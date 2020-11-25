//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include "autogen/addin_h.h"
#include "interfaces.hpp"
#include "module.hpp"
#include "resource.h"
#include "utilities.hpp"

#include <libloaderapi.h> // GetModuleFileName, LoadResource, LockResource, SizeofResource
#include <stringapiset.h> // MultiByteToWideChar
#include <oleauto.h> // DispInvoke, DispGetIDsOfNames, SysAllocStringLen
#include <olectl.h> // OleCreatePictureIndirect
#include <winbase.h> // FindResource
#include <winuser.h> // LoadImage

#include <atomic>
#include <string>

//#include "spdlog/spdlog.h"

class Connect final :
    public _IDTExtensibility2,
    public IRibbonExtensibility,
    public IRibbonCallback
{
public:
    Connect()
    {
        XLLRibbonModule::Instance().Lock();
    }
    
    virtual ~Connect()
    {
        XLLRibbonModule::Instance().Unlock();
    }
    
    //
    // IUnknown implementation
    //
    
    ULONG __stdcall AddRef() override
    {
        //spdlog::debug("Connect: AddRef", m_cRef);
        return ++m_cRef;
    }
    
    ULONG __stdcall Release() override
    {
        //spdlog::debug("Connect: Release", m_cRef);
        if (--m_cRef == 0) {
            delete this;
            return 0;
        }
        return m_cRef;
    }
    
    HRESULT __stdcall QueryInterface(REFIID iid, void **ppv) override
    {
        //spdlog::debug("Connect: QueryInterface {}", GetIIDString(iid));

        if (!ppv)
            return E_INVALIDARG;

        *ppv = nullptr;
        
        if (iid == __uuidof(IUnknown))
            *ppv = static_cast<IUnknown *>((void *)this);
        else if (iid == __uuidof(IDispatch))
            *ppv = static_cast<IDispatch *>((void *)this);
        else if (iid == __uuidof(_IDTExtensibility2))
            *ppv = static_cast<_IDTExtensibility2 *>(this);
        else if (iid == __uuidof(IRibbonExtensibility))
            *ppv = static_cast<IRibbonExtensibility *>(this);
        else if (iid == __uuidof(IRibbonCallback))
            *ppv = static_cast<IRibbonCallback *>(this);
        else
            return E_NOINTERFACE;

        reinterpret_cast<IUnknown *>(*ppv)->AddRef();
        return S_OK;
    }

    //
    // IDispatch implementation
    //

    HRESULT __stdcall GetIDsOfNames(REFIID, LPOLESTR *rgszNames, UINT cNames, LCID, DISPID *rgdispid) override
    {
        //spdlog::debug("IDispatch: GetIDsOfNames");
        ITypeInfo *pTypeInfo = TypeInfoInstance();
        return ::DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgdispid);
    }

    HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID, ITypeInfo **ppTInfo) override
    {
        //spdlog::debug("IDispatch: GetTypeInfo");
        if (iTInfo != 0)
            return DISP_E_BADINDEX;
        return TypeLibInstance()->GetTypeInfo(iTInfo, ppTInfo);
    }

    HRESULT __stdcall GetTypeInfoCount(UINT *pctinfo) override
    {
        //spdlog::debug("IDispatch: GetTypeInfoCount");
        if (!pctinfo)
            return E_INVALIDARG;
        *pctinfo = 1;
        return S_OK;
    }

    HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID, LCID, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) override
    {
        //spdlog::debug("IDispatch: Invoke DISPID {}", dispIdMember);
        ITypeInfo *pTypeInfo = TypeInfoInstance();
        if (!pTypeInfo)
            return E_INVALIDARG;
        
        IRibbonCallback *pCallback = static_cast<IRibbonCallback *>(this);
        return ::DispInvoke(pCallback, pTypeInfo, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
    }
    
    //
    // _IDTExtensibility2 implementation
    //
    
    HRESULT __stdcall OnConnection(IDispatch *Application, ext_ConnectMode, IDispatch *AddInInst, SAFEARRAY **) override
    {
        //spdlog::debug("_IDTExtensibility2: OnConnection");
        Application->QueryInterface(__uuidof(IDispatch), (LPVOID *)&m_pApplication);
        AddInInst->QueryInterface(__uuidof(IDispatch), (LPVOID *)&m_pAddInInst);
        //CComQIPtr<Excel::_Application> app(m_pApplication);
        //using ApplicationPtr = _com_ptr_t<_com_IIID<Excel::_Application, &__uuidof(Excel::_Application)>>;
        return S_OK;
    }
    
    HRESULT __stdcall OnDisconnection(ext_DisconnectMode, SAFEARRAY **) override
    {
        //spdlog::debug("_IDTExtensibility2: OnDisconnection");
        m_pApplication = static_cast<IDispatch *>(nullptr);
        m_pAddInInst = static_cast<IDispatch *>(nullptr);
        return S_OK;
    }
    
    HRESULT __stdcall OnAddInsUpdate(SAFEARRAY **) override
    {
        //spdlog::debug("_IDTExtensibility2: OnAddInsUpdate");
        return S_OK;
    }
    
    HRESULT __stdcall OnStartupComplete(SAFEARRAY **) override
    {
        //spdlog::debug("_IDTExtensibility2: OnStartupComplete");
        return S_OK;
    }
    
    HRESULT __stdcall OnBeginShutdown(SAFEARRAY **) override
    {
        //spdlog::debug("_IDTExtensibility2: OnBeginShutdown");
        return S_OK;
    }
    
    //
    // IRibbonExtensibility implementation
    //
    
    HRESULT __stdcall GetCustomUI(BSTR RibbonID, BSTR *RibbonXml) override
    {
        UNREFERENCED_PARAMETER(RibbonID); // Microsoft.Excel.Workbook

        //spdlog::debug("IRibbonExtensibility: GetCustomUI");

        if (!RibbonXml)
            return E_POINTER;

        HMODULE hModule = (HMODULE)XLLRibbonModule::Instance().GetResourceInstance();
        HRSRC hResource = ::FindResource(hModule, MAKEINTRESOURCE(IDR_RIBBONUI), RT_RCDATA);
        if (!hResource)
            return E_FAIL;

        DWORD dwSize = ::SizeofResource(hModule, hResource);
        HGLOBAL hResData = ::LoadResource(hModule, hResource);
        if (!hResData)
            return E_FAIL;

        char *pResData = reinterpret_cast<char *>(::LockResource(hResData));
        if (!pResData)
            return E_FAIL;
        
        int cchWideChar = ::MultiByteToWideChar(CP_UTF8, 0, pResData, dwSize, nullptr, 0);
        if (cchWideChar == 0)
            return E_FAIL;

        // COM client is responsible for freeing RibbonXml.
        *RibbonXml = ::SysAllocStringLen(nullptr, cchWideChar);
        if (*RibbonXml == nullptr)
            return E_OUTOFMEMORY;

        ::MultiByteToWideChar(CP_UTF8, 0, pResData, dwSize, *RibbonXml, cchWideChar);
		return S_OK;
    }

    //
    // IRibbonCallback implementation
    //
    
    HRESULT __stdcall ButtonClicked(IDispatch *pControl) override
    {
        //spdlog::debug("IRibbonCallback: ButtonClicked");
        return S_OK;
    }
    
    HRESULT __stdcall LoadImageCallback(BSTR *pbstrImageId, IPictureDisp **ppdispImage) override
    {
        //spdlog::debug("IRibbonCallback: LoadImageCallback");
       
        if (!pbstrImageId)
            return E_POINTER;

        std::wstring resourceName = MakeImageResourceName(*pbstrImageId);
        HINSTANCE hInstance = XLLRibbonModule::Instance().GetResourceInstance();
        HANDLE hImage = ::LoadImageW(hInstance, resourceName.c_str(), IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR);
        if (!hImage)
            return E_FAIL;

        PICTDESC pd;
        memset(&pd, 0, sizeof(PICTDESC));
        pd.bmp.hbitmap = (HBITMAP)hImage;
        pd.picType = PICTYPE_BITMAP;
        pd.cbSizeofstruct = sizeof(PICTDESC);

        HRESULT rc = ::OleCreatePictureIndirect(&pd, IID_IPictureDisp, TRUE, reinterpret_cast<void **>(ppdispImage));
        if (FAILED(rc))
            return E_FAIL;

        return S_OK;
    }

private:
    static std::wstring MakeImageResourceName(const std::wstring& imageName)
    {
        // See "Specifying Ribbon Image Resources" in Ribbbon Framework documentation.
        //
        // +---------+--------------+--------------+
        // | DPI     | Small Image  | Large Image  |
        // +---------+--------------+--------------+
        // | 96 dpi  | 16x16 pixels | 32x32 pixels |
        // | 120 dpi | 20x20 pixels | 40x40 pixels |
        // | 144 dpi | 24x24 pixels | 48x48 pixels |
        // | 192 dpi | 32x32 pixels | 64x64 pixels |
        // +---------+--------------+--------------+
        
        HWND hWnd = ::GetActiveWindow();
        int dpi = hWnd ? (int)GetDpiForWindow(hWnd) : 96;
        std::wstring prefix = L"BITMAP_";
        if (dpi >= 192)
            return prefix + imageName + L"_64x";
        else if (dpi >= 144)
            return prefix + imageName + L"_48x";
        else if (dpi >= 120)
            return prefix + imageName + L"_40x";
        else
            return prefix + imageName + L"_32x";
    }

    static ITypeLib * TypeLibInstance()
    {
        static ITypeLib *pTypeLib;
        if (pTypeLib == nullptr) {
            wchar_t szModule[MAX_PATH];
            HMODULE hModule = (HMODULE)XLLRibbonModule::Instance().GetResourceInstance();
            ::GetModuleFileNameW(hModule, szModule, MAX_PATH);
            ::LoadTypeLib(_bstr_t(szModule), &pTypeLib);
        }
        return pTypeLib;
    }

    static ITypeInfo * TypeInfoInstance()
    {
        static ITypeInfo *pTypeInfo;
        if (pTypeInfo == nullptr) {
            ITypeLib *pTypeLib = TypeLibInstance();
            if (pTypeLib)
                pTypeLib->GetTypeInfoOfGuid(IID_IRibbonCallback, &pTypeInfo);
        }
        return pTypeInfo;
    }

    std::atomic_uint m_cRef{ 1 };
	IDispatchPtr m_pApplication;
	IDispatchPtr m_pAddInInst;
};
