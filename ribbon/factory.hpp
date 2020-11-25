//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include "addin.hpp"
#include "module.hpp"

#include <atomic>

//#include "spdlog/spdlog.h"

class XLLRibbonAddInFactory final : public IClassFactory
{
public:
    XLLRibbonAddInFactory()
    {
        XLLRibbonModule::Instance().Lock();
    }

    virtual ~XLLRibbonAddInFactory()
    {
        XLLRibbonModule::Instance().Unlock();
    }
    
    //
    // IUnknown implementation
    //
    
    ULONG __stdcall AddRef() override
    {
        //spdlog::debug("IClassFactory: AddRef", m_cRef);
        return ++m_cRef;
    }

    ULONG __stdcall Release() override
    {
        //spdlog::debug("IClassFactory: Release", m_cRef);
        if (--m_cRef == 0) {
            delete this;
            return 0;
        }
        return m_cRef;
    }
    
    HRESULT __stdcall QueryInterface(REFIID iid, void **ppv) override
    {
        //spdlog::debug("IClassFactory: QueryInterface");
        if (!ppv)
            return E_INVALIDARG;
        
        *ppv = nullptr;
        
        if (iid == IID_IUnknown || iid == IID_IClassFactory) {
            *ppv = static_cast<IClassFactory *>(this);
            reinterpret_cast<IUnknown *>(*ppv)->AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    //
    // IClassFactory implementation
    //
    
    HRESULT __stdcall CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject) override
    {
        //spdlog::debug("IClassFactory: CreateInstance");
        if (pUnkOuter)
            return CLASS_E_NOAGGREGATION;
        
        Connect *p = new Connect;
        if (!p)
            return E_OUTOFMEMORY;
        
        HRESULT rc = p->QueryInterface(riid, ppvObject);
        p->Release();
        return rc;
    }
    
    HRESULT __stdcall LockServer(BOOL) override
    {
        return S_OK;
	}

private:
    std::atomic_uint m_cRef{ 1 };
};
