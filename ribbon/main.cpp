//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#ifndef NOMINMAX
#define NOMINMAX
#endif

#ifndef STRICT
#define STRICT
#endif

#ifndef NTDDI_VERSION
#define NTDDI_VERSION 0x06020000 // Windows 8+
#endif

#include "addin.hpp"
#include "factory.hpp"
#include "module.hpp"
#include "registrar.hpp"

#include <Windows.h>

#include <memory>
#include <string>
#include <system_error>

//#include "spdlog/spdlog.h"
//#include "spdlog/sinks/msvc_sink.h"

//void InitializeLogger()
//{
//    auto sink = std::make_shared<spdlog::sinks::msvc_sink_mt>();
//    auto logger = std::make_shared<spdlog::logger>("ribbon_debug", sink);
//    logger->set_level(spdlog::level::debug);
//    spdlog::set_default_logger(logger);
//}

extern "C" BOOL __stdcall DllMain(HINSTANCE hinstDLL, DWORD dwReason, LPVOID)
{
    switch (dwReason)
    {
    case DLL_PROCESS_ATTACH:
        //InitializeLogger();
        //spdlog::debug("DLL_PROCESS_ATTACH");
        ::DisableThreadLibraryCalls((HMODULE)hinstDLL);
        XLLRibbonModule::Instance().SetResourceInstance(hinstDLL);
        break;
    case DLL_THREAD_ATTACH:
        break;
    case DLL_THREAD_DETACH:
        break;
    case DLL_PROCESS_DETACH:
        //spdlog::debug("DLL_PROCESS_DETACH");
        break;
    }

	return TRUE;
}

extern "C" HRESULT __stdcall DllCanUnloadNow()
{
    //spdlog::debug("DllCanUnloadNow");
    
    if (XLLRibbonModule::Instance().GetLockCount() != 0)
        return S_FALSE;
    return S_OK;
}

extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
    //spdlog::debug("DllGetClassObject");
    
    if (rclsid != CLSID_Connect)
		return CLASS_E_CLASSNOTAVAILABLE;
    
    XLLRibbonAddInFactory *factory = new XLLRibbonAddInFactory;
    if (!factory)
        return E_OUTOFMEMORY;

    HRESULT rc = factory->QueryInterface(riid, ppv);
    factory->Release();
    return rc;
}

extern "C" HRESULT __stdcall DllRegisterServer()
{
    //spdlog::debug("DllRegisterServer");
    
    HRESULT rc;
    rc = Registrar::XLLRegisterServer();
    if (FAILED(rc))
        return SELFREG_E_CLASS;

    rc = Registrar::XLLRegisterTypeLibrary();
    if (FAILED(rc))
        return SELFREG_E_TYPELIB;

    return S_OK;
}

extern "C" HRESULT __stdcall DllUnregisterServer()
{
    //spdlog::debug("DllUnregisterServer");
    
    HRESULT rc;
    rc = Registrar::XLLUnregisterServer();
    if (FAILED(rc))
        return SELFREG_E_CLASS;

    rc = Registrar::XLLUnregisterTypeLibrary();
    if (FAILED(rc))
        return SELFREG_E_TYPELIB;

    return S_OK;
}

// Called by regsvr32 with /i switch. The "user" option enables per-user registration,
// redirecting registry values from HKEY_CLASSES_ROOT to HKEY_CURRENT_USER. This is
// consistent with ATL.

extern "C" HRESULT __stdcall DllInstall(BOOL bInstall, LPCWSTR pszCmdLine)
{
    //spdlog::debug("DllInstall");

    if (pszCmdLine != nullptr) {
        static const wchar_t szUserSwitch[] = L"user";
        bool perUser = (_wcsnicmp(pszCmdLine, szUserSwitch, sizeof(szUserSwitch) / sizeof(szUserSwitch[0])) == 0);
        XLLRibbonModule::Instance().SetPerUserRegistration(perUser);
    }
    
    if (bInstall) {
        HRESULT rc = DllRegisterServer();
        if (FAILED(rc))
            DllUnregisterServer();
        return rc;
    }
    else {
        return DllUnregisterServer();
    }
}
