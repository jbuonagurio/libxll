//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include "config.hpp"
#include "module.hpp"
#include "utilities.hpp"

#include <comdef.h>
#include <comutil.h> // _bstr_t
#include <errhandlingapi.h> // GetLastError
#include <handleapi.h> // CloseHandle
#include <ktmw32.h> // CreateTransaction, CommitTransaction
#include <libloaderapi.h> // GetModuleFileName
#include <oleauto.h> // RegisterTypeLib, RegisterTypeLibForUser
#include <PathCch.h> // PathCchRemoveFileSpec
#include <strsafe.h> // StringCchCopy
#include <winreg.h> // RegCreateKeyTransacted, RegOpenKeyTransacted, RegDeleteKeyTransacted, RegSetValueEx, RegCloseKey

#pragma comment(lib, "Advapi32")
#pragma comment(lib, "KtmW32")
#pragma comment(lib, "Pathcch")

#include <string>
#include <system_error>

//#include "spdlog/spdlog.h"

namespace Registrar {

// See also: atltransactionmanager.h
class RegistryTransaction
{
public:
    RegistryTransaction()
    {
        //spdlog::debug("CreateTransaction");

        m_hTransaction = ::CreateTransaction(nullptr, nullptr, TRANSACTION_DO_NOT_PROMOTE, 0, 0, INFINITE, nullptr);
        if (m_hTransaction == INVALID_HANDLE_VALUE) {
            DWORD ec = ::GetLastError();
            ::CloseHandle(m_hTransaction);
            throw _com_error(HRESULT_FROM_WIN32(ec));
        }
    }

    // If the transaction handle is closed before Commit() is called,
    // KTM will automatically rollback the transaction.
    ~RegistryTransaction()
    {
        if (m_hTransaction)
            ::CloseHandle(m_hTransaction);
    }

    void SetKeyAndValue(HKEY hKey, const std::string& subKey, const std::string& value)
    {
        //spdlog::debug("RegCreateKey [{}]", subKey);

        HKEY hkResult;
        DWORD dwDisposition; // REG_CREATED_NEW_KEY or REG_OPENED_EXISTING_KEY
        HRESULT rc;

        rc = HRESULT_FROM_WIN32(::RegCreateKeyTransactedA(
            hKey, subKey.c_str(), 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_WRITE,
            nullptr, &hkResult, &dwDisposition, m_hTransaction, nullptr));

        if (FAILED(rc))
            throw _com_error(rc);

        if (!value.empty()) {
            const char *lpValue = value.c_str();
            DWORD dataLength = strlen(lpValue) + 1;
            rc = HRESULT_FROM_WIN32(::RegSetValueExA(hkResult, nullptr, 0, REG_SZ, (const BYTE *)lpValue, dataLength));
            ::RegCloseKey(hkResult);
            if (FAILED(rc))
                throw _com_error(rc);
        }
        else {
            ::RegCloseKey(hkResult);
        }
    }

    void SetValue(HKEY hKey, const std::string& subKey, const std::string& valueName, const std::string& value)
    {
        HKEY hkResult;
        HRESULT rc;

        rc = HRESULT_FROM_WIN32(::RegOpenKeyTransactedA(hKey, subKey.c_str(), 0, KEY_ALL_ACCESS, &hkResult, m_hTransaction, nullptr));
        if (FAILED(rc))
            throw _com_error(rc);

        if (!value.empty()) {
            const char *lpValue = value.c_str();
            DWORD dataLength = strlen(lpValue) + 1;
            rc = HRESULT_FROM_WIN32(::RegSetValueExA(hkResult, valueName.c_str(), 0, REG_SZ, (const BYTE *)lpValue, dataLength));
            ::RegCloseKey(hkResult);
            if (FAILED(rc))
                throw _com_error(rc);
        }
        else {
            ::RegCloseKey(hKey);
        }
    }

    void SetValue(HKEY hKey, const std::string& subKey, const std::string& valueName, DWORD value)
    {
        HKEY hkResult;
        HRESULT rc;

        rc = HRESULT_FROM_WIN32(::RegOpenKeyTransactedA(hKey, subKey.c_str(), 0, KEY_ALL_ACCESS, &hkResult, m_hTransaction, nullptr));
        if (FAILED(rc))
            throw _com_error(rc);

        rc = HRESULT_FROM_WIN32(::RegSetValueExA(hkResult, valueName.c_str(), 0, REG_DWORD, (const BYTE *)&value, sizeof(DWORD)));
        ::RegCloseKey(hkResult);
        if (FAILED(rc))
            throw _com_error(rc);
    }

    void DeleteKey(HKEY hKey, const std::string& subKey)
    {
        //spdlog::debug("RegDeleteKey [{}]", subKey);

        HRESULT rc = HRESULT_FROM_WIN32(::RegDeleteKeyTransactedA(hKey, subKey.c_str(), m_samPlatform, 0, m_hTransaction, nullptr));
        if (FAILED(rc))
            throw _com_error(rc);
    }

    void Commit()
    {
        //spdlog::debug("CommitTransaction");

        if (!::CommitTransaction(m_hTransaction)) {
            DWORD ec = ::GetLastError();
            ::CloseHandle(m_hTransaction);
            throw _com_error(HRESULT_FROM_WIN32(ec));
        }
    }

private:
    HANDLE m_hTransaction = nullptr;
#if _WIN64
    REGSAM m_samPlatform = KEY_WOW64_64KEY;
#else
    REGSAM m_samPlatform = KEY_WOW64_32KEY;
#endif
};

//
// Server registration
//

inline HRESULT XLLRegisterTypeLibrary()
{
    HRESULT rc;

    // Load the embedded type library.
    wchar_t szModule[MAX_PATH];
    HMODULE hModule = (HMODULE)XLLRibbonModule::Instance().GetResourceInstance();
    rc = ::GetModuleFileNameW(hModule, szModule, MAX_PATH);
    if (FAILED(rc))
        return rc;

    ITypeLibPtr pTypeLib;
    rc = ::LoadTypeLib(_bstr_t(szModule), &pTypeLib);
    if (FAILED(rc))
        return rc;

    // Register the type library.
    wchar_t szDir[MAX_PATH];
    ::StringCchCopyW(szDir, MAX_PATH, szModule);
    ::PathCchRemoveFileSpec(szDir, MAX_PATH); // pathcch.lib
    
    if (XLLRibbonModule::Instance().GetPerUserRegistration()) {
        //spdlog::debug("RegisterTypeLibForUser");
        rc = ::RegisterTypeLibForUser(pTypeLib, _bstr_t(szModule), _bstr_t(szDir));
    }
    else {
        //spdlog::debug("RegisterTypeLib");
        rc = ::RegisterTypeLib(pTypeLib, _bstr_t(szModule), _bstr_t(szDir));
    }

    return rc;
}

inline HRESULT XLLUnregisterTypeLibrary()
{
    bool perUser = XLLRibbonModule::Instance().GetPerUserRegistration();

    if (perUser) {
        //spdlog::debug("UnRegisterTypeLibForUser");
        return ::UnRegisterTypeLibForUser(LIBID_XLLRibbonAddIn, 1, 0, LANG_NEUTRAL, SYS_WIN32);
    }
    else {
        //spdlog::debug("UnRegisterTypeLib");
        return ::UnRegisterTypeLib(LIBID_XLLRibbonAddIn, 1, 0, LANG_NEUTRAL, SYS_WIN32);
    }
}

// Register the COM server. See also: statreg.h in ATL.
inline HRESULT XLLRegisterServer()
{
    const std::string szDescription = ADDIN_DESCRIPTION;
    const std::string szFriendlyName = ADDIN_FRIENDLY_NAME;
    const std::string szProgID = ADDIN_PROGID;
    const std::string szVIProgID = ADDIN_VI_PROGID;
    const std::string szCLSID = GuidToStdString(CLSID_Connect);
    const std::string szLIBID = GuidToStdString(LIBID_XLLRibbonAddIn);

    bool perUser = XLLRibbonModule::Instance().GetPerUserRegistration();
    HKEY hKeyRoot = perUser ? HKEY_CURRENT_USER : HKEY_LOCAL_MACHINE;

    HMODULE hModule = (HMODULE)XLLRibbonModule::Instance().GetResourceInstance();
    char szModule[MAX_PATH] = { '\0' };
    HRESULT rc = ::GetModuleFileNameA(hModule, szModule, MAX_PATH);
    if (FAILED(rc))
        return rc;

    // [HKEY_CLASSES_ROOT\CLSID\{CLSID}]
    const std::string skCLSID = "Software\\Classes\\CLSID\\" + szCLSID;
    const std::string skCLSIDIPS32 = skCLSID + "\\InprocServer32";
    const std::string skCLSIDProgID = skCLSID + "\\ProgID";
    const std::string skCLSIDVIProgID = skCLSID + "\\VersionIndependentProgID";
    const std::string skCLSIDTypeLib = skCLSID + "\\TypeLib";
    const std::string skCLSIDProgrammable = skCLSID + "\\Programmable";

    // [HKEY_CLASSES_ROOT\{VerIndProgID}]
    const std::string skVIProgID = "Software\\Classes\\" + szVIProgID;
    const std::string skVIProgIDCLSID = skVIProgID + "\\CLSID";
    const std::string skVIProgIDCurVer = skVIProgID + "\\CurVer";

    // [HKEY_CLASSES_ROOT\{ProgID}]
    const std::string skProgID = "Software\\Classes\\" + szProgID;
    const std::string skProgIDCLSID = skProgID + "\\CLSID";

    // [HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\{VerIndProgID}]
    const std::string skAddIn = "Software\\Microsoft\\Office\\Excel\\Addins\\" + szVIProgID;
    
    try {
        RegistryTransaction t;
        t.SetKeyAndValue(hKeyRoot, skCLSID, szFriendlyName);
        t.SetKeyAndValue(hKeyRoot, skCLSIDIPS32, std::string(szModule));
        t.SetValue(hKeyRoot, skCLSIDIPS32, "ThreadingModel", "Apartment");
        t.SetKeyAndValue(hKeyRoot, skCLSIDProgID, szProgID);
        t.SetKeyAndValue(hKeyRoot, skCLSIDVIProgID, szVIProgID);
        t.SetKeyAndValue(hKeyRoot, skCLSIDTypeLib, szLIBID);
        t.SetKeyAndValue(hKeyRoot, skCLSIDProgrammable, {});
        t.SetKeyAndValue(hKeyRoot, skVIProgID, szFriendlyName);
        t.SetKeyAndValue(hKeyRoot, skVIProgIDCLSID, szCLSID);
        t.SetKeyAndValue(hKeyRoot, skVIProgIDCurVer, szProgID);
        t.SetKeyAndValue(hKeyRoot, skProgID, szFriendlyName);
        t.SetKeyAndValue(hKeyRoot, skProgIDCLSID, szCLSID);
        t.SetKeyAndValue(hKeyRoot, skAddIn, {});
        t.SetValue(hKeyRoot, skAddIn, "Description", szDescription);
        t.SetValue(hKeyRoot, skAddIn, "FriendlyName", szFriendlyName);
        t.SetValue(hKeyRoot, skAddIn, "LoadBehavior", ADDIN_LOAD_BEHAVIOR);
        t.Commit();
    }
    catch (const _com_error& e) {
        //spdlog::error(std::system_category().message(e.Error()));
        return e.Error();
    }

    return S_OK;
}

inline HRESULT XLLUnregisterServer()
{
    const std::string szProgID = ADDIN_PROGID;
    const std::string szVIProgID = ADDIN_VI_PROGID;
    const std::string szCLSID = GuidToStdString(CLSID_Connect);

    bool perUser = XLLRibbonModule::Instance().GetPerUserRegistration();
    HKEY hKeyRoot = perUser ? HKEY_CURRENT_USER : HKEY_LOCAL_MACHINE;

    // [HKEY_CLASSES_ROOT\CLSID\{CLSID}]
    const std::string skCLSID = "Software\\Classes\\CLSID\\" + szCLSID;
    const std::string skCLSIDIPS32 = skCLSID + "\\InprocServer32";
    const std::string skCLSIDProgID = skCLSID + "\\ProgID";
    const std::string skCLSIDVIProgID = skCLSID + "\\VersionIndependentProgID";
    const std::string skCLSIDTypeLib = skCLSID + "\\TypeLib";
    const std::string skCLSIDProgrammable = skCLSID + "\\Programmable";

    // [HKEY_CLASSES_ROOT\{VerIndProgID}]
    const std::string skVIProgID = "Software\\Classes\\" + szVIProgID;
    const std::string skVIProgIDCLSID = skVIProgID + "\\CLSID";
    const std::string skVIProgIDCurVer = skVIProgID + "\\CurVer";

    // [HKEY_CLASSES_ROOT\{ProgID}]
    const std::string skProgID = "Software\\Classes\\" + szProgID;
    const std::string skProgIDCLSID = skProgID + "\\CLSID";

    // [HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\{VerIndProgID}]
    const std::string skAddIn = "Software\\Microsoft\\Office\\Excel\\Addins\\" + szVIProgID;

    try {
        RegistryTransaction t;
        t.DeleteKey(hKeyRoot, skCLSIDProgrammable);
        t.DeleteKey(hKeyRoot, skCLSIDTypeLib);
        t.DeleteKey(hKeyRoot, skCLSIDVIProgID);
        t.DeleteKey(hKeyRoot, skCLSIDProgID);
        t.DeleteKey(hKeyRoot, skCLSIDIPS32);
        t.DeleteKey(hKeyRoot, skCLSID);
        t.DeleteKey(hKeyRoot, skVIProgIDCurVer);
        t.DeleteKey(hKeyRoot, skVIProgIDCLSID);
        t.DeleteKey(hKeyRoot, skVIProgID);
        t.DeleteKey(hKeyRoot, skProgIDCLSID);
        t.DeleteKey(hKeyRoot, skProgID);
        t.DeleteKey(hKeyRoot, skAddIn);
        t.Commit();
    }
    catch (const _com_error& e) {
        //spdlog::error(std::system_category().message(e.Error()));
        return e.Error();
    }

    return S_OK;
}

} // namespace Registrar
