//  Copyright 2023 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#include <xll/xll.hpp>

#include <string>

#include <fmt/format.h>
#include <fmt/printf.h>
#include <fmt/ranges.h>
#include <fmt/std.h>

#if BOOST_OS_WINDOWS
#include <boost/winapi/dll.hpp>
#include <boost/winapi/get_last_error.hpp>
#else
#include <dlfcn.h>
#endif

std::string valueString(xll::variant *value)
{
    using namespace xll;

    if (!value)
        return "NULL";

    switch (value->xltype()) {
    case xltypeNum:
        return fmt::format("{}", static_cast<double>(value->get<xlnum>()));
    case xltypeStr:
        return fmt::format("'{}'", static_cast<std::string>(value->get<xlstr>()));
    case xltypeBool:
        return fmt::format("{}", static_cast<bool>(value->get<xlbool>()));
    case xltypeRef:
        return "Ref(...)";
    case xltypeErr:
        return fmt::format("{}", static_cast<std::error_code>(value->get<xlerr>()));
    case xltypeFlow:
        return "Flow(...)";
    case xltypeMulti:
        return "Multi(...)";
    case xltypeMissing:
        return "Missing()";
    case xltypeNil:
        return "Nil()";
    case xltypeSRef:
        return "SRef(...)";
    case xltypeInt:
        return fmt::format("{}", static_cast<int>(value->get<xlint>()));
    case xltypeBigData:
        return fmt::format("BigData(...)");
    default:
        return {};
    }
}

std::string xlfnString(int xlfn)
{
    switch (xlfn) {
    case xll::xlFree:              return "xlFree";
    case xll::xlStack:             return "xlStack";
    case xll::xlCoerce:            return "xlCoerce";
    case xll::xlSet:               return "xlSet";
    case xll::xlSheetId:           return "xlSheetId";
    case xll::xlSheetNm:           return "xlSheetNm";
    case xll::xlAbort:             return "xlAbort";
    case xll::xlGetInst:           return "xlGetInst";
    case xll::xlGetHwnd:           return "xlGetHwnd";
    case xll::xlGetName:           return "xlGetName";
    case xll::xlEnableXLMsgs:      return "xlEnableXLMsgs";
    case xll::xlDisableXLMsgs:     return "xlDisableXLMsgs";
    case xll::xlDefineBinaryName:  return "xlDefineBinaryName";
    case xll::xlGetBinaryName:     return "xlGetBinaryName";
    case xll::xlGetFmlaInfo:       return "xlGetFmlaInfo";
    case xll::xlGetMouseInfo:      return "xlGetMouseInfo";
    case xll::xlAsyncReturn:       return "xlAsyncReturn";
    case xll::xlEventRegister:     return "xlEventRegister";
    case xll::xlRunningOnCluster:  return "xlRunningOnCluster";
    case xll::xlGetInstPtr:        return "xlGetInstPtr";
    case xll::xlfRegister:         return "xlfRegister";
    default:                       return fmt::format("{}", xlfn);
    }
}

int Register(xll::variant *pxRes)
{
    using namespace xll;
    static int iFunc = 0;
    if (pxRes)
        *pxRes = xlnum((double)++iFunc);
    return xlretSuccess;
}

int GetInstPtr(xll::variant *pxRes)
{
    using namespace xll;
    if (pxRes == nullptr)
        return xlretFailed;
    xll::xlbigdata handle;
#if BOOST_OS_WINDOWS
    handle.h = boost::winapi::get_module_handle("");
#else
    handle.h = dlopen(nullptr, RTLD_LAZY);
#endif
    pxRes->emplace<xlbigdata>(handle);
    return xll::xlretSuccess;
}

int GetHwnd(xll::variant *pxRes)
{
    using namespace xll;
    if (pxRes == nullptr)
        return xlretFailed;
    pxRes->emplace<xlint>(0); // HWND_DESKTOP
    return xll::xlretSuccess;
}

int EventRegister(xll::variant *pxRes)
{
    using namespace xll;
    if (pxRes == nullptr)
        return xlretFailed;
    pxRes->emplace<xlbool>(true);
    return xll::xlretSuccess;
}

int GetName(xll::variant *pxRes)
{
    using namespace xll;
    if (pxRes == nullptr)
        return xlretFailed;
    const wchar_t *value = L"minimal.xll";
    pxRes->emplace<xlstr>(value);
    return xlretSuccess;
}

int GetWorkspace(xll::variant *pxRes)
{
    using namespace xll;
    if (pxRes == nullptr)
        return xlretFailed;
#if BOOST_OS_WINDOWS
    const wchar_t *value = L"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE";
#else
    const wchar_t *value = L"/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel";
#endif
    pxRes->emplace<xlstr>(value);
    return xlretSuccess;
}

XLL_EXPORT int __stdcall MdCallBack12(int xlfn, int coper, xll::variant **rgpxloper12, xll::variant *xloper12Res)
{
    using namespace xll;

    fmt::print("  MdCallBack12\n");
    fmt::print("    xlfn: {}\n", xlfnString(xlfn));
    for (size_t i = 0; i < coper; ++i)
        fmt::print("    rgpxloper12[{}]: {}\n", i, valueString(rgpxloper12[i]));
    fmt::print("    xloper12Res: {}\n", valueString(xloper12Res));
    
    if (xlfn == xlfRegister) {
        return Register(xloper12Res);
    }
    else if (xlfn == xlGetInstPtr) {
        return GetInstPtr(xloper12Res);
    }
    else if (xlfn == xlGetHwnd) {
        return GetHwnd(xloper12Res);
    }
    else if (xlfn == xlEventRegister) {
        return EventRegister(xloper12Res);
    }
    else if (xlfn == xlGetName) {
        return GetName(xloper12Res);
    }
    else if (xlfn == xlfGetWorkspace) {
        return GetWorkspace(xloper12Res);
    }
    else if (xlfn == xlFree) {
        rgpxloper12[0] = nullptr;
    }

    return xlretSuccess;
}

int main()
{
    using namespace xll;

#if BOOST_OS_WINDOWS
    auto hmodule = boost::winapi::load_library("minimal.xll");
    if (hmodule == nullptr) {
        fmt::print("ERROR: LoadLibrary failed: {}\n", boost::winapi::GetLastError());
        return -1;
    }
    auto pxlAddInManagerInfo12 = (variant *(*)(variant *))boost::winapi::get_proc_address(hmodule, "xlAddInManagerInfo12");
    auto pxlAutoOpen = (int(*)())boost::winapi::get_proc_address(hmodule, "xlAutoOpen");
#else
    auto hmodule = dlopen("minimal.xll", RTLD_LAZY);
    if (hmodule == nullptr) {
        fmt::print("ERROR: dlopen failed: {}\n", dlerror());
        return -1;
    }
    auto pxlAddInManagerInfo12 = (variant *(*)(variant *))dlsym(hmodule, "xlAddInManagerInfo12");
    auto pxlAutoOpen = (int(*)())dlsym(hmodule, "xlAutoOpen");
#endif

    if (pxlAddInManagerInfo12 != nullptr) {
        fmt::print("xlAddInManagerInfo12\n");
        static variant xAction = xlnum(1.0);
        variant *result = pxlAddInManagerInfo12(&xAction);
        fmt::print("xlAddInManagerInfo12 -> {}\n\n", valueString(result));
    }

    if (pxlAutoOpen != nullptr) {
        fmt::print("xlAutoOpen\n");
        int result = pxlAutoOpen();
        fmt::print("xlAutoOpen -> {}\n\n", result);
    }

    return 0;
}
