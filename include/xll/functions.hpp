//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <xll/constants.hpp>
#include <xll/callback.hpp>
#include <xll/xloper.hpp>

#include <array>
#include <memory>

namespace xll {

/// Returns the full path and file name of the DLL as xltypeStr.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlgetname
inline xlstr get_name()
{
    static variant result; // Not thread-safe
    int rc = Excel12(xlGetName, &result);
    if (rc != XLRET::xlretSuccess || result.xltype() != xltypeStr)
        return {};
    return result.get<xlstr>(); // xlFree
}

/// Returns the number of bytes remaining on the stack.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlstack
inline int stack_size()
{
    variant result;
    int rc = Excel12(xlStack, &result);
    if (rc != XLRET::xlretSuccess || result.xltype() != xltypeInt)
        return 0;
    return static_cast<int>(result.get<xlint>());
}

/// Returns the window handle of the top-level Microsoft Excel window.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlgethwnd
inline int get_hwnd()
{
    variant result;
    int rc = Excel12(xlGetHwnd, &result);
    if (rc != XLRET::xlretSuccess || result.xltype() != xltypeInt)
        return 0; // Invalid HWND
    return static_cast<int>(result.get<xlint>());
}

/// Converts one type of XLOPER to another, if possible.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlcoerce
template<XLTYPE... Ts>
inline variant coerce(const variant *source)
{
    variant result;
    constexpr int flags = (Ts | ...);
    xloper<xlint> types(flags);
    Excel12(xlCoerce, &result, source, &types);
    return result; // xlFree (xltypeStr, xltypeMulti)
}

/// Returns the ID of a named sheet.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlsheetid
//Excel12(xlSheetId, LPXLOPER12 pxRes, 1, LPXLOPER12 pxSheetName);

/// Returns the name of a worksheet or macro sheet from its internal sheet ID.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlsheetnm
//Excel12(xlSheetNm, LPXLOPER12 pxRes, 1, LPXLOPER12 pxExtref);

/// Allocates persistent storage using xltypeBigData.
/// \sa https://learn.microsoft.com/en-us/office/client-developer/excel/xldefinebinaryname
inline int define_binary_name(const std::wstring& name, const xlbigdata& handle)
{
    xloper<xlstr> arg1(name);
    variant arg2(xlmissing{});
    if (handle.h != nullptr)
        arg2.emplace<xlbigdata>(handle);
    return Excel12(xlDefineBinaryName, nullptr, &arg1, &arg2);
}

/// Returns persistent storage using xltypeBigData.
/// \sa https://learn.microsoft.com/en-us/office/client-developer/excel/xlgetbinaryname
//Excel12(xlGetBinaryName, LPXLOPER12 pxRes, 1, LPXLOPER12 pxName);
inline xlbigdata get_binary_name(const std::wstring& name)
{
    xloper<xlstr> arg1(name);
    variant result;
    int rc = Excel12(xlGetBinaryName, &result, &arg1);
    if (rc != XLRET::xlretSuccess || result.xltype() != xltypeBigData)
        return xlbigdata{};
    return result.get<xlbigdata>();
}

/// Puts constant values into cells or ranges very quickly.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlset
//Excel12(xlSet, LPXLOPER12 pxRes, 2, LPXLOPER12 pxReference, LPXLOPER pxValue);

/// Display a message box using xlcAlert.
/// \param[in] text The text to be displayed.
/// \param[in] type Dialog type: 1 = OK/Cancel, 2 = Info, 3 = Warning
inline int message_box(const std::wstring& text, int type = 2)
{
    xloper<xlint> arg1(type);
    xloper<xlstr> arg2(text);
    return Excel12(xlcAlert, nullptr, &arg1, &arg2);
}

/// Display text in the status bar using xlcMessage.
/// \param[in] text The text to be displayed.
inline int status_bar(const std::wstring& text)
{
    xloper<xlbool> arg1((xlbool)true);
    xloper<xlstr> arg2(text);
    return Excel12(xlcMessage, nullptr, &arg1, &arg2);
}

/// Return the result of an asynchronous UDF.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlasyncreturn
inline bool async_return(xloper<xlbigdata>& handle, variant& value)
{
    variant result;
    int rc = Excel12(xlAsyncReturn, &result, &handle, &value);
    if (rc != XLRET::xlretSuccess || result.xltype() != xltypeBool)
        return false;
    return static_cast<bool>(result.get<xlbool>());
}

/// Registers an event handler function for an asynchronous UDF.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xleventregister
inline bool register_event(const std::wstring& procedure, XLEVENT event)
{
    variant result;
    xloper<xlstr> arg1(procedure);
    xloper<xlint> arg2(event);
    int rc = Excel12(xlEventRegister, &result, &arg1, &arg2);
    if (rc != XLRET::xlretSuccess || result.xltype() != xltypeBool)
        return false;
    return static_cast<bool>(result.get<xlbool>());
}

/// Unloads and deactivates the DLL.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlfunregister-form-2
inline bool unload()
{
    variant result;
    variant name(get_name());
    int rc = Excel12(xlfUnregister, &result, &name);
    if (rc != XLRET::xlretSuccess || result.xltype() != xltypeBool)
        return false;
    return static_cast<bool>(result.get<xlbool>());
}

/// Unregisters a function.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlfunregister-form-1
inline bool unregister(double id)
{
    variant result;
    xloper<xlnum> idvar(id);
    int rc = Excel12(xlfUnregister, &result, &idvar);
    if (rc != XLRET::xlretSuccess || result.xltype() != xltypeBool)
        return false;
    return static_cast<bool>(result.get<xlbool>());
}

/// Checks for a user interrupt (pressed ESC key).
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlabort
inline bool check_interrupt()
{
    variant result;
    int rc = Excel12(xlAbort, &result);
    if (rc != XLRET::xlretSuccess || result.xltype() != xltypeBool)
        return false;
    return static_cast<bool>(result.get<xlbool>());
}

} // namespace xll