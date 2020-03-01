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
#include <xll/xloper12.hpp>

#include <array>
#include <memory>

namespace xll {

/// Returns the full path and file name of the DLL as xltypeStr.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlgetname
inline std::unique_ptr<variant12> get_name()
{
    auto result = std::make_unique<variant12>();
    Excel12(xlGetName, result.get());
    return std::move(result);
}

/// Returns the number of bytes remaining on the stack.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlstack
inline int stack_size()
{
    variant12 result;
    XLRET rc = Excel12(xlStack, &result);
    if (rc != XLRET::xlretSuccess && result.xltype() != xltypeInt)
        return 0;
    return static_cast<int>(result.get<xlint>());
}

/// Converts one type of XLOPER to another, if possible.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlcoerce
template <XLTYPE... Ts>
inline std::unique_ptr<variant12> coerce(variant12& source)
{
    constexpr int flags = (Ts | ...);
    auto result = std::make_unique<variant12>();
    auto result_types = std::make_unique<xloper12<xlint>>(flags);
    Excel12(xlCoerce, result.get(), &source, result_types.get());
    return std::move(result);
}

/// Display a message box using xlcAlert.
/// \param[in] text The text to be displayed.
/// \param[in] type Dialog type: 1 = OK/Cancel, 2 = Info, 3 = Warning
inline XLRET message_box(const std::wstring& text, int type = 2)
{
    auto arg1 = std::make_unique<xloper12<xlint>>(type);
    auto arg2 = std::make_unique<xloper12<xlstr>>(text);
    return Excel12(xlcAlert, nullptr, arg1.get(), arg2.get());
}

/// Display text in the status bar using xlcMessage.
/// \param[in] text The text to be displayed.
inline XLRET status_bar(const std::wstring& text)
{
    auto arg1 = std::make_unique<xloper12<xlbool>>(true);
    auto arg2 = std::make_unique<xloper12<xlstr>>(text);
    return Excel12(xlcMessage, nullptr, arg1.get(), arg2.get());
}

/// Return the result of an asynchronous UDF.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlasyncreturn
inline bool async_return(const xloper12<xlbigdata>& handle, const variant12& value)
{
    auto arg1 = std::make_unique<xloper12<xlbigdata>>(handle);
    auto arg2 = std::make_unique<variant12>(value);
    variant12 result;
    XLRET rc = Excel12(xlAsyncReturn, &result, arg1.get(), arg2.get());
    if (rc != XLRET::xlretSuccess && result.xltype() != xltypeBool)
        return false;
    return static_cast<bool>(result.get<xlbool>());
}

/// Registers an event handler function for an asynchronous UDF.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xleventregister
inline bool register_event(const std::wstring& procedure, XLEVENT event)
{
    auto arg1 = std::make_unique<xloper12<xlstr>>(procedure);
    auto arg2 = std::make_unique<xloper12<xlint>>(event);
    variant12 result;
    XLRET rc = Excel12(xlEventRegister, &result, arg1.get(), arg2.get());
    if (rc != XLRET::xlretSuccess && result.xltype() != xltypeBool)
        return false;
}

/// Unloads and deactivates the DLL.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlfunregister-form-2
inline bool unload()
{
    std::unique_ptr<variant12> name = get_name();
    if (!name)
        return false;
    
    variant12 result;
    XLRET rc = Excel12(xlfUnregister, &result, name.get());
    if (rc != XLRET::xlretSuccess && result.xltype() != xltypeBool)
        return false;
    return static_cast<bool>(result.get<xlbool>());
}

/// Unregisters a function.
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlfunregister-form-1
inline bool unregister(double id)
{
    auto idvar = std::make_unique<xloper12<xlnum>>(id);
    variant12 result;
    XLRET rc = Excel12(xlfUnregister, &result, idvar.get());
    if (rc != XLRET::xlretSuccess && result.xltype() != xltypeBool)
        return false;
    return static_cast<bool>(result.get<xlbool>());
}

/// Checks for a user interrupt (pressed ESC key).
/// \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlabort
inline bool check_interrupt()
{
    variant12 result;
    XLRET rc = Excel12(xlAbort, &result);
    if (rc != XLRET::xlretSuccess && result.xltype() != xltypeBool)
        return false;
    return static_cast<bool>(result.get<xlbool>());
}

} // namespace xll