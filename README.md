<!--
  Copyright (c) 2020 John Buonagurio (jbuonagurio at exponent dot com)
  Distributed under the Boost Software License, Version 1.0.
  (See accompanying file LICENSE_1_0.txt or copy at
  https://www.boost.org/LICENSE_1_0.txt)
-->

## libxll - header-only C++ framework for Excel add-ins

**libxll is an early developer preview, and is not yet suitable for general usage. Features and implementation are subject to change.**

libxll provides high-level C++17 utilities to make developing native Excel add-ins easier. Licensed under the [Boost Software License](http://www.boost.org/LICENSE_1_0.txt).

Compared to similar projects for Excel development in modern C++, such as [fancidev/xll](https://github.com/fancidev/xll) and [e4lam/xlkit](https://github.com/e4lam/xlkit), libxll differs in the following ways:

* Intrusive, header-only interface which requires user code to implement only a few functions with minimal boilerplate. By requiring the user to define a suitable API for DLL symbol exports, libxll does not impose any constraints on the build environment or require compiler-specific extensions.
* Designed using cross-platform Boost libraries and STL, enabling future compatibility with Excel for Mac.
* Provides `xll::variant`: a custom type-safe tagged union using XLOPER12-compatible aligned storage and sensible overload resolution.
* Uses template metaprogramming to validate function arguments and generate type text for [`xlfRegister`](https://docs.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1#data-types) at compile time.
* Provides custom Pascal string and string literal classes compatible with XLOPER12, with fast, automatic conversion between UTF-8 and UTF-16.
* Provides interface to Boost uBLAS for FP12 arrays.
* Optional `OutputDebugStringA` logging using [spdlog](https://github.com/gabime/spdlog). Useful with [Sysinternals DebugView](https://docs.microsoft.com/en-us/sysinternals/downloads/debugview).

### Example

The following is a minimal example of functions to be implemented by the user to register a worksheet function. Arguments and return value are parsed and validated at compile-time.

```cpp
#include <xll/xll.hpp>

#define XLL_EXPORT extern "C" __declspec(dllexport)

/// This function is called by the Add-in Manager to find the long name of the
/// add-in. If xAction = 1, this function should return a string containing the
/// long name of this XLL. If xAction = 2 or 3, this function should return
/// xlerrValue (#VALUE!).
XLL_EXPORT xll::variant * __stdcall xlAddInManagerInfo12(xll::variant *xAction)
{
    using namespace xll;
	
    thread_local variant xInfo, xIntAction;
    xloper<xlint> xDestType(xltypeInt);

	Excel12(xlCoerce, &xIntAction, xAction, &xDestType);
    
    if (xIntAction.xltype() == xltypeInt && xIntAction.get<xlint>() == 1)
        xInfo = L"Sample XLL";
    else
        xInfo = error::excel_error::xlerrValue;
    
	return &xInfo;
}

/// Excel calls xlAutoOpen when the XLL is loaded. Register all functions and
/// perform any additional initialization in this function.
/// \return 1 on success, 0 on failure.
XLL_EXPORT int __stdcall xlAutoOpen()
{
    auto opts = xll::function_options();
    opts.argument_text = L"arg1,arg2";
    opts.category = L"Sample";
    opts.function_help = L"Sample function to do something with two arguments.";
    opts.argument_help = { L"first argument.", L"second argument." };
    
    xll::register_function(testFunction, L"testFunction", L"TEST.FUNCTION", opts);

    return 1;
}

/// Excel calls xlAutoClose when it unloads the XLL.
XLL_EXPORT int __stdcall xlAutoClose()
{
    return 1;
}

/// Called when the XLL is added to the list of active add-ins. The Add-in
/// Manager subsequently calls xlAutoOpen.
XLL_EXPORT int __stdcall xlAutoAdd()
{
    return 1;
}

/// Called when the XLL is removed from the list of active add-ins. The Add-in
/// Manager subsequently calls xlAutoRemove() and xlfUnregister.
XLL_EXPORT int __stdcall xlAutoRemove()
{
    return 1;
}

/// Free internally allocated arrays and call destructor.
XLL_EXPORT int __stdcall xlAutoFree12(xll::variant *pxFree)
{
    pxFree->release();
    return 1;
}

```
