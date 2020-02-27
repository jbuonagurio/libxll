//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <cstdlib>

#include <xll/config.hpp>
#include <xll/constants.hpp>
#include <xll/log.hpp>

#include <boost/config.hpp>
#include <boost/current_function.hpp>
#include <boost/winapi/dll.hpp>

#if defined(XLL_DISABLE_ASSERTS) || defined(NDEBUG)

#define XLL_ASSERT(expr) ((void)0)

#else

namespace xll {

inline void assertion_failed(const char *expr, const char *function, const char *file, int line)
{
    xll::log()->critical("Assertion failed: {}, file {}, line {}", expr, file, line);

    // Regular assert would call std::abort and terminate Excel. Instead, call
    // xlfUnregister (Form 2) to force unload and deactivation of the DLL.

    using EXCEL12PROC = int (__stdcall *)(int xlfn, int coper, void **rgpvalue12, void *value12Res);
    auto hmodule = boost::winapi::get_module_handle("");
    auto pfn = (EXCEL12PROC)boost::winapi::get_proc_address(hmodule, "MdCallBack12");
	if (pfn == nullptr)
        std::abort();
    
    void *xDLL = nullptr;
    if ((pfn)(xlGetName, 0, nullptr, xDLL) != 0 &&
        (pfn)(xlfUnregister, 1, &xDLL, nullptr) != 0)
        std::abort();
}

} // namespace xll

#define XLL_ASSERT(expr) (BOOST_LIKELY(!!(expr)) ?((void)0): xll::assertion_failed(#expr, BOOST_CURRENT_FUNCTION, __FILE__, __LINE__))

#endif
