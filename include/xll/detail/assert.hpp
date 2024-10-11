//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <xll/constants.hpp>
#include <xll/detail/callback.hpp>
#include <xll/log.hpp>

#include <cstdlib> // std::abort

#include <boost/current_function.hpp>

#ifdef XLL_DISABLE_ASSERTS

#define XLL_ASSERT(expr) ((void)0)

#else

namespace xll {

inline void assertion_failed(const char *expr, [[maybe_unused]] const char *function, const char *file, int line)
{
    xll::log()->critical("Assertion failed: {}, file {}, line {}", expr, file, line);

    // Regular assert would call std::abort and terminate Excel. Instead, call
    // xlfUnregister (Form 2) to force unload and deactivation of the DLL.

    auto pfn = detail::MdCallBack12();
    if (pfn == nullptr)
        std::abort();
    
    void *xDLL = nullptr;
    if ((pfn)(xlGetName, 0, nullptr, xDLL) != 0 &&
        (pfn)(xlfUnregister, 1, &xDLL, nullptr) != 0)
        std::abort();
}

} // namespace xll

#define XLL_ASSERT(expr) (BOOST_LIKELY(!!(expr)) ?((void)0): xll::assertion_failed(#expr, BOOST_CURRENT_FUNCTION, __FILE__, __LINE__))

#endif // XLL_DISABLE_ASSERTS
