//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#if BOOST_OS_WINDOWS
#include <boost/winapi/dll.hpp>
#else
#include <dlfcn.h>
#endif

namespace xll {
namespace detail {

template <class R, class V>
using EXCEL12PROC = int (__stdcall *)(int xlfn, int coper, V **rgpvalue12, R *value12Res);

using LPENHELPERPROC = int (__stdcall *)(int wCode, void *lpv);

template <class R, class V>
BOOST_FORCEINLINE EXCEL12PROC<R, V> MdCallBack12()
{
#if BOOST_OS_WINDOWS
    static auto hmodule = boost::winapi::get_module_handle("");
    static auto pfn = (EXCEL12PROC<R, V>)boost::winapi::get_proc_address(hmodule, "MdCallBack12");
#else
    static auto hmodule = dlopen(nullptr, RTLD_LAZY);
    static auto pfn = (EXCEL12PROC<R, V>)dlsym(hmodule, "MdCallBack12");
#endif
    return pfn;
}

} // namespace detail
} // namespace xll
