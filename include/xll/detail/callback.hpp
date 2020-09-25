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

template<class R = void, class V = void>
using EXCEL12PROC = int (__stdcall *)(int xlfn, int coper, V **rgpvalue12, R *value12Res);

template<class R = void, class V = void>
BOOST_FORCEINLINE EXCEL12PROC<R, V> MdCallBack12()
{
#if BOOST_OS_WINDOWS
    static auto hmodule = boost::winapi::get_module_handle("");
    static auto pfn = (EXCEL12PROC<R, V>)boost::winapi::get_proc_address(hmodule, EXCEL12ENTRYPT);
#else
    static auto hmodule = dlopen(nullptr, RTLD_LAZY);
    static auto pfn = (EXCEL12PROC<R, V>)dlsym(hmodule, EXCEL12ENTRYPT);
#endif
    return pfn;
}

// LPenHelper symbol from Excel 2010 SDK; not present in Excel 2013 SDK.
using LPENHELPERPROC = long (__stdcall *)(int wCode, void *lpv);

} // namespace detail
} // namespace xll
