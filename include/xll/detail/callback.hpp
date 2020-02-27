//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <boost/winapi/dll.hpp>

namespace xll {
namespace detail {

// Excel 12 entry point definition from XLCALL.CPP.
template <class R, class V>
using EXCEL12PROC = int (__stdcall *)(int xlfn, int coper, V **rgpvalue12, R *value12Res);

template <class R, class V>
BOOST_FORCEINLINE EXCEL12PROC<R, V> MdCallBack12()
{
    static auto hmodule = boost::winapi::get_module_handle("");
    static auto pfn = (EXCEL12PROC<R, V>)boost::winapi::get_proc_address(hmodule, "MdCallBack12");
    return pfn;
}

} // namespace detail
} // namespace xll