//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

/**
 * \file value.hpp
 * XLOPER12 internal structures from XLCALL.H adapted to C++.
 *  
 * Excludes structures `XLREF`, `XLMREF`, `FP`, `FMLAINFO` and `MOUSEINFO`.
 */

#include <xll/config.hpp>

#include <xll/constants.hpp>
#include <xll/detail/variant.hpp>
#include <xll/error.hpp>
#include <xll/pstring.hpp>
#include <xll/value.hpp>

#include <limits>
#include <string>
#include <system_error>
#include <type_traits>

namespace xll {

using variant = detail::variant_base<
    xlnum,
    xlstr,
    xlbool,
    xlerr,
    xlint,
    xlsref,
    xlref,
    xlmulti,
    xlflow,
    xlbigdata,
    xlmissing,
    xlnil>;

template<class T>
using xloper = detail::variant_base<T>;

// Used for async call handle parameters.
using handle = detail::variant_base<xlbigdata>;

} // namespace xll