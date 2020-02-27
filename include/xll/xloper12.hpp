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

using variant12 = detail::variant<
    value<tag::xlnum>,
    value<tag::xlstr>,
    value<tag::xlbool>,
    value<tag::xlerr>,
    value<tag::xlint>,
    value<tag::xlsref>,
    value<tag::xlref>,
    value<tag::xlmulti>,
    value<tag::xlflow>,
    value<tag::xlbigdata>,
    value<tag::xlmissing>,
    value<tag::xlnil>>;

//
// XLOPER12 value types
//

template <class Tag>
using xloper12 = detail::variant<value<Tag>>;

//using string12 = detail::variant<value<tag::xlstr>>;

using handle12 = detail::variant<value<tag::xlbigdata>>;

template <std::size_t N>
using literal12 = detail::variant<basic_pstring_literal<wchar_t, N>>;

} // namespace xll