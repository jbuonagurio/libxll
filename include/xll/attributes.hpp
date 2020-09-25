//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <boost/mp11/list.hpp>
#include <boost/mp11/set.hpp>

namespace xll {
namespace tag {

struct volatile_ {};
struct cluster_safe {};
struct thread_safe {};
struct macro_sheet_equivalent {};

} // namespace tag

template<class T> struct attribute;
template<> struct attribute<tag::volatile_> {};
template<> struct attribute<tag::cluster_safe> {};
template<> struct attribute<tag::thread_safe> {};
template<> struct attribute<tag::macro_sheet_equivalent> {};

template<class... Tags>
struct attribute_set
{
    static_assert(boost::mp11::mp_is_set<
        boost::mp11::mp_list<attribute<Tags>...>>::value,
        "attributes must be unique");
};

} // namespace xll
