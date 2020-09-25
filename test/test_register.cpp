//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#include <xll/xll.hpp>

#include <boost/core/lightweight_test.hpp>

#include <array>

using namespace xll;

const char * __stdcall error1(handle *, handle *)
{
    return "";
}

const char * __stdcall error2(handle *)
{
    return "";
}

void __stdcall async1(handle *)
{
    return;
}

const char * __stdcall func1(variant *)
{
    return "";
}

int main()
{
    {
        // expect static assertion failure: macro sheet equivalent functions cannot be thread-safe
        //constexpr auto attrs = attribute_set<tag::thread_safe, tag::macro_sheet_equivalent>();
        //constexpr auto tt = detail::type_text(func1, attrs);
    }
    {
        // expect static assertion failure: macro sheet equivalent functions cannot be cluster-safe
        //constexpr auto attrs = attribute_set<tag::cluster_safe, tag::macro_sheet_equivalent>();
        //constexpr auto tt = detail::type_text(func1, attrs);
    }
    {
        // expect static assertion failure: multiple async handles in argument list
        //constexpr auto tt = detail::type_text(error1, attribute_set<>());
    }
    {
        // expect static assertion failure: async functions must have void return type
        //constexpr auto tt = detail::type_text(error2, attribute_set<>());
    }
    {
        // expect static assertion failure: async functions cannot be cluster-safe
        //constexpr auto attrs = attribute_set<tag::cluster_safe>();
        //constexpr auto tt = detail::type_text(async1, attrs);
    }
    {
        constexpr auto attrs = attribute_set<tag::thread_safe>();
        constexpr auto tt = detail::type_text(async1, attrs);
        constexpr std::array<wchar_t, 3> expected{{ L'>', L'X', L'$' }};
        BOOST_TEST(tt == expected);
    }
    {
        constexpr auto attrs = attribute_set<tag::volatile_>();
        constexpr auto tt = detail::type_text(func1, attrs);
        constexpr std::array<wchar_t, 3> expected{{ L'C', L'Q', L'!' }};
        BOOST_TEST(tt == expected);
    }
    {
        constexpr auto attrs = attribute_set<tag::cluster_safe>();
        constexpr auto tt = detail::type_text(func1, attrs);
        constexpr std::array<wchar_t, 3> expected{{ L'C', L'Q', L'&' }};
        BOOST_TEST(tt == expected);
    }

    return boost::report_errors();
}