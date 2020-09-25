//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#include <xll/xll.hpp>

#include <boost/core/lightweight_test.hpp>
#include <boost/core/lightweight_test_trait.hpp>

#include <cstdint>
#include <type_traits>
#include <vector>

using namespace xll;

int main()
{
    BOOST_TEST_TRAIT_TRUE((std::is_standard_layout<variant>));

    BOOST_TEST_TRAIT_TRUE((std::is_trivially_constructible<xlnum>));
    BOOST_TEST_TRAIT_FALSE((std::is_trivially_constructible<xlstr>));
    BOOST_TEST_TRAIT_TRUE((std::is_trivially_constructible<xlbool>));
    BOOST_TEST_TRAIT_TRUE((std::is_trivially_constructible<xlerr>));
    BOOST_TEST_TRAIT_TRUE((std::is_trivially_constructible<xlint>));
    BOOST_TEST_TRAIT_FALSE((std::is_trivially_constructible<xlsref>));
    BOOST_TEST_TRAIT_FALSE((std::is_trivially_constructible<xlref>));
    BOOST_TEST_TRAIT_FALSE((std::is_trivially_constructible<xlmulti>));
    BOOST_TEST_TRAIT_TRUE((std::is_trivially_constructible<xlflow>));
    BOOST_TEST_TRAIT_FALSE((std::is_trivially_constructible<xlbigdata>));
    BOOST_TEST_TRAIT_TRUE((std::is_trivially_constructible<xlmissing>));
    BOOST_TEST_TRAIT_TRUE((std::is_trivially_constructible<xlnil>));

#if INTPTR_MAX == INT32_MAX
    BOOST_TEST_EQ(sizeof(xlnum), 8);
    BOOST_TEST_EQ(sizeof(xlstr), 4);
    BOOST_TEST_EQ(sizeof(xlbool), 4);
    BOOST_TEST_EQ(sizeof(xlerr), 4);
    BOOST_TEST_EQ(sizeof(xlint), 4);
    BOOST_TEST_EQ(sizeof(xlsref), 20);
    BOOST_TEST_EQ(sizeof(xlref), 8);
    BOOST_TEST_EQ(sizeof(xlmulti), 12);
    BOOST_TEST_EQ(sizeof(xlflow), 16);
    BOOST_TEST_EQ(sizeof(xlbigdata), 8);
#else
    BOOST_TEST_EQ(sizeof(xlnum), 8);
    BOOST_TEST_EQ(sizeof(xlstr), 8);
    BOOST_TEST_EQ(sizeof(xlbool), 4);
    BOOST_TEST_EQ(sizeof(xlerr), 4);
    BOOST_TEST_EQ(sizeof(xlint), 4);
    BOOST_TEST_EQ(sizeof(xlsref), 20);
    BOOST_TEST_EQ(sizeof(xlref), 16);
    BOOST_TEST_EQ(sizeof(xlmulti), 16);
    BOOST_TEST_EQ(sizeof(xlflow), 24);
    BOOST_TEST_EQ(sizeof(xlbigdata), 16);
#endif

    BOOST_TEST_EQ(sizeof(variant), 32);

    BOOST_TEST_TRAIT_TRUE((std::is_same<detail::resolve_overload_type<const wchar_t *, xlbool, xlstr>, xlstr>));
    BOOST_TEST_TRAIT_TRUE((std::is_same<detail::resolve_overload_type<const char *, xlbool, xlstr>, xlstr>));
    BOOST_TEST_TRAIT_TRUE((std::is_same<detail::resolve_overload_type<xlbool, xlstr, xlbool, xlint>, xlbool>));
    BOOST_TEST_TRAIT_TRUE((std::is_same<detail::resolve_overload_type<xlstr, xlstr, xlbool>, xlstr>));
    BOOST_TEST_TRAIT_TRUE((std::is_same<detail::resolve_overload_type<int, xlnum, xlint, xlbool, xlerr, xlmulti>, xlint>));
    BOOST_TEST_TRAIT_TRUE((std::is_same<detail::resolve_overload_type<double, xlnum, xlint, xlbool, xlerr, xlmulti>, xlnum>));
    BOOST_TEST_TRAIT_TRUE((std::is_same<detail::resolve_overload_type<error::excel_error, xlnum, xlint, xlbool, xlerr, xlmulti>, xlerr>));

    // Requires explicit conversion
    BOOST_TEST_TRAIT_FALSE((std::is_same<detail::resolve_overload_type<bool, xlbool, xlint>, xlbool>));

    {
        variant v(L"TEST");
        BOOST_TEST_EQ(v.xltype(), xltypeStr);
        BOOST_TEST_EQ(v.get<xlstr>(), L"TEST");
    }

    {
        variant v1(L"TEST");
        variant v2(v1);
        BOOST_TEST_EQ(v2.xltype(), xltypeStr);
        BOOST_TEST_EQ(v2.get<xlstr>(), L"TEST");
    }
/*
    {
        variant v(5);
        BOOST_TEST_EQ(v.xltype(), xltypeInt);
        BOOST_TEST_EQ(v.get<xlint>(), 5);
    }

    {
        variant v(10.5);
        BOOST_TEST_EQ(v.xltype(), xltypeNum);
        BOOST_TEST_EQ(v.get<xlnum>(), 10.5);
    }
    
    {
        variant v((xlbool)true);
        BOOST_TEST_EQ(v.xltype(), xltypeBool);
        BOOST_TEST_EQ(v.get<xlbool>(), true);
        v = (xlbool)false;
        BOOST_TEST_EQ(v.xltype(), xltypeBool);
        BOOST_TEST_EQ(v.get<xlbool>(), false);
    }

    {
        std::vector<variant> vv({ L"TEST1" });
        vv.emplace_back("TEST2");
        BOOST_TEST_EQ(vv.at(0).get<xlstr>(), "TEST1");
        BOOST_TEST_EQ(vv.at(1).get<xlstr>(), "TEST2");
    }

    {
        xlmulti m({{"S1", L"S2", 3}, {4}, {7, 8.1, 9}});
        BOOST_TEST_EQ(m.size(), 9);
        BOOST_TEST_EQ(m.at(0, 0).get<xlstr>(), "S1");
        BOOST_TEST_EQ(m.at(0, 1).get<xlstr>(), L"S2");
        BOOST_TEST_EQ(m.at(0, 2).get<xlint>(), 3);
        BOOST_TEST_EQ(m.at(1, 0).get<xlint>(), 4);
        BOOST_TEST_EQ(m.at(1, 1).xltype(), xltypeNil);
        BOOST_TEST_EQ(m.at(1, 2).xltype(), xltypeNil);
        BOOST_TEST_EQ(m.at(2, 0).get<xlint>(), 7);
        BOOST_TEST_EQ(m.at(2, 1).get<xlnum>(), 8.1);
        BOOST_TEST_EQ(m.at(2, 2).get<xlint>(), 9);
    }
*/
    return boost::report_errors();
}