//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#include <xll/xll.hpp>

#include <boost/core/lightweight_test.hpp>

using namespace xll;

int main()
{
    using namespace std::string_literals;
    using namespace std::string_view_literals;

    {
        // compile-time wide pstring literal
        constexpr auto wide = make_wpstring_literal(L"Literal");
        BOOST_TEST(std::wstring_view(wide) == L"Literal");
        BOOST_TEST_EQ(wide.size(), 7);
    }
    {
        // compile-time narrow pstring literal
        constexpr auto narrow = make_pstring_literal("Literal");
        BOOST_TEST(std::string_view(narrow) == "Literal");
        BOOST_TEST_EQ(narrow.size(), 7);
    }
    {
        // compile-time wide pstring array
        constexpr std::array<wchar_t, 4> arr {{ L'T', L'E', L'S', L'T' }};
        constexpr auto wide = make_wpstring_array(arr);
        BOOST_TEST(std::wstring_view(wide) == L"TEST");
        BOOST_TEST_EQ(wide.size(), 4);
    }
    {
        // compile-time narrow pstring array
        constexpr std::array<char, 4> arr {{ 'T', 'E', 'S', 'T' }};
        constexpr auto narrow = make_pstring_array(arr);
        BOOST_TEST(std::string_view(narrow) == "TEST");
        BOOST_TEST_EQ(narrow.size(), 4);
    }
    {
        // narrow pstring with no conversion
        pstring ps(std::string("Test"));
        BOOST_TEST_EQ(ps.size(), 4);
        BOOST_TEST(ps == "Test"s);
        BOOST_TEST(ps == L"Test"s);
        BOOST_TEST(ps == "Test"sv);
    }
    {
        // narrow pstring with UTF-16 to UTF-8 conversion
        pstring ps(std::wstring(L"Test"));
        BOOST_TEST_EQ(ps.size(), 4);
        BOOST_TEST(ps == "Test"s);
        BOOST_TEST(ps == L"Test"s);
        BOOST_TEST(ps == "Test"sv);
    }
    {
        // wide pstring with UTF-8 to UTF-16 conversion
        wpstring wps(std::string("Test"));
        BOOST_TEST_EQ(wps.size(), 4);
        BOOST_TEST(wps == "Test"s);
        BOOST_TEST(wps == L"Test"s);
        BOOST_TEST(wps == L"Test"sv);
    }
    {
        // wide pstring with no conversion
        wpstring wps(std::wstring(L"Test"));
        BOOST_TEST_EQ(wps.size(), 4);
        BOOST_TEST(wps == "Test"s);
        BOOST_TEST(wps == L"Test"s);
        BOOST_TEST(wps == L"Test"sv);
    }
    {
        // copy construct
        pstring ps1(std::string("Copy"));
        pstring ps2(ps1);
        BOOST_TEST_EQ(ps1.size(), 4);
        BOOST_TEST_EQ(ps2.size(), 4);
        BOOST_TEST(ps1 == "Copy");
        BOOST_TEST(ps2 == "Copy");
    }
    {
        // move construct
        pstring ps1(std::string("Move"));
        BOOST_TEST_EQ(ps1.size(), 4);
        pstring ps2(std::move(ps1));
        BOOST_TEST_EQ(ps1.size(), 0);
        BOOST_TEST_EQ(ps2.size(), 4);
        BOOST_TEST(ps2 == "Move");
    }

    return boost::report_errors();
}