//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <boost/mp11/utility.hpp>
#include <boost/winapi/character_code_conversion.hpp>

#include <algorithm>
#include <array>
#include <limits>
#include <memory>
#include <ostream>
#include <string>
#include <string_view>
#include <tuple>
#include <type_traits>
#include <utility>

namespace xll {
namespace detail {

template <class CharT, std::size_t N>
struct basic_pstring_literal_impl
{
    static_assert(N < std::numeric_limits<CharT>::max(),
        "String length exceeds Excel 12 limit");

    using value_type = CharT;
    using size_type = std::make_unsigned_t<CharT>;
    using difference_type = std::make_signed_t<CharT>;

    using xltype =
      std::conditional_t<
        sizeof(CharT) == sizeof(wchar_t),
        tag::xlstr::xltype,
        void>;

protected:
    const CharT data_[1 + N];

public:
    template <class... Chars>
    constexpr basic_pstring_literal_impl(Chars... args)
        : data_{static_cast<CharT>(N), args...}
    {}

    constexpr const CharT * data() const {
        return &data_[1];
    }

    constexpr size_type size() const {
        return N;
    }

    constexpr size_type length() const {
        return N;
    }

    constexpr explicit operator std::basic_string_view<CharT>() const {
        if constexpr (N == 0)
            return std::basic_string_view<CharT>();
        return { &data_[1], N };
    }

    explicit operator std::basic_string<CharT>() const {
        if constexpr (N == 0)
            return std::basic_string<CharT>();
        return std::basic_string<CharT>(&data_[1], N);
    }

    friend bool operator==(const basic_pstring_literal_impl& lhs, const basic_pstring_literal_impl& rhs) {
        return std::basic_string_view<CharT>{ lhs.data(), lhs.size() } ==
               std::basic_string_view<CharT>{ rhs.data(), rhs.size() };
    }

    friend std::basic_ostream<CharT>& operator<<(std::basic_ostream<CharT>& os, const basic_pstring_literal_impl& s) {
        return os << std::basic_string_view<CharT>{ s.data(), s.size() };
    }
};

} // namespace detail

/// Generates pascal string from a C string literal at compile-time.
template <class CharT, std::size_t N>
struct basic_pstring_literal : public detail::basic_pstring_literal_impl<CharT, N - 1>
{
protected:
    template <std::size_t... Is>
    constexpr basic_pstring_literal(const CharT(&str)[N], std::index_sequence<Is...>)
        : detail::basic_pstring_literal_impl<CharT, N - 1>(str[Is]...)
    {}

public:
    constexpr basic_pstring_literal(const CharT(&str)[N])
        : basic_pstring_literal(str, std::make_index_sequence<N - 1>{})
    {}
};

/// Generates pascal string from std::array at compile-time.
template <class CharT, std::size_t N>
struct basic_pstring_array : public detail::basic_pstring_literal_impl<CharT, N>
{
protected:
    template <std::size_t... Is>
    constexpr basic_pstring_array(const std::array<CharT, N>& arr, std::index_sequence<Is...>)
        : detail::basic_pstring_literal_impl<CharT, N>(arr[Is]...)
    {}

public:
    constexpr basic_pstring_array(const std::array<CharT, N>& arr)
        : basic_pstring_array(arr, std::make_index_sequence<N>{})
    {}

    constexpr basic_pstring_array(std::array<CharT, N>&& arr)
        : basic_pstring_array(std::forward<std::array<CharT, N>>(arr))
    {}
};

template <std::size_t N>
using pstring_literal = basic_pstring_literal<char, N>;

template <std::size_t N>
using wpstring_literal = basic_pstring_literal<wchar_t, N>;

template <std::size_t N>
using pstring_array = basic_pstring_array<char, N>;

template <std::size_t N>
using wpstring_array = basic_pstring_array<wchar_t, N>;

template <std::size_t N>
constexpr pstring_literal<N> make_pstring_literal(const char(&str)[N])
{
    return pstring_literal<N>(str);
}

template <std::size_t N>
constexpr wpstring_literal<N> make_wpstring_literal(const wchar_t(&str)[N])
{
    return wpstring_literal<N>(str);
}

template <std::size_t N>
constexpr pstring_array<N> make_pstring_array(std::array<char, N>&& arr)
{
    return wpstring_array<N>(arr);
}

template <std::size_t N>
constexpr pstring_array<N> make_pstring_array(const std::array<char, N>& arr)
{
    return wpstring_array<N>(arr);
}

template <std::size_t N>
constexpr wpstring_array<N> make_wpstring_array(std::array<wchar_t, N>&& arr)
{
    return wpstring_array<N>(arr);
}

template <std::size_t N>
constexpr wpstring_array<N> make_wpstring_array(const std::array<wchar_t, N>& arr)
{
    return wpstring_array<N>(arr);
}

} // namespace xll

/// Allocates new pascal string at runtime, performing UTF-8/UTF-16 conversion
/// as necessary using Win32 API (MultiByteToWideChar, WideCharToMultiByte).
/// For compile-time fixed strings that don't require character set conversion,
/// use xll::basic_pstring_literal.
///
/// Boost.Nowide, Boost.Text, ICU could be used as alternatives. 
/// std::wstring_convert and std::codecvt_utf8_utf16 are depreciated in C++17.

namespace xll {

template <class CharT>
struct basic_pstring
{
    static_assert(sizeof(CharT) <= sizeof(wchar_t), "Unsupported character type.");

    using xltype =
      std::conditional_t<
        sizeof(CharT) == sizeof(wchar_t),
        tag::xlstr::xltype,
        void>;
    
    using value_type = CharT;
    using size_type = std::make_unsigned_t<CharT>;
    using difference_type = std::make_signed_t<CharT>;

protected:
    CharT *data_ = nullptr;

    void copy_construct(const CharT *s, size_type n)
    {
        std::size_t cb = (n + 1) * sizeof(CharT);
        data_ = static_cast<CharT *>(::operator new (cb));
        data_[0] = n;
        if (s != nullptr && n > 0)
            std::copy_n(s, n, &data_[1]);
    }

    void copy_construct(const basic_pstring& rhs)
    {
        if (rhs.data_ != nullptr) {
            const size_type n = rhs.data_[0];
            copy_construct(&rhs.data_[1], n);
        }
    }
    
    void copy_construct(const std::basic_string<CharT>& s)
    {
        const size_type n = static_cast<size_type>(
            std::char_traits<CharT>::length(s.c_str())); // truncate
        copy_construct(s.data(), n);
    }

    void move_construct(basic_pstring& rhs)
    {
        data_ = std::exchange(rhs.data_, nullptr);
    }

    void destroy()
    {
        if (data_ != nullptr) {
            delete data_;
            data_ = nullptr;
        }
    }

public:
    basic_pstring() = default;

    ~basic_pstring()
    {
        destroy();
    }

    basic_pstring(const basic_pstring &rhs)
    {
        copy_construct(rhs);
    }
    
    basic_pstring(basic_pstring&& rhs)
    {
        move_construct(rhs);
    }

    basic_pstring& operator=(const basic_pstring& rhs)
    {
        destroy();
        copy_construct(rhs);
    }

    basic_pstring& operator=(basic_pstring&& rhs)
    {
        destroy();
        move_construct(rhs);
        return *this;
    }

    template <std::size_t N>
    basic_pstring(const detail::basic_pstring_literal_impl<CharT, N>& s)
    {
        copy_construct(s.data(), static_cast<size_type>(N));
    }

    explicit basic_pstring(const std::basic_string<CharT>& s)
    {
        copy_construct(s);
    }
    
    // UTF-8 to UTF-16
    template <typename FromCharT = T, std::enable_if_t<(sizeof(FromCharT) < sizeof(CharT))>* = nullptr>
    explicit basic_pstring(const std::basic_string<FromCharT>& s)
    {
        using namespace boost::winapi;

        const size_type len = static_cast<size_type>(
            std::char_traits<FromCharT>::length(s.c_str()));
        const size_type n = static_cast<size_type>(
            MultiByteToWideChar(CP_UTF8_, 0, s.data(), (int)len, nullptr, 0)); // truncate

        std::size_t cb = (n + 1) * sizeof(CharT);
        data_ = static_cast<CharT *>(::operator new (cb));
        data_[0] = n;
        if (n > 0)
            MultiByteToWideChar(CP_UTF8_, 0, s.data(), (int)len, &data_[1], (int)n);
    }

    // UTF-16 to UTF-8
    template <typename FromCharT = T, std::enable_if_t<(sizeof(FromCharT) > sizeof(CharT))>* = nullptr>
    explicit basic_pstring(const std::basic_string<FromCharT>& s)
    {
        using namespace boost::winapi;

        const size_type len = static_cast<size_type>(
            std::char_traits<FromCharT>::length(s.c_str()));
        const size_type n = static_cast<size_type>(
            WideCharToMultiByte(CP_UTF8_, 0, s.data(), (int)len, nullptr, 0, nullptr, nullptr)); // truncate
        
        std::size_t cb = (n + 1) * sizeof(CharT);
        data_ = static_cast<CharT *>(::operator new (cb));
        data_[0] = n;
        if (n > 0)
            WideCharToMultiByte(CP_UTF8_, 0, s.data(), (int)len, &data_[1], (int)n, nullptr, nullptr);
    }

    explicit operator std::basic_string<CharT>() const
    {
        const std::size_t n = size();
        if (n == 0)
            return std::basic_string<CharT>();
        return std::basic_string<CharT>(&data_[1], &data_[1 + n]);
    }

    // UTF-16 to UTF-8
    template <typename ToCharT = T, std::enable_if_t<(sizeof(ToCharT) < sizeof(CharT))>* = nullptr>
    explicit operator std::basic_string<ToCharT>() const
    {
        using namespace boost::winapi;

        const std::size_t n0 = size();
        if (n0 == 0)
            return std::basic_string<ToCharT>();
        
        const size_type n1 = static_cast<size_type>(
            WideCharToMultiByte(CP_UTF8_, 0, &data_[1], (int)n0, nullptr, 0, nullptr, nullptr));
        
        std::basic_string<ToCharT> out(n1, '\0');
        if (n1 > 0)
            WideCharToMultiByte(CP_UTF8_, 0, &data_[1], (int)n0, out.data(), (int)n1, nullptr, nullptr);

        return out;
    }

    // UTF-8 to UTF-16
    template <typename ToCharT = T, std::enable_if_t<(sizeof(ToCharT) > sizeof(CharT))>* = nullptr>
    explicit operator std::basic_string<ToCharT>() const
    {
        using namespace boost::winapi;
        
        const std::size_t n0 = size();
        if (n0 == 0)
            return std::basic_string<ToCharT>();

        const size_type n1 = static_cast<size_type>(
            MultiByteToWideChar(CP_UTF8_, 0, &data_[1], (int)n0, nullptr, 0));

        std::basic_string<ToCharT> out(n1, L'\0');
        if (n1 > 0)
            MultiByteToWideChar(CP_UTF8_, 0, &data_[1], (int)n0, out.data(), (int)n1);

        return out;
    }

    std::size_t size() const 
    {
        if (data_ == nullptr)
            return 0;
        return data_[0];
    }

    std::size_t length() const 
    {
        return size();
    }

    bool empty() const
    {
        return size() == 0;
    }

    CharT * data() const {
        if (empty())
            return nullptr;
        return &data_[1];
    }

    explicit operator std::basic_string_view<CharT>() const
    {
        const std::size_t n = size();
        if (n == 0)
            return std::basic_string_view<CharT>();
        return { &data_[1], n };
    }

    friend bool operator==(const basic_pstring& lhs, const basic_pstring& rhs)
    {
        const std::size_t n = lhs.size();
        if (n != rhs.size())
            return false;
        return std::basic_string_view<CharT>{ lhs.data(), n } ==
               std::basic_string_view<CharT>{ rhs.data(), n };
    }

    friend std::basic_ostream<CharT>& operator<<(std::basic_ostream<CharT>& os, const basic_pstring& s)
    {
        const std::size_t n = s.size();
        if (n == 0)
            return os;
        return os << std::basic_string_view<CharT>{ &s.data_[1], n };
    }
};

using pstring = basic_pstring<char>;
using wpstring = basic_pstring<wchar_t>;

static_assert(sizeof(pstring) == sizeof(void *));
static_assert(sizeof(wpstring) == sizeof(void *));

} // namespace xll
