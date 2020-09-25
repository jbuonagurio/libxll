//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <boost/mp11/utility.hpp>
#include <boost/nowide/convert.hpp> // Boost 1.73.0

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

template<class CharT, std::size_t N>
struct basic_pstring_literal_impl
{
    static_assert(N < std::numeric_limits<CharT>::max(),
        "string length exceeds Excel 12 limit");

    using value_type = CharT;
    using size_type = std::make_unsigned_t<CharT>;
    using difference_type = std::make_signed_t<CharT>;

protected:
    const CharT data_[1 + N];

public:
    template<class... Chars>
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

    constexpr operator std::basic_string_view<CharT>() const {
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
        os << std::basic_string_view<CharT>{ s.data(), s.size() };
        return os;
    }
};

} // namespace detail

/// Generates pascal string from a C string literal at compile-time.
template<class CharT, std::size_t N>
struct basic_pstring_literal : public detail::basic_pstring_literal_impl<CharT, N - 1>
{
protected:
    template<std::size_t... Is>
    constexpr basic_pstring_literal(const CharT(&str)[N], std::index_sequence<Is...>)
        : detail::basic_pstring_literal_impl<CharT, N - 1>(str[Is]...)
    {}

public:
    constexpr basic_pstring_literal(const CharT(&str)[N])
        : basic_pstring_literal(str, std::make_index_sequence<N - 1>{})
    {}
};

/// Generates pascal string from std::array at compile-time.
template<class CharT, std::size_t N>
struct basic_pstring_array : public detail::basic_pstring_literal_impl<CharT, N>
{
protected:
    template<std::size_t... Is>
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

template<std::size_t N>
using pstring_literal = basic_pstring_literal<char, N>;

template<std::size_t N>
using wpstring_literal = basic_pstring_literal<wchar_t, N>;

template<std::size_t N>
using pstring_array = basic_pstring_array<char, N>;

template<std::size_t N>
using wpstring_array = basic_pstring_array<wchar_t, N>;

template<std::size_t N>
constexpr pstring_literal<N> make_pstring_literal(const char(&str)[N])
{
    return pstring_literal<N>(str);
}

template<std::size_t N>
constexpr wpstring_literal<N> make_wpstring_literal(const wchar_t(&str)[N])
{
    return wpstring_literal<N>(str);
}

template<std::size_t N>
constexpr pstring_array<N> make_pstring_array(std::array<char, N>&& arr)
{
    return pstring_array<N>(arr);
}

template<std::size_t N>
constexpr pstring_array<N> make_pstring_array(const std::array<char, N>& arr)
{
    return pstring_array<N>(arr);
}

template<std::size_t N>
constexpr wpstring_array<N> make_wpstring_array(std::array<wchar_t, N>&& arr)
{
    return wpstring_array<N>(arr);
}

template<std::size_t N>
constexpr wpstring_array<N> make_wpstring_array(const std::array<wchar_t, N>& arr)
{
    return wpstring_array<N>(arr);
}

} // namespace xll

/// Allocates new pascal string at runtime, performing UTF-8/UTF-16 conversion
/// as necessary using Boost.Nowide. For compile-time fixed strings that don't
/// require character set conversion, use xll::basic_pstring_literal.
///
/// Win32 API (MultiByteToWideChar, WideCharToMultiByte), Boost.Text or ICU
/// could be used as alternatives. std::wstring_convert and
/// std::codecvt_utf8_utf16 are depreciated in C++17.

namespace xll {

template<class CharT> // class Allocator
struct basic_pstring
{
    static_assert(sizeof(CharT) <= sizeof(wchar_t), "unsupported character type");
    
    using value_type = CharT;
    using size_type = std::make_unsigned_t<CharT>;
    using difference_type = std::make_signed_t<CharT>;

protected:
    CharT *data_ = nullptr;

    inline void internal_copy(const CharT *s, size_type n)
    {
        std::size_t nbytes = (n + 1) * sizeof(CharT);
        data_ = static_cast<CharT *>(::operator new (nbytes)); // allocate
        data_[0] = n;
        if (s != nullptr && n > 0)
            std::copy_n(s, n, &data_[1]);
    }

    inline void internal_copy(const basic_pstring& rhs)
    {
        if (rhs.data_ != nullptr) {
            const size_type n = rhs.data_[0];
            internal_copy(&rhs.data_[1], n);
        }
    }
    
    inline void internal_copy(const std::basic_string<CharT>& s)
    {
        const size_type n = static_cast<size_type>(
            std::char_traits<CharT>::length(s.c_str())); // truncate
        internal_copy(s.data(), n);
    }

    inline void internal_copy(const CharT* s)
    {
        const size_type n = static_cast<size_type>(
            std::char_traits<CharT>::length(s)); // truncate
        internal_copy(s, n);
    }

    inline void destroy()
    {
        if (data_ != nullptr) {
            delete data_;
            data_ = nullptr;
        }
    }

public:
    basic_pstring() = default;

    ~basic_pstring()
        { destroy(); }

    basic_pstring(const basic_pstring &s)
        { internal_copy(s); }
    
    basic_pstring(basic_pstring&& s) noexcept
        : data_(s.data_)
    {
        if (this != &s) {
            data_ = std::exchange(s.data_, nullptr);
        }
    }

    basic_pstring& operator=(const basic_pstring& rhs)
    {
        destroy();
        internal_copy(rhs);
        return *this;
    }

    basic_pstring& operator=(basic_pstring&& rhs)
    {
        if (this != &rhs) {
            destroy();
            data_ = std::exchange(rhs.data_, nullptr);
        }
        return *this;
    }

    template<std::size_t N>
    explicit basic_pstring(const detail::basic_pstring_literal_impl<CharT, N>& s)
        { internal_copy(s.data(), static_cast<size_type>(N)); }

    // No Conversion
    basic_pstring(const CharT* s)
        { internal_copy(s); }

    // No Conversion
    basic_pstring(const std::basic_string<CharT>& s)
        { internal_copy(s); }
    
    // UTF-8 to UTF-16
    template<typename FromCharT = T, std::enable_if_t<(sizeof(FromCharT) < sizeof(CharT))>* = nullptr>
    basic_pstring(const FromCharT *s)
    {
        const size_type nchars = static_cast<size_type>(
            std::char_traits<FromCharT>::length(s)); // truncate
        std::basic_string<CharT> wide = boost::nowide::widen(s, nchars);
        internal_copy(wide);
    }

    // UTF-8 to UTF-16
    template<typename FromCharT = T, std::enable_if_t<(sizeof(FromCharT) < sizeof(CharT))>* = nullptr>
    basic_pstring(const std::basic_string<FromCharT>& s)
    {
        const FromCharT *cs = s.c_str();
        const size_type nchars = static_cast<size_type>(
            std::char_traits<FromCharT>::length(cs)); // truncate
        std::basic_string<CharT> wide = boost::nowide::widen(cs, nchars);
        internal_copy(wide);
    }

    // UTF-16 to UTF-8
    template<typename FromCharT = T, std::enable_if_t<(sizeof(FromCharT) > sizeof(CharT))>* = nullptr>
    basic_pstring(const FromCharT *s)
    {
        const size_type nchars = static_cast<size_type>(
            std::char_traits<FromCharT>::length(s)); // truncate
        std::basic_string<CharT> narrow = boost::nowide::narrow(cs, nchars);
        internal_copy(narrow);
    }

    // UTF-16 to UTF-8
    template<typename FromCharT = T, std::enable_if_t<(sizeof(FromCharT) > sizeof(CharT))>* = nullptr>
    basic_pstring(const std::basic_string<FromCharT>& s)
    {
        const FromCharT *cs = s.c_str();
        const size_type nchars = static_cast<size_type>(
            std::char_traits<FromCharT>::length(cs)); // truncate
        std::basic_string<CharT> narrow = boost::nowide::narrow(cs, nchars);
        internal_copy(narrow);
    }

    // No Conversion
    operator std::basic_string_view<CharT>() const
    {
        const std::size_t n = size();
        if (n == 0)
            return std::basic_string_view<CharT>();
        return { &data_[1], n };
    }

    // No Conversion
    operator std::basic_string<CharT>() const
    {
        const std::size_t n = size();
        if (n == 0)
            return std::basic_string<CharT>();
        return std::basic_string<CharT>(&data_[1], &data_[1 + n]);
    }

    // UTF-16 to UTF-8
    template<typename ToCharT = T, std::enable_if_t<(sizeof(ToCharT) < sizeof(CharT))>* = nullptr>
    operator std::basic_string<ToCharT>() const
        { return boost::nowide::narrow(data(), size()); }

    // UTF-8 to UTF-16
    template<typename ToCharT = T, std::enable_if_t<(sizeof(ToCharT) > sizeof(CharT))>* = nullptr>
    operator std::basic_string<ToCharT>() const
        { return boost::nowide::widen(data(), size()); }

    std::size_t size() const 
    {
        if (data_ == nullptr)
            return 0;
        return data_[0];
    }

    std::size_t length() const 
        { return size(); }

    bool empty() const
        { return size() == 0; }

    CharT * data() const {
        if (empty())
            return nullptr;
        return &data_[1];
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

    template<typename ToCharT = T, std::enable_if_t<(sizeof(ToCharT) != sizeof(CharT))>* = nullptr>
    friend std::basic_ostream<ToCharT>& operator<<(std::basic_ostream<ToCharT>& os, const basic_pstring& s)
        { return os << std::basic_string<ToCharT>(s); }
};

using pstring = basic_pstring<char>;
using wpstring = basic_pstring<wchar_t>;

static_assert(sizeof(pstring) == sizeof(void *), "invalid pstring size");
static_assert(sizeof(wpstring) == sizeof(void *), "invalid wpstring size");

} // namespace xll
