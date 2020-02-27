//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

/**
 * \file type_text.hpp
 * Compile-time mapping from primitive type to pxTypeText identifier wchar array.
 * \sa https://docs.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1#data-types
 */

#include <xll/config.hpp>

#include <xll/detail/type_traits.hpp>
#include <xll/fp12.hpp>
#include <xll/xloper12.hpp>
#include <xll/pstring.hpp>

#include <boost/mp11/list.hpp>
#include <boost/mp11/algorithm.hpp>
#include <boost/mp11/function.hpp>

#include <array>
#include <cstdint>
#include <functional>
#include <tuple>
#include <type_traits>

namespace xll {
namespace detail {

template <typename T>
struct is_array_type : std::false_type {};

template <std::size_t N>
struct is_array_type<static_fp12<N> *> : std::true_type {};

// Variable-type worksheet values and arrays (XLOPER12)
template <class T, class U = void>
struct type_text_arg {
    static_assert(std::is_base_of_v<detail::tagged_union, T>, "Invalid operand type");
    static constexpr std::array<wchar_t, 1> value = { L'Q' };
};

template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, variant12 *>>> {
    static constexpr std::array<wchar_t, 1> value = { L'Q' };
};

// Asynchronous call handle (XLOPER12, xlTypeBigData)
template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, handle12 *>>> {
    static constexpr std::array<wchar_t, 1> value = { L'X' };
};

// Larger grid floating-point array structure (FP12)
template <class T>
struct type_text_arg<T, std::enable_if_t<is_array_type<T>::value>> {
    static constexpr std::array<wchar_t, 2> value = { L'K', L'%' };
};

// Asynchronous UDF return type
template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, void>>> {
    static constexpr std::array<wchar_t, 1> value = { L'>' };
};

// Integral types
template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, bool>>> {
    static constexpr std::array<wchar_t, 1> value = { L'A' };
};

template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, bool *>>> {
    static constexpr std::array<wchar_t, 1> value = { L'L' };
};

template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, uint16_t>>> {
    static constexpr std::array<wchar_t, 1> value = { L'H' };
};

template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, int16_t>>> {
    static constexpr std::array<wchar_t, 1> value = { L'I' };
};

template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, int16_t *>>> {
    static constexpr std::array<wchar_t, 1> value = { L'M' };
};

template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, int32_t>>> {
    static constexpr std::array<wchar_t, 1> value = { L'J' };
};

template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, int32_t *>>> {
    static constexpr std::array<wchar_t, 1> value = { L'N' };
};

// Floating-point types
template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, double>>> {
    static constexpr std::array<wchar_t, 1> value = { L'B' };
};

template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, double *>>> {
    static constexpr std::array<wchar_t, 1> value = { L'E' };
};

// Null-terminated ASCII byte string (max. 256 characters)
template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, char *>>> {
    static constexpr std::array<wchar_t, 1> value = { L'C' };
};

template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, const char *>>> {
    static constexpr std::array<wchar_t, 1> value = { L'C' };
};

// Null-terminated Unicode wide-character string (max. 32767 characters)
template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, wchar_t *>>> {
    static constexpr std::array<wchar_t, 2> value = { L'C', L'%' };
};

template <class T>
struct type_text_arg<T, std::enable_if_t<std::is_same_v<T, const wchar_t *>>> {
    static constexpr std::array<wchar_t, 2> value = { L'C', L'%' };
};

template <class A>
struct attribute_text_arg;

// Function Attributes
template <>
struct attribute_text_arg<cluster_safe> {
    static constexpr std::array<wchar_t, 1> value = { L'&' };
};

template <>
struct attribute_text_arg<volatile_udf> {
    static constexpr std::array<wchar_t, 1> value = { L'!' };
};

template <>
struct attribute_text_arg<thread_safe> {
    static constexpr std::array<wchar_t, 1> value = { L'$' };
};

// Get a pxTypeText wchar array (null-terminated) for a callable type.
// Concatenates the pxTypeText wchar arrays for the return type and arguments
// at compile-time using tuples.
//
// Requires compiler support for std::tuple_cat with std::array, and C++17
// for std::apply and std::invoke_result_t.
//
// TODO:
// - Ensure that arguments are const qualified.
// - Async
//   - One 'X' parameter which will store the asynchronous call handle (xlTypeBigData)
//   - Cannot be combined with Cluster Safe
//   - void return type, '>'
//   - separate from variant12 - should be its own type ('xlhandle'?).
// - Cluster safe
//   - No XLOPER12 arguments that support range references (type 'U').
//   - Add '&' to end of type text
// - Volatile
//   - Add '!' to end of type text
// - Thread safe
//   - Add '$' to end of type text

template<class Result, class... Args>
constexpr auto type_text_impl()
{
    using namespace boost::mp11;

    static_assert(!mp_any<std::is_void<Args>...>::value,
        "Arguments cannot be void");
    
    using async_handle_count = mp_count<mp_list<Args...>, handle12 *>;
    static_assert(mp_less<async_handle_count, mp_size_t<2>>::value,
        "Multiple async handles in argument list");

    using has_async_handle = mp_to_bool<async_handle_count>;
    using has_void_return = std::is_void<Result>;
    static_assert(!std::is_same_v<has_async_handle, has_void_return>,
        "Async functions must have void return type");
        
    // Construct tuple using std::tuple_cat specialization for std:array.
    constexpr auto tuple = std::tuple_cat(
        type_text_arg<extern_type_t<Result>>::value,
        type_text_arg<extern_type_t<Args>>::value...);
        //attribute_text_arg<Attrs>::value...);

    // Construct wchar array from the flat tuple.
    constexpr auto make_array = [](auto&& ...x) {
        return std::array{x...};
    };

    return std::apply(make_array, tuple);
}

// Overloads for various calling conventions and attributes.

template<class F, class... Args>
constexpr auto type_text(F(__cdecl *)(Args...))
{
    using result_t = std::invoke_result_t<F(Args...), Args...>;
    return detail::type_text_impl<result_t, Args...>();
}

template<class F, class... Args>
constexpr auto type_text(F (__stdcall *)(Args...))
{
    using result_t = std::invoke_result_t<F(Args...), Args...>;
    return detail::type_text_impl<result_t, Args...>();
}

template<class F, class... Args>
constexpr auto type_text(F (__fastcall *)(Args...))
{
    using result_t = std::invoke_result_t<F(Args...), Args...>;
    return detail::type_text_impl<result_t, Args...>();
}

template<class F, class... Args>
constexpr auto type_text(F (__vectorcall *)(Args...))
{
    using result_t = std::invoke_result_t<F(Args...), Args...>;
    return detail::type_text_impl<result_t, Args...>();
}

} // namespace detail
} // namespace xll