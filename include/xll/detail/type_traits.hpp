//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <type_traits>

namespace xll {
namespace detail {

// SFINAE-friendly version of std::underlying_type.

template <class T, bool = std::is_enum_v<T>>
struct safe_underlying_type {
    using type = typename std::underlying_type<T>::type;
};

template <class T>
struct safe_underlying_type<T, false> {
    using type = T;
};

template <class T>
using safe_underlying_type_t = typename safe_underlying_type<T>::type;

// Returns a decayed type for extern "C" linkage. Applies enumeration to
// underlying_type, reference-to-pointer, and array-to-pointer conversions to
// the type T, removes cv-qualifiers, and defines the resulting type as the
// member type alias `type`.
//
// While technically implementation-defined, this assumes that references are
// implemented using a pointer variable.

template <class T>
struct extern_c_type {
private:
  using U =
    std::conditional_t<
      std::is_reference_v<T>,
      std::remove_reference_t<T> *,
      T>;

public:
  using type =
    std::conditional_t<
      std::is_array_v<U>,
      std::remove_extent_t<U> *,
      safe_underlying_type_t<std::remove_cv_t<U>>
    >;
};

template <class T>
using extern_c_type_t = typename extern_c_type<T>::type;

} // namespace detail
} // namespace xll
