//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <cstdint>
#include <new>
#include <intrin.h>

namespace xll {
namespace detail {

template<class T, class U>
T launder_cast(U* u)
{
#ifdef __cpp_lib_launder // P0137R1 (C++17)
    return std::launder(reinterpret_cast<T>(u));
#else
    return reinterpret_cast<T>(u);
#endif
}

BOOST_FORCEINLINE std::uintptr_t stack_base()
{
#ifdef _WIN64
    return __readgsqword(0x08); // offsetof(NT_TIB64, StackBase)
#else
    return __readfsdword(0x04); // offsetof(NT_TIB, StackBase)
#endif
}

BOOST_FORCEINLINE std::uintptr_t stack_limit()
{
#ifdef _WIN64
    return __readgsqword(0x10); // offsetof(NT_TIB64, StackLimit)
#else
    return __readfsdword(0x08); // offsetof(NT_TIB, StackLimit)
#endif
}

} // namespace detail
} // namespace xll