//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <cstdint>
#include <memory>
#include <new>
#include <type_traits>

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

// Used for empty base optimization.
template <class Allocator>
struct alloc_holder : private Allocator
{
    using ebo_eligible = std::conjunction<
        std::is_empty<Allocator>,
        std::is_copy_constructible<Allocator>>;

    static_assert(ebo_eligible::value, "allocator must be stateless");

    alloc_holder() = default;

    explicit alloc_holder(const Allocator& alloc)
        noexcept(std::is_nothrow_copy_constructible_v<Allocator>)
        : Allocator { alloc } {}

    const Allocator& alloc() const noexcept { return *this; }

    Allocator& alloc() noexcept { return *this; }
};

// RAII wrapper for exception-safe allocations.

template <class Allocator>
struct construct_guard
{
    using pointer = typename std::allocator_traits<Allocator>::pointer;

    construct_guard(Allocator& a, pointer p) noexcept
        : a_(a), p_(p), n_(1) {}

    construct_guard(Allocator& a, pointer p, std::size_t n) noexcept
        : a_(a), p_(p), n_(n) {}

    construct_guard(const construct_guard&) = delete;
    construct_guard& operator=(const construct_guard&) = delete;

    ~construct_guard() {
        if (p_) { a_.deallocate(p_, n_); }
    }

    void release() { p_ = pointer(); }

private:
    Allocator& a_;
    pointer p_;
    std::size_t n_;
};

} // namespace detail
} // namespace xll