//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <xll/detail/assert.hpp>
#include <xll/detail/memory.hpp>

#include <algorithm>
#include <iterator>
#include <limits>
#include <memory>
#include <type_traits>

#include <boost/container/vector.hpp>
#include <boost/numeric/ublas/storage.hpp>

namespace xll {
namespace detail {

// Boost.uBLAS Storage concept using a fixed size buffer and configurable
// stored size type (SST). Used with FP12 arrays.

template <class T, std::size_t N, class SST = std::size_t>
struct static_array
    : public boost::numeric::ublas::storage_array<static_array<T, N, SST>>
{
    static_assert(N <= std::numeric_limits<SST>::max());

    using value_type = T;
    using size_type = SST;
    using difference_type = std::make_signed_t<SST>;
    using const_reference = const T&;
    using reference = T&;
    using const_pointer = const T*;
    using pointer = T*;
    using const_iterator = const_pointer;
    using iterator = pointer;
    using const_reverse_iterator = std::reverse_iterator<const_iterator>;
    using reverse_iterator = std::reverse_iterator<iterator>;

public:
    inline static_array() noexcept
    {}

    explicit inline static_array(size_type size)
    {
        (void)size;
        XLL_ASSERT(size == N);
        std::uninitialized_value_construct(begin(), end());
    }

    inline static_array(size_type size, value_type init)
    {
        XLL_ASSERT(size == N);
        std::uninitialized_fill(begin(), end(), init);
    }

    inline static_array(const static_array& rhs)
    {
        std::uninitialized_copy(rhs.begin(), rhs.end(), begin());
    }

    inline static_array(static_array&& rhs)
        noexcept(std::is_nothrow_move_constructible_v<value_type>)
    {
        std::uninitialized_move(rhs.begin(), rhs.end(), begin());
    }

    inline size_type size() const
    {
        return static_cast<size_type>(N);
    }

    inline size_type max_size() const
    {
        return static_cast<size_type>(N);
    }

    inline const_reference operator[](size_type i) const
    {
        XLL_ASSERT(i < N);
        return begin()[i];
    }
    
    inline reference operator[](size_type i)
    {
        XLL_ASSERT(i < N);
        return begin()[i];
    }

    inline static_array& operator=(const static_array& rhs)
    {
        if (this != &rhs) {
            std::uninitialized_copy(rhs.begin(), rhs.end(), begin());
        }
        return *this;
    }

    inline static_array& operator=(static_array&& rhs)
    {
        if (this != &rhs) {
            std::uninitialized_move(rhs.begin(), rhs.end(), begin());
        }
        return *this;
    }

    inline static_array& assign_temporary(static_array& a)
    {
        *this = a;
        return *this;
    }

    inline void swap(static_array& rhs)
    {
        if (this != &rhs) {
            std::swap_ranges(begin(), begin() + N, rhs.begin());
        }
    }
    
    friend void swap(static_array &a1, static_array &a2) {
        a1.swap(a2);
    }

    inline const_iterator begin() const
    {
        return reinterpret_cast<const_iterator>(&data_);
    }

    inline const_iterator cbegin() const
    {
        return reinterpret_cast<const_iterator>(&data_);
    }

    inline const_iterator end() const
    {
        return reinterpret_cast<const_iterator>(&data_) + N;
    }
    
    inline const_iterator cend() const
    {
        return reinterpret_cast<const_iterator>(&data_) + N;
    }

    inline iterator begin()
    {
        return reinterpret_cast<iterator>(&data_);
    }
    
    inline iterator end()
    {
        return reinterpret_cast<iterator>(&data_) + N;
    }

    inline const_reverse_iterator rbegin() const
    {
        return const_reverse_iterator(end());
    }
    
    inline const_reverse_iterator crbegin() const
    {
        return const_reverse_iterator(end());
    }
    
    inline const_reverse_iterator rend() const
    {
        return const_reverse_iterator(begin());
    }
    
    inline const_reverse_iterator crend() const
    {
        return const_reverse_iterator(begin());
    }
    
    inline reverse_iterator rbegin()
    {
        return reverse_iterator(end());
    }
    
    inline reverse_iterator rend()
    {
        return reverse_iterator(begin());
    }

private:
    std::aligned_storage_t<sizeof(T)*N, alignof(T)> data_;
};

} // namespace detail
} // namespace xll