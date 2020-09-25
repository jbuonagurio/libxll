//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <xll/constants.hpp>
#include <xll/error.hpp>
#include <xll/pstring.hpp>
#include <xll/detail/memory.hpp>
#include <xll/detail/variant.hpp>

#include <algorithm>
#include <initializer_list>
#include <iterator>
#include <functional>
#include <stdexcept>
#include <vector>

namespace xll {

//
// XLOPER12 internal value types
//

struct xlnum
{
    using xltype = std::integral_constant<XLTYPE, xltypeNum>;
    xlnum() = default;
    xlnum(double x) : num(x) {}

    operator double() const noexcept
        { return num; }

    double num;
};

struct xlstr : public wpstring
{
    using xltype = std::integral_constant<XLTYPE, xltypeStr>;
    using wpstring::wpstring;
};

struct xlbool
{
    using xltype = std::integral_constant<XLTYPE, xltypeBool>;
    xlbool() = default;
    explicit xlbool(bool x) : xbool(x) {}
    
    operator bool() const noexcept
        { return xbool != 0 ? true : false; }

    int32_t xbool; // BOOL (INT32)
};

struct xlerr
{
    using xltype = std::integral_constant<XLTYPE, xltypeErr>;
    xlerr() = default;
    xlerr(error::excel_error x) : err(x) {}

    operator std::error_code() const noexcept
        { return error::make_error_code(err); }

    operator error::excel_error() const noexcept
        { return err; }

    error::excel_error err;
};

struct xlint
{
    using xltype = std::integral_constant<XLTYPE, xltypeInt>;
    xlint() = default;
    xlint(int32_t i) : w(i) {}
    
    operator int32_t() const noexcept
        { return w; }

    int32_t w;
};

struct xlsref
{
    using xltype = std::integral_constant<XLTYPE, xltypeSRef>;
    uint16_t count = 1; // WORD, always = 1
    struct {
        int32_t rwFirst = 0; // INT32
        int32_t rwLast = 0; // INT32
        int32_t colFirst = 0; // INT32
        int32_t colLast = 0; // INT32
    } ref;
};

struct xlref
{
    using xltype = std::integral_constant<XLTYPE, xltypeRef>;
    struct {
        uint16_t count = 1; // WORD
        struct {
            int32_t rwFirst = 0; // INT32
            int32_t rwLast = 0; // INT32
            int32_t colFirst = 0; // INT32
            int32_t colLast = 0; // INT32
        } reftbl[1]; // actually reftbl[count]
    } *lpmref = nullptr;
    uintptr_t idSheet; // DWORD_PTR (ULONG_PTR)
};

struct xlmissing
{
    using xltype = std::integral_constant<XLTYPE, xltypeMissing>;
};

struct xlnil
{
    using xltype = std::integral_constant<XLTYPE, xltypeNil>;
};

struct xlflow
{
    using xltype = std::integral_constant<XLTYPE, xltypeFlow>;
    union {
        int level; // xlflowRestart
        int tbctrl; // xlflowPause
        uintptr_t idSheet; // DWORD_PTR (ULONG_PTR), xlflowGoto
    } valflow;
    int32_t rw; // INT32, xlflowGoto
    int32_t col; // INT32, xlflowGoto
    XLFLOW xlflow; // BYTE
};

struct xlbigdata
{
    using xltype = std::integral_constant<XLTYPE, xltypeBigData>;
    void *h = nullptr; // BYTE*, HANDLE
    long cbData = 0;
};

namespace detail {

template <class T, class Allocator = std::allocator<T>>
struct xlmulti_base : detail::alloc_holder<Allocator>
{
    using xltype = std::integral_constant<XLTYPE, xltypeMulti>;

    using allocator_type = Allocator;
    using value_type = T;
    using element_type = value_type;
    using index_type = std::ptrdiff_t;
    using pointer = value_type*;
    using reference = value_type&;
    using iterator = pointer;
    using reverse_iterator = std::reverse_iterator<iterator>;
    using const_pointer = const value_type*;
    using const_reference = const value_type&;
    using const_iterator = const_pointer;
    using const_reverse_iterator = std::reverse_iterator<const_iterator>;
    
    // Storage layout is compatible with Eigen: Eigen/src/Core/DenseStorage.h

    pointer lparray = nullptr;
    int32_t rows_ = 0; // INT32
    int32_t cols_ = 0; // INT32
};

} // namespace detail

struct xlmulti : detail::xlmulti_base<
    detail::variant_base<
        xlnum,
        xlstr,
        xlbool,
        xlerr,
        xlint,
        xlsref,
        xlref,
        xlmulti,
        xlflow,
        xlbigdata,
        xlmissing,
        xlnil>>
{
    xlmulti(unsigned rows, unsigned cols)
    {   
        std::size_t n = rows * cols;
        if (rows != 0 && n / rows != cols)
            throw std::length_error("xlmulti too long");

        lparray = alloc().allocate(n);
        std::uninitialized_fill_n(lparray, n, xlnil());
        rows_ = static_cast<int32_t>(rows);
        cols_ = static_cast<int32_t>(cols);
    }

    xlmulti(const xlmulti& other)
    {
        std::size_t n = static_cast<std::size_t>(other.rows_ * other.cols_);
        lparray = alloc().allocate(n);
        std::uninitialized_copy_n(other.lparray, n, lparray);
        rows_ = other.rows_;
        cols_ = other.cols_;
    }

    xlmulti(xlmulti&& other) noexcept
    {
        if (this != &other) {
            lparray = std::exchange(other.lparray, nullptr);
            rows_ = std::exchange(other.rows_, 0);
            cols_ = std::exchange(other.cols_, 0);
        }
    }

    xlmulti(std::initializer_list<std::initializer_list<value_type>> list)
    {
        rows_ = static_cast<int32_t>(list.size());
        cols_ = 0;

        for (const auto& row : list) {
            if (row.size() > static_cast<std::size_t>(cols_)) {
                cols_ = static_cast<int32_t>(row.size());
            }
        }

        std::size_t n = static_cast<std::size_t>(rows_ * cols_);
        lparray = alloc().allocate(n);
        pointer current = lparray;

        for (const auto& row : list) {
            current = std::uninitialized_copy(row.begin(), row.end(), current);
            auto cols = row.size();
            if (cols < static_cast<std::size_t>(cols_)) {
                current = std::uninitialized_fill_n(current, cols_ - cols, xlnil());
            }
        }
    }

    ~xlmulti() noexcept
    {
        if (lparray) {
            std::destroy_n(lparray, size());
            alloc().deallocate(lparray, size());
        }
    }

    xlmulti& operator=(const xlmulti& other)
    {
        this->~xlmulti();
        std::size_t n = static_cast<std::size_t>(other.rows_ * other.cols_);
        lparray = alloc().allocate(n);
        std::uninitialized_copy_n(other.lparray, n, lparray);
        rows_ = other.rows_;
        cols_ = other.cols_;
        return *this;
    }

    xlmulti& operator=(xlmulti&& other) noexcept
    {
        if (this != &other) {
            this->~xlmulti();
            lparray = std::exchange(other.lparray, nullptr);
            rows_ = std::exchange(other.rows_, 0);
            cols_ = std::exchange(other.cols_, 0);
        }
        return *this;
    }
    
    inline bool empty() const noexcept {
        return size() == 0;
    }

    inline pointer data() const noexcept {
        return lparray;
    }

    inline std::size_t size() const noexcept {
        return static_cast<std::size_t>(rows_ * cols_);
    }

    inline reference operator[](std::size_t n) noexcept {
        return *std::next(lparray, static_cast<std::ptrdiff_t>(n));
    }

    inline reference operator()(unsigned i, unsigned j) noexcept {
        auto n = static_cast<std::size_t>(i * cols_ + j);
        return *std::next(lparray, static_cast<std::ptrdiff_t>(n));
    }

    inline reference at(std::size_t n) {
        if (n > size())
            throw std::out_of_range("invalid xlmulti subscript");
        return operator[](n);
    }

    inline reference at(unsigned i, unsigned j) {
        auto n = static_cast<std::size_t>(i * cols_ + j);
        if (n > size())
            throw std::out_of_range("invalid xlmulti subscript");
        return operator()(i, j);
    }

    inline const_iterator begin() const noexcept
        { return lparray; }

    inline const_iterator cbegin() const noexcept
        { return lparray; }

    inline const_iterator end() const noexcept
        { return std::next(lparray, static_cast<std::ptrdiff_t>(size())); }
    
    inline const_iterator cend() const noexcept
        { return std::next(lparray, static_cast<std::ptrdiff_t>(size())); }

    inline iterator begin() noexcept
        { return lparray; }
    
    inline iterator end() noexcept
        { return std::next(lparray, static_cast<std::ptrdiff_t>(size())); }

    inline const_reverse_iterator rbegin() const noexcept
        { return const_reverse_iterator(end()); }
    
    inline const_reverse_iterator crbegin() const noexcept
        { return const_reverse_iterator(end()); }
    
    inline const_reverse_iterator rend() const noexcept
        { return const_reverse_iterator(begin()); }
    
    inline const_reverse_iterator crend() const noexcept
        { return const_reverse_iterator(begin()); }
    
    inline reverse_iterator rbegin() noexcept
        { return reverse_iterator(end()); }
    
    inline reverse_iterator rend() noexcept
        { return reverse_iterator(begin()); }
};

} // namespace xll