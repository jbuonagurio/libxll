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

#include <boost/core/alloc_construct.hpp>

#include <iterator>
#include <stdexcept>

namespace xll {

struct xlref12 {
    int32_t rwFirst = 0; // INT32
    int32_t rwLast = 0; // INT32
    int32_t colFirst = 0; // INT32
    int32_t colLast = 0; // INT32
};

struct xlmref12 {
    uint16_t count = 1; // WORD
    xlref12 reftbl[1]; // actually reftbl[count]
};

//
// XLOPER12 internal value types
//

struct xlnum
{
    using xltype = std::integral_constant<XLTYPE, xltypeNum>;
    xlnum() = default;
    xlnum(double x) : num(x) {}

    explicit operator double() const {
        return num;
    }

    friend bool operator==(const xlnum& lhs, const xlnum& rhs) {
        return lhs.num == rhs.num;
    }

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
    xlbool(bool x) : xbool(x) {}
    
    explicit operator bool() const {
        return xbool != 0 ? true : false;
    }

    friend bool operator==(const xlbool& lhs, const xlbool& rhs) {
        return lhs.xbool == rhs.xbool;
    }

    int32_t xbool; // BOOL (INT32)
};

struct xlerr
{
    using xltype = std::integral_constant<XLTYPE, xltypeErr>;
    xlerr() = default;
    xlerr(error::excel_error x) : err(x) {}

    explicit operator std::error_code() const {
        return error::make_error_code(err);
    }

    explicit operator error::excel_error() const {
        return err;
    }

    friend bool operator==(const xlerr& lhs, const xlerr& rhs) {
        return lhs.err == rhs.err;
    }

    error::excel_error err;
};

struct xlint
{
    using xltype = std::integral_constant<XLTYPE, xltypeInt>;
    xlint() = default;
    xlint(int32_t i) : w(i) {}
    
    explicit operator int32_t() const {
        return w;
    }

    friend bool operator==(const xlint& lhs, const xlint& rhs) {
        return lhs.w == rhs.w;
    }

    int32_t w;
};

struct xlsref
{
    using xltype = std::integral_constant<XLTYPE, xltypeSRef>;
    uint16_t count = 1; // WORD, always = 1
    xlref12 ref;
};

struct xlref
{
    using xltype = std::integral_constant<XLTYPE, xltypeRef>;
    xlmref12 *lpmref = nullptr;
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
    using const_pointer = const value_type*;
    using const_reference = const value_type&;

    pointer lparray = nullptr;
    int32_t rows = 0; // INT32
    int32_t columns = 0; // INT32
};

// forward declaration
template <class... Ts>
struct variant;

} // namespace detail

struct xlmulti : detail::xlmulti_base<
    detail::variant<
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
    xlmulti(unsigned i, unsigned j)
    {
        lparray = alloc().allocate(i * j);
        detail::construct_guard<allocator_type> guard(alloc(), lparray, i * j);
        boost::alloc_construct_n(alloc(), lparray, i * j);
        guard.release();
    }

    ~xlmulti() noexcept
    {
        boost::alloc_destroy_n(alloc(), lparray, size());
        alloc().deallocate(lparray, size());
    }

    bool empty() const noexcept {
        return size() == 0;
    }

    pointer data() const noexcept {
        return lparray;
    }

    std::size_t size() const noexcept {
        return static_cast<std::size_t>(rows * columns);
    }

    reference operator[](std::size_t i) noexcept {
        return *std::next(lparray, i);
    }

    reference operator()(unsigned i, unsigned j) noexcept {
        auto n = static_cast<std::size_t>(i * columns + j);
        return operator[](n);
    }

    reference at(std::size_t n) {
        if (n > size())
            throw std::out_of_range("invalid xlmulti<T> subscript");
        return operator[](n);
    }

    reference at(unsigned i, unsigned j) {
        auto n = static_cast<std::size_t>(i * columns + j);
        if (n > size())
            throw std::out_of_range("invalid xlmulti<T> subscript");
        return operator[](n);
    }
};

} // namespace xll