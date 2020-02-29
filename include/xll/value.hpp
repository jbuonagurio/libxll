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

namespace xll {
namespace detail {

template <class... Ts>
struct variant;

} // namespace detail

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

template <class Tag>
struct value;

template <>
struct value<tag::xlnum> : tag::xlnum
{
    value() = default;
    value(double x) : num(x) {}

    explicit operator double() const {
        return num;
    }

    friend bool operator==(const value& lhs, const value& rhs) {
        return lhs.num == rhs.num;
    }

    double num;
};

template <>
struct value<tag::xlstr> : public wpstring
{
    using wpstring::wpstring;
};

template <>
struct value<tag::xlbool> : tag::xlbool
{
    value() = default;
    value(bool x) : xbool(x) {}
    
    explicit operator bool() const {
        return xbool != 0 ? true : false;
    }

    friend bool operator==(const value& lhs, const value& rhs) {
        return lhs.xbool == rhs.xbool;
    }

    int32_t xbool; // BOOL (INT32)
};

template <>
struct value<tag::xlerr> : tag::xlerr
{
    value() = default;
    value(error::excel_error x) : err(x) {}

    explicit operator std::error_code() const {
        return error::make_error_code(err);
    }

    explicit operator error::excel_error() const {
        return err;
    }

    friend bool operator==(const value& lhs, const value& rhs) {
        return lhs.err == rhs.err;
    }

    error::excel_error err;
};

template <>
struct value<tag::xlint> : tag::xlint
{
    value() = default;
    value(int32_t i) : w(i) {}
    
    explicit operator int32_t() const {
        return w;
    }

    friend bool operator==(const value& lhs, const value& rhs) {
        return lhs.w == rhs.w;
    }

    int32_t w;
};

template <>
struct value<tag::xlsref>: tag::xlsref
{
    uint16_t count = 1; // WORD, always = 1
    xlref12 ref;
};

template <>
struct value<tag::xlref> : tag::xlref
{
    xlmref12 *lpmref = nullptr;
    uintptr_t idSheet; // DWORD_PTR (ULONG_PTR)
};

template <>
struct value<tag::xlmissing> : tag::xlmissing
{};

template <>
struct value<tag::xlnil> : tag::xlnil
{};

template <>
struct value<tag::xlflow> : tag::xlflow
{
    union {
        int level; // xlflowRestart
        int tbctrl; // xlflowPause
        uintptr_t idSheet; // DWORD_PTR (ULONG_PTR), xlflowGoto
    } valflow;
    int32_t rw; // INT32, xlflowGoto
    int32_t col; // INT32, xlflowGoto
    XLFLOW xlflow; // BYTE
};

template <>
struct value<tag::xlbigdata> : tag::xlbigdata
{
    void *h = nullptr; // BYTE*, HANDLE
    long cbData = 0;
};

template <>
struct value<tag::xlmulti> : tag::xlmulti
{
    detail::variant<
        value<tag::xlnum>,
        value<tag::xlstr>,
        value<tag::xlbool>,
        value<tag::xlerr>,
        value<tag::xlint>,
        value<tag::xlsref>,
        value<tag::xlref>,
        value<tag::xlmulti>,
        value<tag::xlflow>,
        value<tag::xlbigdata>,
        value<tag::xlmissing>,
        value<tag::xlnil>> *lparray = nullptr;
    
    int32_t rows = 0; // INT32
    int32_t columns = 0; // INT32
};

} // namespace xll