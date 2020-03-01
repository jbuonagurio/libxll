//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <xll/value.hpp>
#include <xll/detail/memory.hpp>

#include <type_traits>

/// XLOPER12-compatible aligned storage with heap tracking. If xlbitDLLFree is
/// set, assume the XLOPER and associated xltypeStr, xltypeRef and xltypeMulti
/// values are dynamically allocated and should be destroyed in xlAutoFree12.
/// If xlbitXLFree is set, assume memory is allocated by Excel and should only
/// be destroyed with Excel callback function xlFree.

namespace xll {
namespace detail {

struct tagged_union
{
    using storage_type = std::aligned_union_t<0,
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
        xlnil>;

    storage_type val_;
    uint32_t xltype_; // DWORD

protected:
    tagged_union()
        : tagged_union(xltypeMissing)
    {}

    tagged_union(uint32_t xltype) {
        set_xltype(xltype);
    }

    void set_xltype(uint32_t xltype) {
        if (reinterpret_cast<std::uintptr_t>(this) > stack_base() ||
            reinterpret_cast<std::uintptr_t>(this) < stack_limit()) {
            xltype_ = xltype | xlbitDLLFree; // destroy in xlAutoFree12
        }
        else {
            xltype_ = xltype;
        }
    }

public:
    constexpr void set_flags(uint32_t flags) {
        xltype_ |= (flags & 0xF000);
    }

    constexpr void clear_flags(uint32_t flags) {
        xltype_ &= ~(flags & 0xF000); 
    }

    /// Returns xltype without xlbit flags (xlbitXLFree, xlbitDLLFree).
    constexpr uint32_t xltype() const {
        return (xltype_ & 0x0FFF);
    }

    /// Returns xlbit flags (xlbitXLFree, xlbitDLLFree).
    constexpr uint32_t flags() {
        return (xltype_ & 0xF000);
    }
};

} // namespace detail
} // namespace xll
