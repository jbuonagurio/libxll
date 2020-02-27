//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <system_error>

namespace xll {
namespace error {

//
// Error codes
//
// Used for val.err field of XLOPER and XLOPER12 structures
// when constructing error XLOPERs and XLOPER12s
//

enum excel_error : int {
    xlerrNull = 0,         // #NULL!
    xlerrDiv0 = 7,         // #DIV/0!
    xlerrValue = 15,       // #VALUE!
    xlerrRef = 23,         // #REF!
    xlerrName = 29,        // #NAME?
    xlerrNum = 36,         // #NUM!
    xlerrNA = 42,          // #N/A
    xlerrGettingData = 43  // #GETTING_DATA
};

extern inline const std::error_category& get_excel_category();

static const std::error_category& excel_category = get_excel_category();

} // namespace error
} // namespace xll

namespace std {

template <>
struct is_error_code_enum<xll::error::excel_error> : public true_type {};

} // namespace std

namespace xll {
namespace error {

inline std::error_code make_error_code(excel_error e) {
    return std::error_code(static_cast<int>(e), get_excel_category());
}

} // namespace error
} // namespace xll

#include <xll/impl/error.ipp>
