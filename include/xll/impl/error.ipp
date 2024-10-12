//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <string>

namespace xll {
namespace error {
namespace detail {

struct excel_category : public std::error_category
{
    const char* name() const noexcept
    {
        return "Excel";
    }
    
    std::string message(int value) const
    {
        if (value == error::xlerrNull)
            return "#NULL!";
        if (value == error::xlerrDiv0)
            return "#DIV/0!";
        if (value == error::xlerrValue)
            return "#VALUE!";
        if (value == error::xlerrRef)
            return "#REF!";
        if (value == error::xlerrName)
            return "#NAME?";
        if (value == error::xlerrNum)
            return "#NUM!";
        if (value == error::xlerrNA)
            return "#N/A";
        if (value == error::xlerrGettingData)
            return "#GETTING_DATA";
        if (value == error::xlerrSpill)
            return "#SPILL!";
        if (value == error::xlerrConnect)
            return "#CONNECT!";
        if (value == error::xlerrBlocked)
            return "#BLOCKED!";
        if (value == error::xlerrUnknown)
            return "#UNKNOWN!";
        if (value == error::xlerrField)
            return "#FIELD!";
        if (value == error::xlerrCalc)
            return "#CALC!";
        
        return {};
    }
};

} // namespace detail

const std::error_category& get_excel_category()
{
    static detail::excel_category instance;
    return instance;
}

} // namespace error
} // namespace xll
