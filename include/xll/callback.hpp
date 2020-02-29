//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <xll/constants.hpp>
#include <xll/detail/assert.hpp>
#include <xll/detail/callback.hpp>
#include <xll/detail/tagged_union.hpp>
#include <xll/log.hpp>

#include <boost/winapi/dll.hpp>

#include <array>
#include <type_traits>

namespace xll {

/**
 * Handles Excel 12 callbacks.
 * \param[in] xlfn Excel function number.
 * \param[in,out] result Pointer to the XLOPER12 that will hold the result.
 * \param[in,out] opers Array of arguments as pointers to XLOPER12 values.
 * \param[in] count Position of last argument in array.
 * \return Entry point return code (xlret prefix).
 */

template <class R, class V, std::size_t N>
inline XLRET Excel12v(int xlfn, R *result, std::array<V*, N>& opers, std::size_t count = N)
{
    static_assert(N <= 255, "parameter count exceeds Excel 12 limit");
    static_assert(std::is_base_of_v<detail::tagged_union, V>, "invalid operand type");
    static_assert(std::is_base_of_v<detail::tagged_union, R>, "invalid result type");

    auto pfn = detail::MdCallBack12<R, V>();
    XLL_ASSERT(pfn != nullptr);
	if (pfn == nullptr)
        return XLRET::xlretAbort;
    
    int retval = (pfn)(xlfn, static_cast<int>(count), opers.data(), result);
    XLRET rc = static_cast<XLRET>(retval);
    if (rc != XLRET::xlretSuccess) {
        xll::log()->error("Callback failed: {}, xlfn {}", rc, xlfn);
    }
    
    // Set xlBitXLFree on return value from C API. xlFree is called in the destructor.
    if (result) {
        uint32_t xt = result->xltype();
        if (xt == xltypeStr || xt == xltypeRef || xt == xltypeMulti) {
            result->set_flags(xlbitXLFree);
        }
    }
    
    return rc;
}

template <class R>
inline XLRET Excel12(int xlfn, R *result)
{
    std::array<R*, 1> opers{ nullptr };
    return Excel12v(xlfn, result, opers);
}

template <class R, class... Args>
inline XLRET Excel12(int xlfn, R *result, Args*... args)
{
    std::array<detail::tagged_union *, sizeof...(Args)> opers =
        {( static_cast<detail::tagged_union *>(args), ... )};
    
    return Excel12v(xlfn, result, opers);
}

template <class... Args>
inline XLRET Excel12(int xlfn, std::nullptr_t, Args*... args)
{
    std::array<detail::tagged_union *, sizeof...(Args)> opers =
        {( static_cast<detail::tagged_union *>(args), ... )};
    
    return Excel12v<detail::tagged_union>(xlfn, nullptr, opers);
}

} // namespace xll