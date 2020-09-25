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
#include <xll/detail/variant.hpp>
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

template<class R, class V, std::size_t N>
inline int Excel12v(int xlfn, R *result, std::array<V*, N>& opers, std::size_t count = N)
{
    static_assert(N <= 255, "parameter count exceeds Excel 12 limit");
    static_assert(std::is_base_of_v<detail::variant_common_type, V>, "invalid operand type");
    static_assert(std::is_base_of_v<detail::variant_common_type, R>, "invalid result type");

    auto pfn = detail::MdCallBack12<R, V>();
	if (pfn == nullptr)
        return XLRET::xlretFailed;
    
    int rc = (pfn)(xlfn, static_cast<int>(count), opers.data(), result);
    if (rc != XLRET::xlretSuccess) {
        xll::log()->error("Callback failed: xlfn {}, return code {:#06x}", xlfn, rc);
        return rc;
    }
    
    // Set xlBitXLFree on return value from C API. xlFree is called in the destructor.
    if (result) {
        auto *p = reinterpret_cast<detail::variant_opaque_ptr>(result);
        uint32_t xt = p->xltype();
        if (xt == xltypeStr || xt == xltypeRef || xt == xltypeMulti) {
            p->set_flags(xlbitXLFree);
        }
    }
    
    return rc;
}

template<class R>
inline int Excel12(int xlfn, R *result)
{
    std::array<R*, 1> opers{ nullptr };
    return Excel12v(xlfn, result, opers);
}

template<class R, class... Args>
inline int Excel12(int xlfn, R *result, Args*... args)
{
    std::array<detail::variant_common_type *, sizeof...(Args)> opers =
        {( static_cast<detail::variant_common_type *>(args), ... )};
    return Excel12v(xlfn, result, opers);
}

template<class... Args>
inline int Excel12(int xlfn, std::nullptr_t, Args*... args)
{
    std::array<detail::variant_common_type *, sizeof...(Args)> opers =
        {( static_cast<detail::variant_common_type *>(args), ... )};
    return Excel12v<detail::variant_common_type>(xlfn, nullptr, opers);
}

} // namespace xll