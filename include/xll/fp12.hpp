//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

#include <xll/config.hpp>

#include <xll/detail/array_adaptor.hpp>

#include <boost/numeric/ublas/matrix.hpp>
#include <boost/numeric/ublas/storage.hpp>

/**
 * \file fp12.hpp
 * FP12 structure from XLCALL.H adapted to C++.
 */

namespace xll {

/// Original C-style variable length struct.

struct fp12
{
    int32_t rows; // INT32
    int32_t columns; // INT32
    double array[1];
};

/// uBLAS dense matrix adapted to the storage layout of the FP12 structure from
/// XLCALL.H: INT32 rows, INT32 columns, double array[1].
///
/// A::size_type size1_;
/// A::size_type size2_;
/// A data_;

template<std::size_t N>
using static_fp12_storage = detail::static_array<double, N, int32_t>;

template<std::size_t N = 1>
using static_fp12 = boost::numeric::ublas::matrix<double,
    boost::numeric::ublas::row_major,
    static_fp12_storage<N>>;

} // namespace xll
