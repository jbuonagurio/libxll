//  Copyright 2020 John Buonagurio
//
//  Distributed under the Boost Software License, Version 1.0.
//
//  See accompanying file LICENSE_1_0.txt or copy at
//  http://www.boost.org/LICENSE_1_0.txt

#pragma once

// Boost Dependencies:
// - Boost.Assert
// - Boost.Config
// - Boost.MP11
// - Boost.NoWide
// - Boost.Predef
// - Boost.WinAPI
// - Boost.uBLAS

#include <boost/config.hpp>
#include <boost/predef.h>

#ifndef XLL_EXPORT
#define XLL_EXPORT extern "C" BOOST_SYMBOL_EXPORT
#endif

#if defined(NDEBUG) && !defined(XLL_DISABLE_ASSERTS)
#define XLL_DISABLE_ASSERTS
#endif

// Enable OutputDebugStringA logging using spdlog.
//#ifndef XLL_USE_SPDLOG
//#define XLL_USE_SPDLOG
//#endif

// Custom entry point may be used for an HPC cluster container DLL.
#ifndef EXCEL12ENTRYPT
#define EXCEL12ENTRYPT "MdCallBack12"
#endif
