# Copyright 2020 John Buonagurio
#
# Distributed under the Boost Software License, Version 1.0.
# (See accompanying file LICENSE_1_0.txt
# or copy at http://www.boost.org/LICENSE_1_0.txt)

cmake_minimum_required(VERSION 3.15)

# Enable MSVC runtime library flags.
cmake_policy(SET CMP0091 NEW)

if(DEFINED ENV{VCPKG_ROOT} AND NOT DEFINED CMAKE_TOOLCHAIN_FILE)
  file(TO_CMAKE_PATH "$ENV{VCPKG_ROOT}" VCPKG_ROOT)
  set(CMAKE_TOOLCHAIN_FILE "${VCPKG_ROOT}/scripts/buildsystems/vcpkg.cmake"
      CACHE STRING "")
  unset(VCPKG_ROOT)
  message(STATUS "CMAKE_TOOLCHAIN_FILE: ${CMAKE_TOOLCHAIN_FILE}")
endif()

if(DEFINED ENV{VCPKG_DEFAULT_TRIPLET} AND NOT DEFINED VCPKG_TARGET_TRIPLET)
  set(VCPKG_TARGET_TRIPLET "$ENV{VCPKG_DEFAULT_TRIPLET}" CACHE STRING "")
  message(STATUS "VCPKG_TARGET_TRIPLET: ${VCPKG_TARGET_TRIPLET}")
else()
  # Use default triplet with static CRT and library linkage if not specified.
  if(CMAKE_CL_64)
    set(VCPKG_TARGET_TRIPLET "x64-windows-static" CACHE STRING "")
  else()
    set(VCPKG_TARGET_TRIPLET "x86-windows-static" CACHE STRING "")
  endif()
endif()

project(xll LANGUAGES CXX)

include(CTest)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Always use static CRT.
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

# Always build shared libraries.
set(BUILD_SHARED_LIBS ON)

# Options
option(XLL_USE_SPDLOG "Enable debug logging wth spdlog" OFF)

option(BUILD_EXAMPLES "Build examples" ON)
option(BUILD_TESTING "Build tests" ON)

# Dependencies
find_package(Boost REQUIRED)

if(XLL_USE_SPDLOG)
  find_package(spdlog CONFIG REQUIRED)
endif()

# Target
add_library(xll INTERFACE)

target_include_directories(xll INTERFACE
    "$<BUILD_INTERFACE:${Boost_INCLUDE_DIRS}>")

target_include_directories(xll INTERFACE
    "$<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>"
    "$<INSTALL_INTERFACE:${CMAKE_INSTALL_INCLUDEDIR}>")

if(XLL_USE_SPDLOG)
  target_link_libraries(xll INTERFACE spdlog::spdlog)
  target_compile_definitions(xll INTERFACE XLL_USE_SPDLOG=1)
endif()

if(BUILD_TESTING)
  add_executable(test_pstring ${CMAKE_CURRENT_SOURCE_DIR}/test/test_pstring.cpp)
  target_link_libraries(test_pstring PRIVATE xll)
  add_test(test_pstring test_pstring)

  add_executable(test_register ${CMAKE_CURRENT_SOURCE_DIR}/test/test_register.cpp)
  target_link_libraries(test_register PRIVATE xll)
  add_test(test_register test_register)  

  add_executable(test_xloper ${CMAKE_CURRENT_SOURCE_DIR}/test/test_xloper.cpp)
  target_link_libraries(test_xloper PRIVATE xll)
  add_test(test_xloper test_xloper)

  set_tests_properties(test_pstring test_register test_xloper PROPERTIES
    FAIL_REGULAR_EXPRESSION "[^a-z]Failed")
endif()

if(BUILD_EXAMPLES)
  add_library(generic SHARED
    ${CMAKE_CURRENT_SOURCE_DIR}/src/examples/generic/addin.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/src/examples/generic/addin.def)
  if(MSVC)
    target_compile_options(generic PRIVATE /W4)
  endif()
  target_link_libraries(generic PRIVATE xll)
  set_target_properties(generic PROPERTIES
    OUTPUT_NAME "generic"
    SUFFIX ".xll")
endif()