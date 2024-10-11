# Copyright 2024 John Buonagurio
#
# Distributed under the Boost Software License, Version 1.0.
# (See accompanying file LICENSE_1_0.txt
# or copy at http://www.boost.org/LICENSE_1_0.txt)

cmake_minimum_required(VERSION 3.15)

# Enable CMAKE_MSVC_RUNTIME_LIBRARY.
cmake_policy(SET CMP0091 NEW)

if(DEFINED ENV{VCPKG_ROOT} AND NOT DEFINED CMAKE_TOOLCHAIN_FILE)
  file(TO_CMAKE_PATH "$ENV{VCPKG_ROOT}" VCPKG_ROOT)
  set(CMAKE_TOOLCHAIN_FILE "${VCPKG_ROOT}/scripts/buildsystems/vcpkg.cmake" CACHE STRING "")
  unset(VCPKG_ROOT)
endif()

if(DEFINED ENV{VCPKG_DEFAULT_TRIPLET} AND NOT DEFINED VCPKG_TARGET_TRIPLET)
  set(VCPKG_TARGET_TRIPLET "$ENV{VCPKG_DEFAULT_TRIPLET}" CACHE STRING "")
else()
  # Use default triplet with static CRT and library linkage if not specified.
  if(CMAKE_CL_64)
    set(VCPKG_TARGET_TRIPLET "x64-windows-static" CACHE STRING "")
  elseif(CMAKE_CL_32)
    set(VCPKG_TARGET_TRIPLET "x86-windows-static" CACHE STRING "")
  endif()
endif()

set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>" CACHE STRING "")

message(STATUS "CMAKE_TOOLCHAIN_FILE: ${CMAKE_TOOLCHAIN_FILE}")
message(STATUS "VCPKG_TARGET_TRIPLET: ${VCPKG_TARGET_TRIPLET}")
message(STATUS "CMAKE_MSVC_RUNTIME_LIBRARY: ${CMAKE_MSVC_RUNTIME_LIBRARY}")

if(APPLE)
  # Excel for Mac requires Universal Binary.
  set(CMAKE_OSX_ARCHITECTURES "arm64;x86_64" CACHE STRING "" FORCE)
endif()

#-------------------------------------------------------------------------------
# Project
#-------------------------------------------------------------------------------

project(xll LANGUAGES CXX VERSION 1.0.0)

include(CTest)
include(GNUInstallDirs)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(BUILD_SHARED_LIBS ON)

# Common Options
option(BUILD_EXAMPLES "Build examples" OFF)
option(BUILD_TESTING "Build tests" OFF)

# Dependencies
find_package(Boost REQUIRED)

#-------------------------------------------------------------------------------
# Interface Target
#-------------------------------------------------------------------------------

add_library(xll INTERFACE)
add_library(xll::xll ALIAS xll)

target_compile_features(xll INTERFACE cxx_std_17)

target_include_directories(xll INTERFACE
  "$<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>"
  "$<INSTALL_INTERFACE:${CMAKE_INSTALL_INCLUDEDIR}>")

target_link_libraries(xll INTERFACE Boost::boost)

#-------------------------------------------------------------------------------
# Tests
#-------------------------------------------------------------------------------

if(BUILD_TESTING)
  add_executable(test_pstring ${CMAKE_CURRENT_SOURCE_DIR}/test/test_pstring.cpp)
  add_executable(test_register ${CMAKE_CURRENT_SOURCE_DIR}/test/test_register.cpp)
  add_executable(test_xloper ${CMAKE_CURRENT_SOURCE_DIR}/test/test_xloper.cpp)
  
  target_link_libraries(test_pstring PRIVATE xll)
  target_link_libraries(test_register PRIVATE xll)
  target_link_libraries(test_xloper PRIVATE xll)

  if(MSVC)
    set_target_properties(test_pstring test_register test_xloper PROPERTIES COMPILE_FLAGS "/W4")
  elseif(CMAKE_CXX_COMPILER_ID MATCHES "Clang")
    set_target_properties(test_pstring test_register test_xloper PROPERTIES COMPILE_FLAGS "-Wall")
  endif()

  add_test(test_pstring test_pstring)
  add_test(test_register test_register)
  add_test(test_xloper test_xloper)

  set_tests_properties(test_pstring test_register test_xloper PROPERTIES FAIL_REGULAR_EXPRESSION "[^a-z]Failed")
endif()

#-------------------------------------------------------------------------------
# Examples
#-------------------------------------------------------------------------------

if(BUILD_EXAMPLES)
  # Minimal
  add_library(minimal SHARED
    ${CMAKE_CURRENT_SOURCE_DIR}/src/examples/minimal/addin.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/src/examples/minimal/addin.def)
  target_link_libraries(minimal PRIVATE xll)
  set_target_properties(minimal PROPERTIES OUTPUT_NAME "minimal" PREFIX "" SUFFIX ".xll")
  
  # Generic
  add_library(generic SHARED
    ${CMAKE_CURRENT_SOURCE_DIR}/src/examples/generic/addin.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/src/examples/generic/addin.def)
  target_link_libraries(generic PRIVATE xll)
  set_target_properties(generic PROPERTIES OUTPUT_NAME "generic" PREFIX "" SUFFIX ".xll")

  # Geodesic
  find_package(GeographicLib CONFIG)
  if(GeographicLib_FOUND)
    add_library(geodesic SHARED
      ${CMAKE_CURRENT_SOURCE_DIR}/src/examples/geodesic/addin.cpp
      ${CMAKE_CURRENT_SOURCE_DIR}/src/examples/geodesic/addin.def)
    target_link_libraries(geodesic PRIVATE xll ${GeographicLib_LIBRARIES})
    set_target_properties(geodesic PROPERTIES OUTPUT_NAME "geodesic" PREFIX "" SUFFIX ".xll")
  endif()
endif()

#-------------------------------------------------------------------------------
# Install
#-------------------------------------------------------------------------------

set(XLL_CMAKE_DIR ${CMAKE_INSTALL_LIBDIR}/cmake/xll)

install(DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}/include/xll
        DESTINATION ${CMAKE_INSTALL_INCLUDEDIR})

install(TARGETS xll EXPORT xll-targets
        LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
        ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
        RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR})

install(EXPORT xll-targets
        FILE xll-targets.cmake
        NAMESPACE xll::
        DESTINATION ${XLL_CMAKE_DIR})

include(CMakePackageConfigHelpers)

configure_package_config_file(
  ${CMAKE_CURRENT_SOURCE_DIR}/cmake/xll-config.cmake.in
  ${CMAKE_CURRENT_BINARY_DIR}/xll-config.cmake
  INSTALL_DESTINATION ${XLL_CMAKE_DIR})
  
install(FILES ${CMAKE_CURRENT_BINARY_DIR}/xll-config.cmake
        DESTINATION ${XLL_CMAKE_DIR})

write_basic_package_version_file(
  ${CMAKE_CURRENT_BINARY_DIR}/xll-config-version.cmake
  VERSION ${PROJECT_VERSION}
  COMPATIBILITY AnyNewerVersion
  ARCH_INDEPENDENT)

install(FILES ${CMAKE_CURRENT_BINARY_DIR}/xll-config-version.cmake
        DESTINATION ${XLL_CMAKE_DIR})
