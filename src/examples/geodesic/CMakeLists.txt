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

project(geodesic LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Set default build type to Release.
set_property(CACHE CMAKE_BUILD_TYPE PROPERTY VALUE "Release")

# Always use statically-linked CRT.
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

# Always build a shared library.
set(BUILD_SHARED_LIBS ON)

# Set default install prefix to local Excel add-ins directory.
if(CMAKE_INSTALL_PREFIX_INITIALIZED_TO_DEFAULT)
  set(CMAKE_INSTALL_PREFIX "$ENV{USERPROFILE}/AppData/Roaming/Microsoft/AddIns/"
      CACHE PATH "" FORCE)
endif()

set(CMAKE_EXE_LINKER_FLAGS    "${CMAKE_EXE_LINKER_FLAGS} /MANIFEST:NO")
set(CMAKE_MODULE_LINKER_FLAGS "${CMAKE_MODULE_LINKER_FLAGS} /MANIFEST:NO")
set(CMAKE_SHARED_LINKER_FLAGS "${CMAKE_SHARED_LINKER_FLAGS} /MANIFEST:NO")

# Dependencies
set(Boost_USE_STATIC_LIBS ON)
set(Boost_USE_MULTITHREADED ON)
set(Boost_USE_STATIC_RUNTIME ON)
set(Boost_ALL_DYN_LINK OFF)

find_package(GeographicLib CONFIG REQUIRED)

add_subdirectory(xll)

# Target
add_library(addin SHARED
  ${CMAKE_CURRENT_SOURCE_DIR}/addin.cpp
  ${CMAKE_CURRENT_SOURCE_DIR}/addin.def)

target_compile_options(addin PRIVATE /W4)

target_link_libraries(addin PRIVATE xll GeographicLib::GeographicLib_STATIC)

set_target_properties(addin PROPERTIES
  OUTPUT_NAME "geodesic"
  SUFFIX ".xll")

install(TARGETS addin
        RUNTIME DESTINATION ${CMAKE_INSTALL_PREFIX})
