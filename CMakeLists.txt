cmake_minimum_required(VERSION 3.20)
project(libtorrent_downloader LANGUAGES CXX)

# Use vcpkg toolchain if available
if(NOT DEFINED CMAKE_TOOLCHAIN_FILE AND EXISTS "$ENV{VCPKG_ROOT}/scripts/buildsystems/vcpkg.cmake")
  set(CMAKE_TOOLCHAIN_FILE "$ENV{VCPKG_ROOT}/scripts/buildsystems/vcpkg.cmake" CACHE STRING "Vcpkg toolchain file")
endif()

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

add_executable(libtorrent_downloader src/main.cpp)

# Prefer static runtime on Windows for easier distribution
if(WIN32)
  foreach(flag_var CMAKE_C_FLAGS_RELEASE CMAKE_C_FLAGS_DEBUG CMAKE_CXX_FLAGS_RELEASE CMAKE_CXX_FLAGS_DEBUG)
    if(${flag_var} MATCHES "/MDd?")
      string(REGEX REPLACE "/MDd?" "/MT" ${flag_var} "${${flag_var}}")
    endif()
  endforeach()
endif()

find_package(libtorrent-rasterbar CONFIG REQUIRED)

# Figure out the correct imported target name across ecosystems
set(LT_TARGET "")
if(TARGET torrent-rasterbar)
  set(LT_TARGET torrent-rasterbar)
elseif(TARGET libtorrent-rasterbar)
  set(LT_TARGET libtorrent-rasterbar)
elseif(TARGET LibtorrentRasterbar::torrent-rasterbar)
  set(LT_TARGET LibtorrentRasterbar::torrent-rasterbar)
elseif(TARGET libtorrent-rasterbar::torrent-rasterbar)
  set(LT_TARGET libtorrent-rasterbar::torrent-rasterbar)
else()
  message(FATAL_ERROR "libtorrent imported target not found. Known names: torrent-rasterbar, libtorrent-rasterbar, *::torrent-rasterbar")
endif()

# On Windows, link with Iphlpapi if needed by libtorrent
if(WIN32)
  target_link_libraries(libtorrent_downloader PRIVATE ${LT_TARGET} Iphlpapi)
else()
  target_link_libraries(libtorrent_downloader PRIVATE ${LT_TARGET})
endif()

# Enable more warnings
if(MSVC)
  target_compile_options(libtorrent_downloader PRIVATE /W4 /WX-)
else()
  target_compile_options(libtorrent_downloader PRIVATE -Wall -Wextra -Wpedantic)
endif()
