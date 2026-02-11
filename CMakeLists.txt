cmake_minimum_required(VERSION 3.24)
project(MermaidPreview LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# WebView2 via NuGet-style package (manual fetch)
include(FetchContent)
FetchContent_Declare(
    webview2
    URL https://www.nuget.org/api/v2/package/Microsoft.Web.WebView2/1.0.2903.40
    DOWNLOAD_EXTRACT_TIMESTAMP TRUE
)
FetchContent_MakeAvailable(webview2)

# Source files
set(SOURCES
    src/DllMain.cpp
    src/MermaidPreview.cpp
    src/WebView2Manager.cpp
    src/MarkdownParser.cpp
    src/BunRenderer.cpp
)

# Resource file
set(RESOURCES
    resources/MermaidPreview.rc
)

# Create DLL
add_library(${PROJECT_NAME} SHARED ${SOURCES} ${RESOURCES})

# Include directories
target_include_directories(${PROJECT_NAME} PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/include
    ${CMAKE_CURRENT_SOURCE_DIR}/src
    ${webview2_SOURCE_DIR}/build/native/include
)

# Module definition for exports
target_sources(${PROJECT_NAME} PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/exports.def)

# Link WebView2 loader
if(CMAKE_SIZEOF_VOID_P EQUAL 8)
    set(WEBVIEW2_ARCH "x64")
else()
    set(WEBVIEW2_ARCH "x86")
endif()

target_link_libraries(${PROJECT_NAME} PRIVATE
    ${webview2_SOURCE_DIR}/build/native/${WEBVIEW2_ARCH}/WebView2LoaderStatic.lib
    shlwapi.lib
    comctl32.lib
    pathcch.lib
)

# Compiler options
target_compile_definitions(${PROJECT_NAME} PRIVATE
    UNICODE
    _UNICODE
    NOMINMAX
)

# Use static CRT
set_property(TARGET ${PROJECT_NAME} PROPERTY
    MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>"
)

# Disable C++ exceptions warning for /EHsc
target_compile_options(${PROJECT_NAME} PRIVATE /EHsc /W3)

# Output name
set_target_properties(${PROJECT_NAME} PROPERTIES
    OUTPUT_NAME "MermaidPreview"
    SUFFIX ".dll"
)
