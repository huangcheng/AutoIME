cmake_minimum_required(VERSION 3.28)
project(AutoIME)

set(CMAKE_CXX_STANDARD 26)

#set(CMAKE_WINDOWS_EXPORT_ALL_SYMBOLS ON)

add_definitions(-DUNICODE -D_UNICODE)

add_library(ime SHARED library.cpp)

set_target_properties(ime PROPERTIES
        RUNTIME_OUTPUT_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}/lib
)
