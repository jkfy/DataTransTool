cmake_minimum_required(VERSION 3.21)

project(DataTransTool VERSION 0.1.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC_SEARCH_PATHS "${CMAKE_CURRENT_SOURCE_DIR}")

find_package(Qt6 6.5 REQUIRED COMPONENTS Widgets)

add_executable(DataTransTool WIN32
    main.cpp
    MainWindow.cpp
    MainWindow.h
    MainWindow.ui
    app.qrc
    BackendClient.cpp
    BackendClient.h
)

target_link_libraries(DataTransTool PRIVATE Qt6::Widgets)

if (MSVC)
    target_compile_options(DataTransTool PRIVATE /utf-8)
endif()

# 开发时把模板复制到可执行文件旁边，方便直接运行 build 目录里的 exe。
add_custom_command(TARGET DataTransTool POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_directory
        "${CMAKE_CURRENT_SOURCE_DIR}/../templates"
        "$<TARGET_FILE_DIR:DataTransTool>/templates"
    COMMENT "Copying templates next to DataTransTool.exe"
)

install(TARGETS DataTransTool RUNTIME DESTINATION .)

option(DATATRANS_BUILD_SMOKE_TEST "Build BackendClient smoke test" OFF)
if (DATATRANS_BUILD_SMOKE_TEST)
    add_executable(backend_client_smoke
        tests/backend_client_smoke.cpp
        BackendClient.cpp
        BackendClient.h
    )
    target_include_directories(backend_client_smoke PRIVATE
        "${CMAKE_CURRENT_SOURCE_DIR}"
    )
    target_link_libraries(backend_client_smoke PRIVATE Qt6::Core)
endif()

option(DATATRANS_BUILD_UI_TEST "Build MainWindow UI smoke test" OFF)
if (DATATRANS_BUILD_UI_TEST)
    find_package(Qt6 6.5 REQUIRED COMPONENTS Test)
    add_executable(ui_smoke_test
        tests/ui_smoke_test.cpp
        MainWindow.cpp
        MainWindow.h
        MainWindow.ui
        app.qrc
        BackendClient.cpp
        BackendClient.h
    )
    target_include_directories(ui_smoke_test PRIVATE
        "${CMAKE_CURRENT_SOURCE_DIR}"
    )
    target_link_libraries(ui_smoke_test PRIVATE Qt6::Widgets Qt6::Test)
    if (MSVC)
        target_compile_options(ui_smoke_test PRIVATE /utf-8)
    endif()
endif()
