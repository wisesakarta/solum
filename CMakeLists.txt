cmake_minimum_required(VERSION 3.16)
project(technical-standard-note VERSION 1.3.9 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# Use static runtime for MSVC
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

option(LEGACY_NOTEPAD_WERROR "Treat warnings as errors" OFF)
option(TECHNICAL_STANDARD_NOTE_BUILD_TESTS "Build lightweight logic tests" ON)

if(MSVC)
    add_compile_options(/W4 /WX /permissive- /utf-8)
    set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} /SUBSYSTEM:WINDOWS /ENTRY:wWinMainCRTStartup")
else()
    add_compile_options(-Wall -Wextra -municode -Wno-error=missing-field-initializers)
    if(LEGACY_NOTEPAD_WERROR)
        add_compile_options(-Werror)
    endif()
    if(CMAKE_CXX_COMPILER_ID STREQUAL "GNU")
        set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} -static -static-libgcc -static-libstdc++ -mwindows")
    else()
        set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} -mwindows")
    endif()
endif()

add_executable(technical-standard-note WIN32
    src/main.cpp
    src/core/crash_diagnostics.cpp
    src/core/file_dialog_filters.cpp
    src/core/globals.cpp
    src/core/list_continuation.cpp
    src/core/session_io.cpp
    src/core/text_codec.cpp
    src/core/versioning.cpp
    src/lang/lang.cpp
    src/modules/theme.cpp
    src/modules/editor.cpp
    src/modules/file.cpp
    src/modules/ui.cpp
    src/modules/background.cpp
    src/modules/command_routing.cpp
    src/modules/dialog.cpp
    src/modules/commands.cpp
    src/modules/icon_actions.cpp
    src/modules/perf_benchmark.cpp
    src/modules/tab_layout.cpp
    src/modules/tab_model_ops.cpp
    src/modules/tab_session_io.cpp
    src/modules/update_check.cpp
    src/modules/menu.cpp
    src/modules/settings.cpp
    src/notepad.rc
)

target_include_directories(technical-standard-note PRIVATE ${CMAKE_SOURCE_DIR}/src ${CMAKE_SOURCE_DIR}/src/modules)
target_compile_definitions(technical-standard-note PRIVATE UNICODE _UNICODE NOMINMAX TECHNICAL_STANDARD_NOTE_SPLIT_COMMAND_MODULES=1)

target_link_libraries(technical-standard-note PRIVATE
    comctl32
    shlwapi
    comdlg32
    shell32
    gdiplus
    msimg32
    user32
    gdi32
    kernel32
    dwmapi
    uxtheme
    winhttp
)

set(TECHNICAL_STANDARD_NOTE_FONT_SOURCE_DIR "${CMAKE_SOURCE_DIR}/assets/fonts")
set(TECHNICAL_STANDARD_NOTE_BUNDLED_FONT_FILES
    "Berkeley Mono Condensed.ttf"
    "Berkeley Mono Condensed Bold.ttf"
    "Berkeley Mono Condensed Oblique.ttf"
    "Berkeley Mono Condensed Bold Oblique.ttf"
)
set(TECHNICAL_STANDARD_NOTE_BUNDLED_FONT_PATHS "")
foreach(font_file IN LISTS TECHNICAL_STANDARD_NOTE_BUNDLED_FONT_FILES)
    set(font_path "${TECHNICAL_STANDARD_NOTE_FONT_SOURCE_DIR}/${font_file}")
    if(EXISTS "${font_path}")
        list(APPEND TECHNICAL_STANDARD_NOTE_BUNDLED_FONT_PATHS "${font_path}")
    endif()
endforeach()
if(TECHNICAL_STANDARD_NOTE_BUNDLED_FONT_PATHS)
    add_custom_command(TARGET technical-standard-note POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E make_directory "$<TARGET_FILE_DIR:technical-standard-note>/fonts"
        COMMAND ${CMAKE_COMMAND} -E copy_if_different ${TECHNICAL_STANDARD_NOTE_BUNDLED_FONT_PATHS} "$<TARGET_FILE_DIR:technical-standard-note>/fonts"
    )
endif()

if(MSVC)
    # Manifest already embedded via src/notepad.rc.
    # Disable linker auto-manifest to avoid duplicate MANIFEST resource (type 1).
    target_link_options(technical-standard-note PRIVATE /MANIFEST:NO)
endif()

if(NOT MSVC)
    target_link_options(technical-standard-note PRIVATE -mwindows -municode)
    if(CMAKE_CXX_COMPILER_ID STREQUAL "GNU")
        target_link_options(technical-standard-note PRIVATE -static)
        target_link_libraries(technical-standard-note PRIVATE -static-libgcc -static-libstdc++)
    endif()
    if(CMAKE_STRIP)
        add_custom_command(TARGET technical-standard-note POST_BUILD COMMAND ${CMAKE_STRIP} --strip-all $<TARGET_FILE:technical-standard-note>)
    endif()
endif()

if(TECHNICAL_STANDARD_NOTE_BUILD_TESTS)
    enable_testing()
    add_executable(technical-standard-note-versioning-tests
        tests/versioning_tests.cpp
        src/core/versioning.cpp
    )
    target_include_directories(technical-standard-note-versioning-tests PRIVATE ${CMAKE_SOURCE_DIR}/src)
    target_compile_definitions(technical-standard-note-versioning-tests PRIVATE UNICODE _UNICODE NOMINMAX)
    add_test(NAME technical-standard-note-versioning-tests COMMAND technical-standard-note-versioning-tests)

    add_executable(technical-standard-note-encoding-tests
        tests/encoding_tests.cpp
        src/core/text_codec.cpp
    )
    target_include_directories(technical-standard-note-encoding-tests PRIVATE ${CMAKE_SOURCE_DIR}/src)
    target_compile_definitions(technical-standard-note-encoding-tests PRIVATE UNICODE _UNICODE NOMINMAX)
    add_test(NAME technical-standard-note-encoding-tests COMMAND technical-standard-note-encoding-tests)

    add_executable(technical-standard-note-session-io-tests
        tests/session_io_tests.cpp
        src/core/session_io.cpp
    )
    target_include_directories(technical-standard-note-session-io-tests PRIVATE ${CMAKE_SOURCE_DIR}/src)
    target_compile_definitions(technical-standard-note-session-io-tests PRIVATE UNICODE _UNICODE NOMINMAX)
    add_test(NAME technical-standard-note-session-io-tests COMMAND technical-standard-note-session-io-tests)

    add_executable(technical-standard-note-tab-session-io-tests
        tests/tab_session_io_tests.cpp
        src/core/session_io.cpp
        src/core/text_codec.cpp
        src/modules/tab_session_io.cpp
    )
    target_include_directories(technical-standard-note-tab-session-io-tests PRIVATE ${CMAKE_SOURCE_DIR}/src ${CMAKE_SOURCE_DIR}/src/modules)
    target_compile_definitions(technical-standard-note-tab-session-io-tests PRIVATE UNICODE _UNICODE NOMINMAX)
    target_link_libraries(technical-standard-note-tab-session-io-tests PRIVATE shlwapi)
    add_test(NAME technical-standard-note-tab-session-io-tests COMMAND technical-standard-note-tab-session-io-tests)

    add_executable(technical-standard-note-tab-model-ops-tests
        tests/tab_model_ops_tests.cpp
        src/modules/tab_model_ops.cpp
    )
    target_include_directories(technical-standard-note-tab-model-ops-tests PRIVATE ${CMAKE_SOURCE_DIR}/src ${CMAKE_SOURCE_DIR}/src/modules)
    target_compile_definitions(technical-standard-note-tab-model-ops-tests PRIVATE UNICODE _UNICODE NOMINMAX)
    add_test(NAME technical-standard-note-tab-model-ops-tests COMMAND technical-standard-note-tab-model-ops-tests)

    add_executable(technical-standard-note-list-continuation-tests
        tests/list_continuation_tests.cpp
        src/core/list_continuation.cpp
    )
    target_include_directories(technical-standard-note-list-continuation-tests PRIVATE ${CMAKE_SOURCE_DIR}/src)
    target_compile_definitions(technical-standard-note-list-continuation-tests PRIVATE UNICODE _UNICODE NOMINMAX)
    add_test(NAME technical-standard-note-list-continuation-tests COMMAND technical-standard-note-list-continuation-tests)
endif()

# Installation
install(TARGETS technical-standard-note DESTINATION .)
install(FILES LICENSE README.md DESTINATION .)
if(TECHNICAL_STANDARD_NOTE_BUNDLED_FONT_PATHS)
    install(FILES ${TECHNICAL_STANDARD_NOTE_BUNDLED_FONT_PATHS} DESTINATION fonts)
endif()

# Installer / packaging (CPack with NSIS)
set(CPACK_GENERATOR "NSIS")
set(CPACK_PACKAGE_NAME "technical-standard-note")
set(CPACK_PACKAGE_VENDOR "Saka Studio & Engineering")
set(CPACK_PACKAGE_VERSION ${PROJECT_VERSION})
if(CMAKE_GENERATOR_PLATFORM)
    set(CPACK_PACKAGE_FILE_NAME "technical-standard-note-${PROJECT_VERSION}-${CMAKE_GENERATOR_PLATFORM}")
else()
    set(CPACK_PACKAGE_FILE_NAME "technical-standard-note-${PROJECT_VERSION}")
endif()
set(CPACK_RESOURCE_FILE_LICENSE "${CMAKE_SOURCE_DIR}/LICENSE")
set(CPACK_NSIS_DISPLAY_NAME "Technical Standard Note")
set(CPACK_NSIS_PACKAGE_NAME "Technical Standard Note")
set(CPACK_NSIS_URL_INFO_ABOUT "https://github.com/wisesakarta/technical-standard-note")
set(CPACK_NSIS_MODIFY_PATH ON)
set(CPACK_NSIS_ENABLE_UNINSTALL_BEFORE_INSTALL ON)
set(CPACK_NSIS_MUI_ICON "${CMAKE_SOURCE_DIR}/src/icon.ico")
set(CPACK_NSIS_MUI_UNIICON "${CMAKE_SOURCE_DIR}/src/icon.ico")
set(CPACK_NSIS_EXECUTABLES_DIRECTORY ".")
set(CPACK_PACKAGE_EXECUTABLES "technical-standard-note.exe" "Technical Standard Note")
set(CPACK_NSIS_INSTALLED_ICON_NAME "technical-standard-note.exe")
set(CPACK_NSIS_MUI_FINISHPAGE_RUN "technical-standard-note.exe")
set(CPACK_NSIS_MENU_LINKS
    "technical-standard-note.exe" "Technical Standard Note"
)

include(CPack)

