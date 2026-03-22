cmake_minimum_required(VERSION 3.22)
project(OpenSheetET VERSION 1.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)
set(CMAKE_INCLUDE_CURRENT_DIR ON)

find_package(Qt6 REQUIRED COMPONENTS
    Core Gui Widgets PrintSupport
)

# Platform-specific
if(WIN32)
    set(APP_ICON_RESOURCE_WINDOWS "${CMAKE_CURRENT_SOURCE_DIR}/resources/app.rc")
endif()

# Core library (model/formula engine)
add_library(OSCore STATIC
    core/Cell.cpp
    core/Sheet.cpp
    core/Workbook.cpp
    core/FormulaEngine.cpp
    core/CsvLoader.cpp
    core/UndoStack.cpp
)
target_include_directories(OSCore PUBLIC core)
target_link_libraries(OSCore PUBLIC Qt6::Core)

# UI library (ribbon + dialogs)
add_library(OSUI STATIC
    ui/RibbonWidget.cpp
    ui/ColorButton.cpp
    ui/FormatCellsDialog.cpp
    ui/FindReplaceDialog.cpp
    ui/FunctionWizard.cpp
    ui/CondFmtDialog.cpp
    ui/CellStyleDialog.cpp
)
target_include_directories(OSUI PUBLIC ui)
target_link_libraries(OSUI PUBLIC Qt6::Widgets OSCore)

# Main executable
set(APP_SOURCES
    app/main.cpp
    app/MainWindow.cpp
    app/SpreadsheetView.cpp
    app/CellDelegate.cpp
    app/SheetTabBar.cpp
    app/FormulaBar.cpp
    app/NotifBar.cpp
    app/StatusBar.cpp
)

if(WIN32)
    add_executable(OpenSheetET WIN32 ${APP_SOURCES} ${APP_ICON_RESOURCE_WINDOWS})
else()
    add_executable(OpenSheetET ${APP_SOURCES})
endif()

target_include_directories(OpenSheetET PRIVATE app ui core)
target_link_libraries(OpenSheetET PRIVATE
    OSCore OSUI
    Qt6::Core Qt6::Gui Qt6::Widgets Qt6::PrintSupport
)

# Install rules
install(TARGETS OpenSheetET DESTINATION bin)
if(WIN32)
    # windeployqt handled by build.bat
endif()
