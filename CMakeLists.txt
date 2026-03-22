cmake_minimum_required(VERSION 3.22)
project(OpenSheetET VERSION 1.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)
set(CMAKE_INCLUDE_CURRENT_DIR ON)

# ── Find Qt6 ──────────────────────────────────────────────────────────────
# jurplel/install-qt-action sets QT_ROOT_DIR env var.
# Pass it as CMAKE_PREFIX_PATH so find_package(Qt6) works correctly.
find_package(Qt6 REQUIRED COMPONENTS
    Core         # QObject, QString, QVariant, QHash, QVector, QJsonDocument
    Gui          # QColor, QFont, QBrush, QPainter, QPixmap, QIcon, QKeyEvent
    Widgets      # QMainWindow, QTableView, QToolButton, QDialog, QMenu
    PrintSupport # QPrinter, QPrintDialog
)

message(STATUS "Qt6 version  : ${Qt6_VERSION}")
message(STATUS "Qt6Core dir  : ${Qt6Core_INCLUDE_DIRS}")
message(STATUS "Qt6Gui  dir  : ${Qt6Gui_INCLUDE_DIRS}")

# ── Platform-specific app icon ──────────────────────────────────────────────
if(WIN32)
    set(APP_ICON_RC "${CMAKE_CURRENT_SOURCE_DIR}/resources/app.rc")
endif()

# ════════════════════════════════════════════════════════════════════════════
#  OSCore  — data model + formula engine  (static lib)
#  Headers include: QColor QFont QBrush QPainter QPixmap (Qt6::Gui)
#                   QObject QString QVariant QHash     (Qt6::Core)
# ════════════════════════════════════════════════════════════════════════════
add_library(OSCore STATIC
    core/Cell.cpp
    core/Sheet.cpp
    core/Workbook.cpp
    core/FormulaEngine.cpp
    core/CsvLoader.cpp
    core/UndoStack.cpp
)

# Belt-and-suspenders: both target_link_libraries (INTERFACE propagation)
# AND explicit target_include_directories (for MSVC generator quirks).
target_include_directories(OSCore
    PUBLIC
        core
        ${Qt6Core_INCLUDE_DIRS}
        ${Qt6Gui_INCLUDE_DIRS}
)

target_link_libraries(OSCore
    PUBLIC
        Qt6::Core
        Qt6::Gui
)

# ════════════════════════════════════════════════════════════════════════════
#  OSUI  — ribbon + all dialogs  (static lib)
# ════════════════════════════════════════════════════════════════════════════
add_library(OSUI STATIC
    ui/RibbonWidget.cpp
    ui/ColorButton.cpp
    ui/FormatCellsDialog.cpp
    ui/FindReplaceDialog.cpp
    ui/FunctionWizard.cpp
    ui/CondFmtDialog.cpp
    ui/CellStyleDialog.cpp
)

target_include_directories(OSUI
    PUBLIC
        ui
        ${Qt6Widgets_INCLUDE_DIRS}
)

target_link_libraries(OSUI
    PUBLIC
        Qt6::Widgets
        OSCore           # brings Qt6::Core + Qt6::Gui transitively
)

# ════════════════════════════════════════════════════════════════════════════
#  OpenSheetET  — main executable
# ════════════════════════════════════════════════════════════════════════════
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
    add_executable(OpenSheetET WIN32 ${APP_SOURCES} ${APP_ICON_RC})
else()
    add_executable(OpenSheetET ${APP_SOURCES})
endif()

target_include_directories(OpenSheetET
    PRIVATE
        app
        ui
        core
        ${Qt6Core_INCLUDE_DIRS}
        ${Qt6Gui_INCLUDE_DIRS}
        ${Qt6Widgets_INCLUDE_DIRS}
)

target_link_libraries(OpenSheetET
    PRIVATE
        OSCore
        OSUI
        Qt6::Core
        Qt6::Gui
        Qt6::Widgets
        Qt6::PrintSupport
)

# ── Install ──────────────────────────────────────────────────────────────────
install(TARGETS OpenSheetET DESTINATION bin)
