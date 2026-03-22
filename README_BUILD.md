# OpenSheet ET — Windows Build Guide

## Requirements

| Tool | Version | Download |
|------|---------|----------|
| **Qt** | 6.5+ (MSVC 2022 x64) | https://www.qt.io/download |
| **CMake** | 3.22+ | https://cmake.org/download |
| **Visual Studio** | 2019 or 2022 (C++ Desktop workload) | https://visualstudio.microsoft.com |
| **NSIS** | 3.x (installer only) | https://nsis.sourceforge.io |

## Quick Build (Windows)

```bat
:: Simple — auto-detects Qt
build.bat

:: Explicit Qt path
build.bat Release "C:\Qt\6.5.3\msvc2022_64"
```

This runs 5 steps:
1. **CMake configure** — generates Visual Studio project
2. **MSBuild compile** — builds all 3 targets (OSCore, OSUI, OpenSheetET.exe)
3. **Stage** — copies .exe to `deploy\`
4. **windeployqt** — copies Qt DLLs/plugins into `deploy\`
5. **NSIS** — packages `deploy\` into `OpenSheetET-1.0.0-Setup.exe`

## Manual Build (GitHub Actions / CI)

```yaml
# See .github/workflows/build-windows.yml
```

## Project Structure

```
OpenSheetET/
├── core/              ← Data model + formula engine (static lib OSCore)
│   ├── Cell.h/cpp        Cell data struct + CellFormat
│   ├── Sheet.h/cpp       Single worksheet (QHash-backed)
│   ├── Workbook.h/cpp    Multi-sheet workbook + undo + file I/O
│   └── FormulaEngine.h/cpp  60+ functions: SUM AVERAGE IF VLOOKUP etc.
│
├── ui/                ← Ribbon + dialogs (static lib OSUI)
│   ├── RibbonWidget.h/cpp   WPS ET-style 7-group ribbon
│   ├── ColorButton.h/cpp    Color picker swatch button
│   ├── FormatCellsDialog    Full format cells dialog
│   ├── FindReplaceDialog    Find & Replace with regex support
│   └── FunctionWizard       60+ function browser
│
├── app/               ← Main executable
│   ├── main.cpp             Splash screen + app entry
│   ├── MainWindow.h/cpp     Top-level window, all signal wiring
│   ├── SpreadsheetView.h/cpp QTableView + SheetModel + CellDelegate
│   ├── FormulaBar.h/cpp     Cell ref + formula input bar
│   ├── SheetTabBar.h/cpp    WPS-style sheet tab bar
│   ├── NotifBar.h/cpp       Green promo notification bar
│   └── StatusBar.h/cpp      Stats + zoom slider
│
├── installer/
│   └── OpenSheetET.nsi      NSIS installer script
│
├── resources/
│   └── app.rc               Windows manifest + icon reference
│
├── build.bat                Windows build script (all-in-one)
├── build.sh                 Linux build script
└── CMakeLists.txt           Root CMake (C++20, Qt6)
```

## Features

- **WPS ET 2018white_dark theme** — pixel-perfect matching of WPS Spreadsheet
- **60+ formulas** — SUM, AVERAGE, IF, IFS, VLOOKUP, DATEDIF, STDEV, TEXTJOIN…
- **Full ribbon** — Clipboard | Font | Alignment | Number | Styles | Cells | Editing
- **Format Conversion** — WPS-exclusive: Convert to Number/Text/Upper/Lower/Proper Case
- **Conditional Formatting** — 10 rule types including Color Scale and Data Bars
- **Cell Styles** — 14 presets (Good/Bad/Heading/Total/Input…)
- **Table Styles** — 8 themes (Light/Medium/Dark in 5 colors)
- **Undo/Redo** — 100-level snapshot history
- **CSV import/export** — full RFC 4180 compliance
- **Find & Replace** — with case-sensitive, whole-cell options
- **Function Wizard** — 65 functions with descriptions
- **Custom CellDelegate** — 0.7px WPS-style grid lines, formula color, error color
- **Splash screen** — animated green splash on startup
- **NSIS installer** — per-user or per-machine, file association (.oset), Start Menu

## Troubleshooting

**CMake can't find Qt:**
```
cmake -B build -DCMAKE_PREFIX_PATH="C:\Qt\6.5.3\msvc2022_64"
```

**C2079 QPainterPath:**
Already fixed — `#include <QPainterPath>` is in RibbonWidget.cpp.

**Missing moc files:**
Run `cmake --build build --target clean` then rebuild.

**linker errors on Windows (undefined `main`):**
The CMakeLists.txt uses `add_executable(OpenSheetET WIN32 ...)` which sets the correct
`WinMain` entry point for Qt apps. Do not use `add_executable(OpenSheetET ...)` without WIN32.
