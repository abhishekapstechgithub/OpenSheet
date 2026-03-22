# OpenSheet ET

A modern desktop spreadsheet application built with **Qt 6 & C++20**, inspired by WPS ET.

[![Build OpenSheet ET](https://github.com/YOUR_USERNAME/OpenSheetET/actions/workflows/build.yml/badge.svg)](https://github.com/YOUR_USERNAME/OpenSheetET/actions/workflows/build.yml)
[![Latest Release](https://img.shields.io/github/v/release/YOUR_USERNAME/OpenSheetET)](https://github.com/YOUR_USERNAME/OpenSheetET/releases/latest)

---

## 📥 Download

Go to **[Releases](https://github.com/YOUR_USERNAME/OpenSheetET/releases/latest)** and download:
- `OpenSheetET-x.x.x-Setup.exe` — Windows Installer (recommended)
- `OpenSheetET-x.x.x-Portable.zip` — Portable, no install needed

---

## ✨ Features

- **WPS ET-style ribbon** — Clipboard · Font · Alignment · Number · Styles · Cells · Editing
- **60+ formulas** — SUM, AVERAGE, IF, IFS, VLOOKUP, DATEDIF, STDEV, TEXTJOIN…
- **Conditional Formatting** — 10 rule types (color scale, data bars, top/bottom, above/below average…)
- **Cell Styles** — 14 presets (Good/Bad/Heading/Total/Input/Output…)
- **Table Styles** — 8 themes (Light/Medium/Dark in Green, Blue, Orange, Red)
- **Format Conversion** — Convert to Number/Text/Upper/Lower/Proper Case, trim spaces
- **CSV import & export** — RFC 4180 compliant
- **100-level undo/redo** — full snapshot history
- **Find & Replace** — case-sensitive, whole-cell, forward search
- **Function Wizard** — 65 functions with descriptions
- **Splash screen** — animated green WPS-style splash on startup

---

## 🚀 GitHub Actions: How to build automatically

Every time you push code to `main`, GitHub builds the app automatically and produces:
- A **portable ZIP** you can download from the Actions tab
- On version tags (`v1.0.0`), a full **GitHub Release** with installer + portable ZIP

### Step 1 — Push this repository to GitHub

```bash
cd OpenSheetET
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/OpenSheetET.git
git push -u origin main
```

### Step 2 — Watch the build

1. Go to your repo on GitHub
2. Click **Actions** tab
3. See **"Build OpenSheet ET"** running — takes ~8 minutes (Qt download is cached after first run)
4. Click the workflow run → scroll to **Artifacts** → download:
   - `OpenSheetET-Windows-Installer`
   - `OpenSheetET-Windows-Portable`

### Step 3 — Create a release (optional)

**Option A: Tag and push (fully automatic)**
```bash
git tag -a v1.0.0 -m "Release v1.0.0"
git push origin v1.0.0
```
GitHub Actions will automatically:
- Build the app
- Run windeployqt
- Build the NSIS installer
- Create a GitHub Release with all files attached

**Option B: Manual release from GitHub UI**
1. Go to **Actions** → **Release** workflow
2. Click **Run workflow**
3. Enter version number → click **Run**

---

## 🔧 Build locally (Windows)

### Requirements

| Tool | Version | Download |
|------|---------|----------|
| **Qt** | 6.5+ (MSVC 2022 x64) | https://www.qt.io/download |
| **CMake** | 3.22+ | https://cmake.org |
| **Visual Studio** | 2019 or 2022 | https://visualstudio.microsoft.com |
| **NSIS** | 3.x | https://nsis.sourceforge.io |

### One command build

```bat
build.bat
```

Or with explicit Qt path:
```bat
build.bat Release "C:\Qt\6.7.3\msvc2022_64"
```

Output:
- `deploy\OpenSheetET.exe` — the executable with Qt runtime
- `OpenSheetET-1.0.0-Setup.exe` — NSIS installer

---

## 📁 Project Structure

```
OpenSheetET/
├── .github/workflows/
│   ├── build.yml       ← Main CI/CD (runs on every push + releases on tags)
│   └── release.yml     ← Manual release trigger from GitHub UI
│
├── core/               ← Data model + formula engine (static lib)
│   ├── Cell.h            CellFormat: font, colors, alignment, numFmt, border
│   ├── Sheet.h/cpp       QHash-backed worksheet with row/col insert/delete/sort
│   ├── Workbook.h/cpp    Multi-sheet, 100-level undo, CSV/JSON save/load
│   └── FormulaEngine.cpp 60+ functions (SUM,IF,VLOOKUP,STDEV,DATEDIF,TEXTJOIN…)
│
├── ui/                 ← Ribbon + dialogs (static lib)
│   ├── RibbonWidget.cpp  7 groups, WPS ET icons, Format Conversion pill
│   ├── ColorButton.cpp   Swatch-indicator color picker button
│   ├── FormatCellsDialog Full format dialog (font/numfmt/align/colors/decimals)
│   ├── FindReplaceDialog Find & Replace with case/whole-cell options
│   └── FunctionWizard    65-function browser with descriptions
│
├── app/                ← Main executable
│   ├── main.cpp          Splash screen + QApplication setup
│   ├── MainWindow.cpp    All signal wiring, file ops, cell/format operations
│   ├── SpreadsheetView   QTableView + SheetModel + CellDelegate (WPS grid lines)
│   ├── FormulaBar        Cell reference box + fx button + formula input
│   ├── SheetTabBar       WPS-style sheet tabs with drag-reorder + context menu
│   ├── NotifBar          Green dismissible promo bar
│   └── StatusBar         Stats (avg/count/sum) + zoom slider
│
├── installer/
│   └── OpenSheetET.nsi   NSIS script: per-user/per-machine, .oset association
│
├── resources/
│   └── app.rc            Windows manifest + icon
│
├── CMakeLists.txt        C++20, Qt6, 3 targets: OSCore + OSUI + OpenSheetET
├── build.bat             Windows all-in-one build script
└── build.sh              Linux build script
```

---

## 📝 Changing the Qt Version

To use a different Qt version, edit `.github/workflows/build.yml`:
```yaml
env:
  QT_VERSION: 6.7.3   # ← change this
  QT_ARCH:    win64_msvc2022_64
```

Supported Qt versions for this project: **6.5.x, 6.6.x, 6.7.x**

---

## 🪪 License

MIT License — see [LICENSE.txt](LICENSE.txt)
