; ============================================================
;  OpenSheet ET — NSIS Installer
;  Builds a professional Windows installer (.exe)
;  Requirements: NSIS 3.x — https://nsis.sourceforge.io
; ============================================================

!define APP_NAME        "OpenSheet ET"
!define APP_VERSION     "1.0.0"
!define APP_PUBLISHER   "OpenSheet"
!define APP_EXE         "OpenSheetET.exe"
!define APP_ICON        "..\resources\app.ico"
!define REG_KEY         "Software\OpenSheet\ET"
!define UNINSTALL_KEY   "Software\Microsoft\Windows\CurrentVersion\Uninstall\OpenSheetET"

Name         "${APP_NAME} ${APP_VERSION}"
OutFile      "..\OpenSheetET-${APP_VERSION}-Setup.exe"
Unicode      True
SetCompressor /SOLID lzma

; Request admin for per-machine install, user for per-user
RequestExecutionLevel user

; ── Pages ────────────────────────────────────────────────────────────────────
!include "MUI2.nsh"
!include "LogicLib.nsh"

!define MUI_ABORTWARNING
!define MUI_ICON         "${APP_ICON}"
!define MUI_UNICON       "${APP_ICON}"
!define MUI_HEADERIMAGE
!define MUI_HEADERIMAGE_BITMAP "${NSISDIR}\Contrib\Graphics\Header\orange-r.bmp"
!define MUI_WELCOMEPAGE_TITLE  "Welcome to ${APP_NAME} Setup"
!define MUI_WELCOMEPAGE_TEXT   "This wizard will guide you through the installation of ${APP_NAME} ${APP_VERSION}.$\n$\nA modern, fast spreadsheet application for Windows.$\n$\nClick Next to continue."
!define MUI_FINISHPAGE_RUN     "$INSTDIR\${APP_EXE}"
!define MUI_FINISHPAGE_RUN_TEXT "Launch ${APP_NAME}"
!define MUI_FINISHPAGE_SHOWREADME ""

; Scope selection page (All Users vs Just Me)
Var InstallScope
Page custom PageScope PageScopeLeave

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "..\LICENSE.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

; ── Scope page ───────────────────────────────────────────────────────────────
Function PageScope
  nsDialogs::Create 1044
  Pop $0
  ${NSD_CreateLabel} 0 0 100% 20u "Install ${APP_NAME} for:"
  Pop $0
  ${NSD_CreateRadioButton} 10u 25u 100% 12u "Anyone who uses this computer (requires administrator)"
  Pop $1
  ${NSD_CreateRadioButton} 10u 42u 100% 12u "Only for me (no administrator required)"
  Pop $2
  StrCmp $InstallScope "user" +2
  ${NSD_Check} $1
  ${NSD_Check} $2
  nsDialogs::Show
FunctionEnd

Function PageScopeLeave
  ${NSD_GetState} $1 $0
  ${If} $0 == ${BST_CHECKED}
    StrCpy $InstallScope "machine"
    SetShellVarContext all
    StrCpy $INSTDIR "$PROGRAMFILES64\${APP_NAME}"
  ${Else}
    StrCpy $InstallScope "user"
    SetShellVarContext current
    StrCpy $INSTDIR "$LOCALAPPDATA\Programs\${APP_NAME}"
  ${EndIf}
FunctionEnd

; ── Install ───────────────────────────────────────────────────────────────────
Section "Main Application" SEC_MAIN
  SectionIn RO
  SetOutPath "$INSTDIR"

  ; Deploy all files from deploy directory
  File /r "..\deploy\*.*"

  ; Write registry
  WriteRegStr   HKCU "${REG_KEY}" "InstallDir"    "$INSTDIR"
  WriteRegStr   HKCU "${REG_KEY}" "Version"        "${APP_VERSION}"
  WriteRegStr   HKCU "${UNINSTALL_KEY}" "DisplayName"          "${APP_NAME}"
  WriteRegStr   HKCU "${UNINSTALL_KEY}" "DisplayVersion"       "${APP_VERSION}"
  WriteRegStr   HKCU "${UNINSTALL_KEY}" "Publisher"            "${APP_PUBLISHER}"
  WriteRegStr   HKCU "${UNINSTALL_KEY}" "InstallLocation"      "$INSTDIR"
  WriteRegStr   HKCU "${UNINSTALL_KEY}" "DisplayIcon"          "$INSTDIR\${APP_EXE}"
  WriteRegStr   HKCU "${UNINSTALL_KEY}" "UninstallString"      '"$INSTDIR\Uninstall.exe"'
  WriteRegStr   HKCU "${UNINSTALL_KEY}" "QuietUninstallString" '"$INSTDIR\Uninstall.exe" /S'
  WriteRegDWORD HKCU "${UNINSTALL_KEY}" "NoModify"             1
  WriteRegDWORD HKCU "${UNINSTALL_KEY}" "NoRepair"             1

  ; File associations (.oset files)
  WriteRegStr HKCU "Software\Classes\.oset"                ""                "OpenSheetET.Document"
  WriteRegStr HKCU "Software\Classes\OpenSheetET.Document" ""                "${APP_NAME} Spreadsheet"
  WriteRegStr HKCU "Software\Classes\OpenSheetET.Document\DefaultIcon" "" "$INSTDIR\${APP_EXE},0"
  WriteRegStr HKCU "Software\Classes\OpenSheetET.Document\shell\open\command" "" '"$INSTDIR\${APP_EXE}" "%1"'

  ; Uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd

Section "Desktop Shortcut" SEC_DESKTOP
  CreateShortcut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}" "" "$INSTDIR\${APP_EXE}" 0
SectionEnd

Section "Start Menu Shortcuts" SEC_START
  CreateDirectory "$SMPROGRAMS\${APP_NAME}"
  CreateShortcut "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk"    "$INSTDIR\${APP_EXE}" "" "$INSTDIR\${APP_EXE}" 0
  CreateShortcut "$SMPROGRAMS\${APP_NAME}\Uninstall.lnk"      "$INSTDIR\Uninstall.exe"
SectionEnd

; ── Uninstall ────────────────────────────────────────────────────────────────
Section "Uninstall"
  SetShellVarContext current
  ; Remove all installed files
  RMDir /r "$INSTDIR"
  ; Remove shortcuts
  Delete "$DESKTOP\${APP_NAME}.lnk"
  RMDir /r "$SMPROGRAMS\${APP_NAME}"
  ; Remove registry
  DeleteRegKey HKCU "${REG_KEY}"
  DeleteRegKey HKCU "${UNINSTALL_KEY}"
  DeleteRegKey HKCU "Software\Classes\.oset"
  DeleteRegKey HKCU "Software\Classes\OpenSheetET.Document"
  ; Done
  MessageBox MB_OK "${APP_NAME} has been uninstalled."
SectionEnd

; ── Descriptions ─────────────────────────────────────────────────────────────
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC_MAIN}    "The main application files (required)."
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC_DESKTOP} "Create a desktop shortcut."
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC_START}   "Create Start Menu entries."
!insertmacro MUI_FUNCTION_DESCRIPTION_END
