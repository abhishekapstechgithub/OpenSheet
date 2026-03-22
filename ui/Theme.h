#pragma once
#include <QString>

// WPS ET 2018white colour theme
namespace Theme {
    // Greens
    constexpr auto GREEN        = "#1e7145";
    constexpr auto GREEN_DARK   = "#155c34";
    constexpr auto GREEN_MID    = "#2d9e5f";
    constexpr auto GREEN_LIGHT  = "#e8f5ee";
    constexpr auto GREEN_MED    = "#c8e8d8";
    // Backgrounds
    constexpr auto BG           = "#ffffff";
    constexpr auto BG2          = "#f5f6f7";
    constexpr auto BG3          = "#eef0f3";
    // Borders
    constexpr auto BD           = "#d8dce3";
    constexpr auto BD2          = "#c4c9d4";
    // Text
    constexpr auto TX           = "#1a1d23";
    constexpr auto TX2          = "#454d5c";
    constexpr auto TX3          = "#7a8494";
    constexpr auto TX4          = "#b8bfcc";
    // Misc
    constexpr auto RED          = "#d32f2f";
    constexpr auto AMBER        = "#e65100";

    // Global stylesheet
    inline QString appStyle() {
        return R"(
QWidget { font-family:'Segoe UI',Arial,sans-serif; font-size:12px; color:#1a1d23; }
QMainWindow { background:#f5f6f7; }
QMenuBar { background:#155c34; color:white; padding:0 6px; font-size:12px; }
QMenuBar::item { padding:5px 10px; border-radius:3px; }
QMenuBar::item:selected { background:rgba(255,255,255,0.18); }
QMenuBar::item:pressed  { background:rgba(255,255,255,0.28); }
QMenu { background:white; border:1px solid #d8dce3; border-radius:4px; padding:3px 0; font-size:12px; }
QMenu::item { padding:6px 20px 6px 14px; }
QMenu::item:selected { background:#e8f5ee; color:#1e7145; }
QMenu::separator { height:1px; background:#d8dce3; margin:3px 0; }
QToolBar { background:#f5f6f7; border:none; spacing:2px; }
QStatusBar { background:#f0f2f5; border-top:1px solid #d8dce3; font-size:11px; }
QScrollBar:horizontal { height:8px; background:#f5f6f7; }
QScrollBar:vertical   { width:8px;  background:#f5f6f7; }
QScrollBar::handle { background:#c4c9d4; border-radius:4px; margin:2px; }
QScrollBar::handle:hover { background:#7a8494; }
QScrollBar::add-line, QScrollBar::sub-line { width:0; height:0; }
QDialog { background:white; }
QLineEdit,QSpinBox,QDoubleSpinBox,QComboBox {
  border:1px solid #d8dce3; border-radius:3px; background:white;
  padding:2px 6px; selection-background-color:#c8e8d8; }
QLineEdit:focus,QSpinBox:focus,QComboBox:focus { border-color:#1e7145; }
QPushButton {
  border:1px solid #d8dce3; border-radius:3px; background:#f5f6f7;
  padding:5px 14px; font-size:12px; }
QPushButton:hover { background:#e8f5ee; border-color:#c8e8d8; }
QPushButton:pressed { background:#c8e8d8; border-color:#1e7145; }
QPushButton[primary="true"] { background:#1e7145; color:white; border:none; }
QPushButton[primary="true"]:hover { background:#155c34; }
QGroupBox { border:1px solid #d8dce3; border-radius:3px; margin-top:14px; }
QGroupBox::title { subcontrol-origin:margin; left:6px; padding:0 4px; color:#7a8494; font-size:11px; }
QTabWidget::pane { border:1px solid #d8dce3; background:white; }
QTabBar::tab { padding:6px 16px; background:#eef0f3; border:1px solid #d8dce3; border-bottom:none; margin-right:1px; border-radius:3px 3px 0 0; }
QTabBar::tab:selected { background:white; color:#1e7145; font-weight:600; border-color:#1e7145; }
QCheckBox::indicator { width:13px; height:13px; border:1px solid #c4c9d4; border-radius:2px; }
QCheckBox::indicator:checked { background:#1e7145; border-color:#1e7145; image:url(none); }
QRadioButton::indicator { width:13px; height:13px; border-radius:7px; border:1px solid #c4c9d4; }
QRadioButton::indicator:checked { background:#1e7145; border-color:#1e7145; }
        )";
    }
}
