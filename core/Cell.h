#pragma once
#include <QString>
#include <QColor>
#include <QFont>
#include <QVariant>
#include <optional>

// ─── Number format ────────────────────────────────────────────────────────────
enum class NumFmt {
    General, Number, Currency, Accounting, Percent,
    Scientific, ShortDate, LongDate, Time, Fraction, Text, Thousands
};

// ─── Border style ─────────────────────────────────────────────────────────────
enum class BorderStyle { None, All, Outside, Top, Bottom, Left, Right, Thick };

// ─── Cell format ──────────────────────────────────────────────────────────────
struct CellFormat {
    // Font
    QString     fontFamily   { "Calibri" };
    int         fontSize     { 11 };
    bool        bold         { false };
    bool        italic       { false };
    bool        underline    { false };
    bool        strikeThrough{ false };
    bool        wrapText     { false };
    int         indent       { 0 };

    // Colours
    QColor      textColor    { Qt::black };
    QColor      fillColor    { Qt::transparent };

    // Alignment
    Qt::AlignmentFlag hAlign { Qt::AlignLeft };
    Qt::AlignmentFlag vAlign { Qt::AlignVCenter };

    // Number formatting
    NumFmt      numFmt       { NumFmt::General };
    int         decimals     { 2 };

    // Border
    BorderStyle border       { BorderStyle::None };

    // Merge
    bool        merged       { false };

    // Conditional formatting colour override (separate from user fill)
    QColor      condFillColor{ Qt::transparent };

    bool hasCondFill() const { return condFillColor.isValid() && condFillColor != Qt::transparent; }
};

// ─── Cell ─────────────────────────────────────────────────────────────────────
struct Cell {
    QVariant  rawValue;     // What the user typed  (string/number/formula)
    QVariant  cachedValue;  // Last evaluated result
    CellFormat format;
    bool      dirty { true }; // needs re-evaluation?

    bool isEmpty() const { return !rawValue.isValid() || rawValue.isNull() || rawValue.toString().isEmpty(); }
    bool isFormula() const { return rawValue.userType() == QMetaType::QString && rawValue.toString().startsWith('='); }
};
