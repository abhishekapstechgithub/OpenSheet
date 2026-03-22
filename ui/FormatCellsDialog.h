#pragma once
#include <QDialog>
#include "Cell.h"
class QComboBox; class QSpinBox; class QCheckBox; class QLineEdit; class QLabel;

class FormatCellsDialog : public QDialog {
    Q_OBJECT
public:
    explicit FormatCellsDialog(const CellFormat& fmt, QWidget* parent=nullptr);
    CellFormat format() const { return m_fmt; }
private:
    void apply();
    CellFormat m_fmt;
    QComboBox*  m_fontFam; QSpinBox* m_fontSize;
    QCheckBox  *m_bold,*m_italic,*m_underline,*m_strike,*m_wrap;
    QComboBox  *m_hAlign,*m_vAlign,*m_numFmt;
    QSpinBox*   m_decimals;
    QLabel     *m_textColorSw,*m_fillColorSw;
    QColor      m_textColor,m_fillColor;
};
