#pragma once
#include <QWidget>
#include <QToolButton>
#include <QComboBox>
#include <QSpinBox>
#include <QTabBar>
#include <QColor>
#include "Cell.h"

class ColorButton;

// ─────────────────────────────────────────────────────────────────────────────
//  RibbonWidget — WPS ET 2018white_dark style ribbon
//  Groups: Clipboard | Font | Alignment | Number | Styles | Cells | Editing
// ─────────────────────────────────────────────────────────────────────────────
class RibbonWidget : public QWidget {
    Q_OBJECT
public:
    explicit RibbonWidget(QWidget* parent = nullptr);

    // Sync ribbon state to cell format (called when selection changes)
    void syncToFormat(const CellFormat& fmt);

signals:
    // ── Clipboard ────────────────────────────────────────────────────────────
    void cutRequested();
    void copyRequested();
    void pasteRequested();
    void formatPainterRequested();

    // ── Font ─────────────────────────────────────────────────────────────────
    void fontFamilyChanged(const QString& family);
    void fontSizeChanged(int size);
    void boldToggled(bool on);
    void italicToggled(bool on);
    void underlineToggled(bool on);
    void strikeThroughToggled(bool on);
    void textColorChanged(const QColor& color);
    void fillColorChanged(const QColor& color);
    void borderRequested(BorderStyle style);

    // ── Alignment ─────────────────────────────────────────────────────────────
    void hAlignChanged(Qt::AlignmentFlag align);
    void vAlignChanged(Qt::AlignmentFlag align);
    void wrapTextToggled(bool on);
    void mergeCellsRequested(bool merge);
    void indentChanged(int delta);

    // ── Number ────────────────────────────────────────────────────────────────
    void numberFormatChanged(NumFmt fmt);
    void decimalsChanged(int delta);
    void formatConversionRequested(const QString& type);

    // ── Styles ────────────────────────────────────────────────────────────────
    void conditionalFormatRequested(const QString& type);
    void cellStyleRequested(const QString& styleName);
    void tableStyleRequested(const QString& styleName);

    // ── Cells ─────────────────────────────────────────────────────────────────
    void insertRowRequested();
    void deleteRowRequested();
    void insertColRequested();
    void deleteColRequested();
    void formatCellsRequested();
    void rowHeightRequested();
    void colWidthRequested();

    // ── Editing ───────────────────────────────────────────────────────────────
    void autoSumRequested();
    void fillRequested(const QString& direction);
    void sortRequested(bool ascending);
    void filterToggled();
    void freezeRequested(const QString& type);
    void findReplaceRequested();
    void functionWizardRequested();

    // ── File ──────────────────────────────────────────────────────────────────
    void newFileRequested();
    void openFileRequested();
    void saveFileRequested();
    void saveAsRequested();
    void exportCsvRequested();
    void importCsvRequested();
    void printRequested();

private:
    // Tab bar
    QTabBar*    m_tabs       { nullptr };
    QWidget*    m_homePanel  { nullptr };
    QWidget*    m_body       { nullptr };

    // Font group controls (need access to sync)
    QComboBox*  m_fontFamily  { nullptr };
    QSpinBox*   m_fontSize    { nullptr };
    QToolButton *m_btnBold, *m_btnItalic, *m_btnUnderline, *m_btnStrike;
    ColorButton *m_textColorBtn, *m_fillColorBtn;
    QToolButton *m_btnWrap, *m_btnFilter;
    QComboBox*   m_numFmtCombo { nullptr };
    QToolButton *m_btnAlL, *m_btnAlC, *m_btnAlR;
    QToolButton *m_btnVTop, *m_btnVMid, *m_btnVBot;

    void buildLayout();
    QWidget* buildHomeTab();
    QWidget* buildClipboardGroup();
    QWidget* buildFontGroup();
    QWidget* buildAlignGroup();
    QWidget* buildNumberGroup();
    QWidget* buildStylesGroup();
    QWidget* buildCellsGroup();
    QWidget* buildEditingGroup();

    // Factory helpers
    QToolButton* mkXLBtn(const QIcon& icon, const QString& label, const QString& tip);
    QToolButton* mkSmBtn(const QIcon& icon, const QString& tip, bool checkable=false);
    QToolButton* mkMdBtn(const QIcon& icon, const QString& label, const QString& tip);
    static QFrame* vSep();
    static QWidget* groupLabel(const QString& text);

    // Icon factory (QPainter-drawn)
    static QIcon makeIcon(std::function<void(QPainter&,int)> fn, int sz=20);
};
