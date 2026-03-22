#pragma once
#include <QMainWindow>
#include <QCloseEvent>
#include "Workbook.h"
#include "RibbonWidget.h"
#include "SpreadsheetView.h"
#include "FormulaBar.h"
#include "SheetTabBar.h"
#include "NotifBar.h"
#include "StatusBar.h"
#include "Cell.h"

class MainWindow : public QMainWindow {
    Q_OBJECT
public:
    explicit MainWindow(QWidget* parent=nullptr);
    ~MainWindow() override;

protected:
    void closeEvent(QCloseEvent* e) override;
    void keyPressEvent(QKeyEvent* e) override;

private slots:
    // File
    void newFile();
    void openFile();
    void saveFile();
    void saveFileAs();
    void exportCsv();
    void importCsv();
    void printDoc();

    // Edit
    void undoAction();
    void redoAction();

    // Sheet navigation
    void onSheetActivated(int idx);
    void onSheetAdded();
    void onSheetListChanged();

    // Cell selection
    void onCurrentCellChanged(int row, int col);
    void onFormatChanged(const CellFormat& fmt, const QString& ref);

    // Ribbon handlers
    void onFontFamily(const QString& f);
    void onFontSize(int sz);
    void onBold(bool on);
    void onItalic(bool on);
    void onUnderline(bool on);
    void onStrikeThrough(bool on);
    void onTextColor(const QColor& c);
    void onFillColor(const QColor& c);
    void onBorder(BorderStyle style);
    void onHAlign(Qt::AlignmentFlag a);
    void onVAlign(Qt::AlignmentFlag a);
    void onWrapText(bool on);
    void onMergeCells(bool merge);
    void onIndent(int delta);
    void onNumberFormat(NumFmt fmt);
    void onDecimals(int delta);
    void onFmtConversion(const QString& type);
    void onCondFmt(const QString& type);
    void onCellStyle(const QString& name);
    void onTableStyle(const QString& name);
    void onInsertRow();
    void onDeleteRow();
    void onInsertCol();
    void onDeleteCol();
    void onFormatCells();
    void onRowHeight();
    void onColWidth();
    void onAutoSum();
    void onFill(const QString& dir);
    void onSort(bool asc);
    void onFilter();
    void onFreeze(const QString& type);
    void onFindReplace();
    void onFunctionWizard();

    // Formula bar
    void onFormulaCommitted(const QString& text);
    void onCellRefNavigated(const QString& ref);

    // Zoom
    void onZoom(int pct);

private:
    void buildMenuBar();
    void buildCentralWidget();
    void updateTitle();
    void updateFormulaBar();
    void updateStatusStats();
    void applyFmtDelta(std::function<void(CellFormat&)> fn);
    void pushUndo();
    bool confirmDiscard();

    // Core
    Workbook*        m_wb      { nullptr };

    // UI
    RibbonWidget*    m_ribbon  { nullptr };
    NotifBar*        m_notif   { nullptr };
    FormulaBar*      m_fbar    { nullptr };
    SpreadsheetView* m_view    { nullptr };
    SheetTabBar*     m_sheetBar{ nullptr };
    OSStatusBar*     m_stbar   { nullptr };

    // State
    bool m_ignoreFormatSignals { false };
};
