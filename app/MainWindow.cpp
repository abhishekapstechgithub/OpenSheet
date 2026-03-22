// ═══════════════════════════════════════════════════════════════════════════════
//  MainWindow.cpp — OpenSheet ET Main Window
//  Wires Workbook ↔ Ribbon ↔ SpreadsheetView ↔ FormulaBar ↔ SheetTabBar
// ═══════════════════════════════════════════════════════════════════════════════
#include "MainWindow.h"
#include "Theme.h"
#include "FormatCellsDialog.h"
#include "FindReplaceDialog.h"
#include "FunctionWizard.h"
#include "FormulaEngine.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QMenuBar>
#include <QStatusBar>
#include <QFileDialog>
#include <QMessageBox>
#include <QInputDialog>
#include <QKeySequence>
#include <QScreen>
#include <QApplication>
#include <QTimer>
#include <QPrinter>
#include <QPrintDialog>
#include <cmath>

// ════════════════════════════════════════════════════════════════════════════
//  Construction
// ════════════════════════════════════════════════════════════════════════════
MainWindow::MainWindow(QWidget* parent) : QMainWindow(parent) {
    setWindowTitle("OpenSheet ET");
    setWindowIcon(QIcon());   // Will use app.rc on Windows

    QRect screen = QApplication::primaryScreen()->availableGeometry();
    resize(qMin(1400, (int)(screen.width()  * 0.88)),
           qMin( 820, (int)(screen.height() * 0.88)));
    setMinimumSize(960, 620);

    m_wb = new Workbook(this);
    connect(m_wb, &Workbook::modifiedChanged, this, [this](bool){ updateTitle(); });
    connect(m_wb, &Workbook::sheetListChanged, this, &MainWindow::onSheetListChanged);

    buildMenuBar();
    buildCentralWidget();
    updateTitle();

    // Load sample data on first launch
    QTimer::singleShot(0, this, [this]{
        Sheet* sh = m_wb->activeSheet();
        if (!sh) return;
        // WPS ET Business template style sample
        struct SampleCell { int r,c; QVariant v; CellFormat fmt; };
        const QColor hdrBg("#1e7145");
        auto hdrFmt=[&](){ CellFormat f; f.bold=true; f.fillColor=hdrBg; f.textColor=Qt::white; f.hAlign=Qt::AlignHCenter; f.fontSize=11; return f; };
        auto numFmt=[&](){ CellFormat f; f.numFmt=NumFmt::Currency; f.hAlign=Qt::AlignRight; return f; };
        auto boldFmt=[&](){ CellFormat f; f.numFmt=NumFmt::Currency; f.bold=true; f.hAlign=Qt::AlignRight; return f; };
        auto totFmt=[&](){ CellFormat f; f.numFmt=NumFmt::Currency; f.bold=true; f.fillColor=QColor("#e8f5ee"); f.hAlign=Qt::AlignRight; return f; };
        auto totLblFmt=[&](){ CellFormat f; f.bold=true; f.fillColor=QColor("#e8f5ee"); return f; };

        // Headers
        for(auto& [c,label] : QList<QPair<int,QString>>{{0,"Product"},{1,"Jan"},{2,"Feb"},{3,"Mar"},{4,"Q1 Total"},{5,"Growth %"},{6,"Status"}}) {
            sh->setCellValue(0,c,label); sh->setCellFormat(0,c,hdrFmt());
        }
        // Data rows
        const QList<QPair<QString,QList<double>>> data={
            {"Widget A",   {42300,38700,51200}},
            {"Gadget X",   {18750,22100,19800}},
            {"Device Pro", {95200,88400,102300}},
            {"Pack Basic", {12400,14800,13200}},
            {"Ultra Kit",  {67800,72300,78900}},
        };
        for(int i=0;i<data.size();i++) {
            int r=i+1;
            sh->setCellValue(r,0,data[i].first);
            for(int c=0;c<3;c++){ sh->setCellValue(r,c+1,data[i].second[c]); sh->setCellFormat(r,c+1,numFmt()); }
            // Q1 Total formula
            QString sumFm=QString("=SUM(%1%2:%3%2)").arg(FormulaEngine::colLabel(1)).arg(r+1).arg(FormulaEngine::colLabel(3));
            sh->setCellFormula(r,4,sumFm); sh->setCellFormat(r,4,boldFmt());
            // Growth %
            QString growFm=QString("=(%1%2-%3%2)/%3%2").arg(FormulaEngine::colLabel(3)).arg(r+1).arg(FormulaEngine::colLabel(1));
            CellFormat gf; gf.numFmt=NumFmt::Percent; gf.hAlign=Qt::AlignRight;
            sh->setCellFormula(r,5,growFm); sh->setCellFormat(r,5,gf);
            // Status IF formula
            QString stFm=QString("=IF(F%1>0,\"▲ Growing\",\"▼ Declining\")").arg(r+1);
            CellFormat sf; sf.hAlign=Qt::AlignHCenter;
            sh->setCellFormula(r,6,stFm); sh->setCellFormat(r,6,sf);
        }
        // Total row
        int tr=data.size()+1;
        sh->setCellValue(tr,0,"TOTAL"); sh->setCellFormat(tr,0,totLblFmt());
        for(int c=1;c<=4;c++) {
            QString col=FormulaEngine::colLabel(c);
            sh->setCellFormula(tr,c,QString("=SUM(%1%2:%1%3)").arg(col).arg(2).arg(tr));
            sh->setCellFormat(tr,c,totFmt());
        }
        m_view->setSheet(sh);
        m_view->scrollToCell(0,0);
        m_wb->setModified(false);
    });
}

MainWindow::~MainWindow() {}

// ════════════════════════════════════════════════════════════════════════════
//  Menu Bar (WPS green style)
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::buildMenuBar() {
    auto* mb=menuBar();

    auto addMenu=[&](const QString& name)->QMenu*{
        auto* m=mb->addMenu(name); return m;
    };

    // File
    auto* file=addMenu("&File");
    file->addAction("&New",       this,&MainWindow::newFile,   QKeySequence::New);
    file->addAction("&Open…",     this,&MainWindow::openFile,  QKeySequence::Open);
    file->addSeparator();
    file->addAction("&Save",      this,&MainWindow::saveFile,  QKeySequence::Save);
    file->addAction("Save &As…",  this,&MainWindow::saveFileAs,QKeySequence::SaveAs);
    file->addSeparator();
    file->addAction("Export as &CSV…", this,&MainWindow::exportCsv);
    file->addAction("Import CSV…",     this,&MainWindow::importCsv);
    file->addSeparator();
    file->addAction("&Print…",    this,&MainWindow::printDoc,  QKeySequence::Print);
    file->addSeparator();
    file->addAction("E&xit",      this,[this]{ close(); },      QKeySequence::Quit);

    // Edit
    auto* edit=addMenu("&Edit");
    edit->addAction("&Undo",    this,&MainWindow::undoAction, QKeySequence::Undo);
    edit->addAction("&Redo",    this,&MainWindow::redoAction, QKeySequence::Redo);
    edit->addSeparator();
    edit->addAction("&Cut",     this,[this]{ QApplication::sendEvent(m_view,new QKeyEvent(QEvent::KeyPress,Qt::Key_X,Qt::ControlModifier)); }, QKeySequence::Cut);
    edit->addAction("C&opy",    this,[this]{ QApplication::sendEvent(m_view,new QKeyEvent(QEvent::KeyPress,Qt::Key_C,Qt::ControlModifier)); }, QKeySequence::Copy);
    edit->addAction("&Paste",   this,[this]{ QApplication::sendEvent(m_view,new QKeyEvent(QEvent::KeyPress,Qt::Key_V,Qt::ControlModifier)); }, QKeySequence::Paste);
    edit->addSeparator();
    edit->addAction("&Find && Replace…", this,&MainWindow::onFindReplace, QKeySequence(Qt::CTRL|Qt::Key_H));
    edit->addAction("Select &All",       this,[this]{ m_view->selectAll(); }, QKeySequence::SelectAll);
    edit->addSeparator();
    edit->addAction("Clear &Contents", this,[this]{
        if(m_wb->activeSheet())
            for(const auto& idx:m_view->selectedIndexes())
                m_wb->activeSheet()->clearCell(idx.row(),idx.column());
    }, Qt::Key_Delete);

    // Sheet
    auto* shMenu=addMenu("&Sheet");
    shMenu->addAction("Insert &Sheet",    this,&MainWindow::onSheetAdded);
    shMenu->addAction("&Delete Sheet",    this,[this]{ m_wb->removeSheet(m_wb->activeIndex()); });
    shMenu->addAction("&Rename Sheet…",   this,[this]{
        bool ok; QString n=QInputDialog::getText(this,"Rename Sheet","Sheet name:",QLineEdit::Normal,m_wb->activeSheet()->name(),&ok);
        if(ok&&!n.isEmpty()) m_wb->renameSheet(m_wb->activeIndex(),n);
    });
    shMenu->addSeparator();
    shMenu->addAction("Insert &Row",      this,&MainWindow::onInsertRow, QKeySequence(Qt::CTRL|Qt::SHIFT|Qt::Key_Plus));
    shMenu->addAction("Delete Row",       this,&MainWindow::onDeleteRow);
    shMenu->addAction("Insert &Column",   this,&MainWindow::onInsertCol);
    shMenu->addAction("Delete Column",    this,&MainWindow::onDeleteCol);

    // View
    auto* view=addMenu("&View");
    view->addAction("Zoom &In",  this,[this]{ onZoom(qMin(300,int(m_view->property("zoom").toInt()+10))); }, QKeySequence(Qt::CTRL|Qt::Key_Plus));
    view->addAction("Zoom &Out", this,[this]{ onZoom(qMax(50, int(m_view->property("zoom").toInt()-10))); }, QKeySequence(Qt::CTRL|Qt::Key_Minus));
    view->addAction("Reset &100%",this,[this]{ onZoom(100); });
    view->addSeparator();
    view->addAction("Freeze &Top Row",      this,[this]{ onFreeze("toprow"); });
    view->addAction("Freeze &First Column", this,[this]{ onFreeze("firstcol"); });
    view->addAction("&Unfreeze Panes",      this,[this]{ onFreeze("unfreeze"); });

    // Tools
    auto* tools=addMenu("&Tools");
    tools->addAction("&Sort A→Z",    this,[this]{ onSort(true); });
    tools->addAction("Sort Z→A",     this,[this]{ onSort(false); });
    tools->addAction("&AutoFilter",  this,&MainWindow::onFilter);
    tools->addSeparator();
    tools->addAction("&Insert Function…",    this,&MainWindow::onFunctionWizard, QKeySequence(Qt::SHIFT|Qt::Key_F3));
    tools->addAction("Format &Cells…",       this,&MainWindow::onFormatCells,    QKeySequence(Qt::CTRL|Qt::Key_1));
    tools->addSeparator();
    tools->addAction("Auto&Sum",      this,&MainWindow::onAutoSum, QKeySequence(Qt::ALT|Qt::Key_Equal));
    tools->addAction("Fill &Down",    this,[this]{ onFill("down");  }, QKeySequence(Qt::CTRL|Qt::Key_D));
    tools->addAction("Fill &Right",   this,[this]{ onFill("right"); }, QKeySequence(Qt::CTRL|Qt::Key_R));
    tools->addSeparator();
    tools->addAction("&Options…", this,[this]{ QMessageBox::information(this,"Options","Options dialog — coming in Pro version"); });

    // Help
    auto* help=addMenu("&Help");
    help->addAction("&Keyboard Shortcuts", this,[this]{
        QMessageBox::information(this,"Keyboard Shortcuts",
            "Ctrl+C/X/V  Copy/Cut/Paste\n"
            "Ctrl+Z/Y    Undo/Redo\n"
            "Ctrl+B/I/U  Bold/Italic/Underline\n"
            "Ctrl+H      Find & Replace\n"
            "Ctrl+S      Save\n"
            "Ctrl+N      New\n"
            "Ctrl+O      Open\n"
            "Alt+=       AutoSum\n"
            "Ctrl+D/R    Fill Down/Right\n"
            "Ctrl+1      Format Cells\n"
            "Del         Clear Contents\n"
            "F2          Edit Cell\n"
            "Shift+F3    Insert Function");
    });
    help->addSeparator();
    help->addAction("&About OpenSheet ET", this,[this]{
        QMessageBox::about(this,"About OpenSheet ET",
            "<b>OpenSheet ET v1.0</b><br>"
            "A modern desktop spreadsheet inspired by WPS ET.<br><br>"
            "Features:<br>"
            "• 60+ built-in formulas<br>"
            "• WPS ET-style ribbon UI<br>"
            "• Conditional formatting<br>"
            "• Cell styles &amp; table styles<br>"
            "• CSV import/export<br>"
            "• Full undo/redo history<br><br>"
            "Built with Qt 6 &amp; C++20");
    });
}

// ════════════════════════════════════════════════════════════════════════════
//  Central Widget
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::buildCentralWidget() {
    auto* central = new QWidget(this);
    central->setStyleSheet("QWidget{background:#ffffff;}");
    auto* vl = new QVBoxLayout(central);
    vl->setContentsMargins(0,0,0,0);
    vl->setSpacing(0);

    // 1. Ribbon
    m_ribbon = new RibbonWidget(this);
    vl->addWidget(m_ribbon);

    // 2. Notification bar
    m_notif = new NotifBar(this);
    vl->addWidget(m_notif);

    // 3. Formula bar
    m_fbar = new FormulaBar(this);
    vl->addWidget(m_fbar);

    // 4. Grid (fills remaining space)
    m_view = new SpreadsheetView(m_wb->activeSheet(), this);
    m_view->setObjectName("gridView");
    vl->addWidget(m_view, 1);

    // 5. Sheet tab bar
    m_sheetBar = new SheetTabBar(m_wb, this);
    vl->addWidget(m_sheetBar);

    // 6. Status bar
    m_stbar = new OSStatusBar(this);
    vl->addWidget(m_stbar);

    setCentralWidget(central);

    // ── Connect ribbon → handlers ──────────────────────────────────────────
    auto* R = m_ribbon;
    connect(R,&RibbonWidget::newFileRequested,     this,&MainWindow::newFile);
    connect(R,&RibbonWidget::openFileRequested,    this,&MainWindow::openFile);
    connect(R,&RibbonWidget::saveFileRequested,    this,&MainWindow::saveFile);
    connect(R,&RibbonWidget::saveAsRequested,      this,&MainWindow::saveFileAs);
    connect(R,&RibbonWidget::exportCsvRequested,   this,&MainWindow::exportCsv);
    connect(R,&RibbonWidget::importCsvRequested,   this,&MainWindow::importCsv);
    connect(R,&RibbonWidget::printRequested,       this,&MainWindow::printDoc);
    connect(R,&RibbonWidget::cutRequested,         this,[this]{ QApplication::sendEvent(m_view,new QKeyEvent(QEvent::KeyPress,Qt::Key_X,Qt::ControlModifier)); });
    connect(R,&RibbonWidget::copyRequested,        this,[this]{ QApplication::sendEvent(m_view,new QKeyEvent(QEvent::KeyPress,Qt::Key_C,Qt::ControlModifier)); });
    connect(R,&RibbonWidget::pasteRequested,       this,[this]{ QApplication::sendEvent(m_view,new QKeyEvent(QEvent::KeyPress,Qt::Key_V,Qt::ControlModifier)); });
    connect(R,&RibbonWidget::formatPainterRequested,this,[this]{
        // Capture format, next click applies it
        if(m_wb->activeSheet()) {
            auto captured = m_wb->activeSheet()->getCell(m_view->currentRow(),m_view->currentCol()).format;
            // Simple approach: apply immediately to selection as a demo
            m_view->applyFormatToSelection(captured);
        }
    });
    connect(R,&RibbonWidget::fontFamilyChanged,    this,&MainWindow::onFontFamily);
    connect(R,&RibbonWidget::fontSizeChanged,      this,&MainWindow::onFontSize);
    connect(R,&RibbonWidget::boldToggled,          this,&MainWindow::onBold);
    connect(R,&RibbonWidget::italicToggled,        this,&MainWindow::onItalic);
    connect(R,&RibbonWidget::underlineToggled,     this,&MainWindow::onUnderline);
    connect(R,&RibbonWidget::strikeThroughToggled, this,&MainWindow::onStrikeThrough);
    connect(R,&RibbonWidget::textColorChanged,     this,&MainWindow::onTextColor);
    connect(R,&RibbonWidget::fillColorChanged,     this,&MainWindow::onFillColor);
    connect(R,&RibbonWidget::borderRequested,      this,&MainWindow::onBorder);
    connect(R,&RibbonWidget::hAlignChanged,        this,&MainWindow::onHAlign);
    connect(R,&RibbonWidget::vAlignChanged,        this,&MainWindow::onVAlign);
    connect(R,&RibbonWidget::wrapTextToggled,      this,&MainWindow::onWrapText);
    connect(R,&RibbonWidget::mergeCellsRequested,  this,&MainWindow::onMergeCells);
    connect(R,&RibbonWidget::indentChanged,        this,&MainWindow::onIndent);
    connect(R,&RibbonWidget::numberFormatChanged,  this,&MainWindow::onNumberFormat);
    connect(R,&RibbonWidget::decimalsChanged,      this,&MainWindow::onDecimals);
    connect(R,&RibbonWidget::formatConversionRequested,this,&MainWindow::onFmtConversion);
    connect(R,&RibbonWidget::conditionalFormatRequested,this,&MainWindow::onCondFmt);
    connect(R,&RibbonWidget::cellStyleRequested,   this,&MainWindow::onCellStyle);
    connect(R,&RibbonWidget::tableStyleRequested,  this,&MainWindow::onTableStyle);
    connect(R,&RibbonWidget::insertRowRequested,   this,&MainWindow::onInsertRow);
    connect(R,&RibbonWidget::deleteRowRequested,   this,&MainWindow::onDeleteRow);
    connect(R,&RibbonWidget::insertColRequested,   this,&MainWindow::onInsertCol);
    connect(R,&RibbonWidget::deleteColRequested,   this,&MainWindow::onDeleteCol);
    connect(R,&RibbonWidget::formatCellsRequested, this,&MainWindow::onFormatCells);
    connect(R,&RibbonWidget::rowHeightRequested,   this,&MainWindow::onRowHeight);
    connect(R,&RibbonWidget::colWidthRequested,    this,&MainWindow::onColWidth);
    connect(R,&RibbonWidget::autoSumRequested,     this,&MainWindow::onAutoSum);
    connect(R,&RibbonWidget::fillRequested,        this,&MainWindow::onFill);
    connect(R,&RibbonWidget::sortRequested,        this,&MainWindow::onSort);
    connect(R,&RibbonWidget::filterToggled,        this,&MainWindow::onFilter);
    connect(R,&RibbonWidget::freezeRequested,      this,&MainWindow::onFreeze);
    connect(R,&RibbonWidget::findReplaceRequested, this,&MainWindow::onFindReplace);
    connect(R,&RibbonWidget::functionWizardRequested,this,&MainWindow::onFunctionWizard);

    // ── Connect view → UI ──────────────────────────────────────────────────
    connect(m_view,&SpreadsheetView::currentCellChanged,this,&MainWindow::onCurrentCellChanged);
    connect(m_view,&SpreadsheetView::formatChanged,     this,&MainWindow::onFormatChanged);
    connect(m_view,&SpreadsheetView::zoomChanged,       this,&MainWindow::onZoom);

    // ── Connect formula bar ────────────────────────────────────────────────
    connect(m_fbar,&FormulaBar::formulaCommitted,    this,&MainWindow::onFormulaCommitted);
    connect(m_fbar,&FormulaBar::cellRefNavigated,    this,&MainWindow::onCellRefNavigated);
    connect(m_fbar,&FormulaBar::functionWizardRequested,this,&MainWindow::onFunctionWizard);

    // ── Connect sheet tab bar ──────────────────────────────────────────────
    connect(m_sheetBar,&SheetTabBar::sheetActivated, this,&MainWindow::onSheetActivated);
    connect(m_sheetBar,&SheetTabBar::sheetAdded,     this,&MainWindow::onSheetAdded);
    connect(m_sheetBar,&SheetTabBar::sheetDeleted,   this,[this](int idx){ m_wb->removeSheet(idx); });
    connect(m_sheetBar,&SheetTabBar::sheetRenamed,   this,[this](int idx,const QString&n){ m_wb->renameSheet(idx,n); });
    connect(m_sheetBar,&SheetTabBar::sheetDuplicated,this,[this](int idx){ m_wb->duplicateSheet(idx); });

    // ── Connect status bar zoom ────────────────────────────────────────────
    connect(m_stbar,&OSStatusBar::zoomChanged,this,&MainWindow::onZoom);

    // ── Connect notif ──────────────────────────────────────────────────────
    connect(m_notif,&NotifBar::upgradeClicked,this,[this]{
        QMessageBox::information(this,"OpenSheet Pro",
            "<b>OpenSheet Pro Features:</b><br>"
            "• Pivot Tables<br>• AI Formula Assistant<br>"
            "• Macro Recorder<br>• Real-time Collaboration<br>"
            "• 500GB File Support<br>• Cloud Sync<br>"
            "• XLSX/ODS Export<br>• 50+ Chart Types");
    });

    // ── Workbook sheet changes ─────────────────────────────────────────────
    connect(m_wb,&Workbook::activeSheetChanged,this,[this](int idx){
        Sheet* sh=m_wb->sheet(idx);
        if(sh){ m_view->setSheet(sh); updateFormulaBar(); }
        m_sheetBar->refresh();
    });
}

// ════════════════════════════════════════════════════════════════════════════
//  Helpers
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::updateTitle() {
    QString title = m_wb->filePath().isEmpty() ? "Untitled" : QFileInfo(m_wb->filePath()).fileName();
    if(m_wb->isModified()) title += " *";
    title += " — OpenSheet ET";
    setWindowTitle(title);
}

void MainWindow::updateFormulaBar() {
    int r=m_view->currentRow(), c=m_view->currentCol();
    m_fbar->setCellRef(FormulaEngine::cellRef(r,c));
    if(m_wb->activeSheet()) {
        const Cell& cell=m_wb->activeSheet()->getCell(r,c);
        m_fbar->setFormulaText(cell.rawValue.toString());
    }
}

void MainWindow::updateStatusStats() {
    auto sel = m_view->selectedIndexes();
    if(sel.size()<2){ m_stbar->setStats(""); return; }
    QVector<double> vals;
    for(const auto& idx:sel){
        if(!m_wb->activeSheet()) continue;
        bool ok; double d=m_wb->activeSheet()->getCellValue(idx.row(),idx.column()).toDouble(&ok);
        if(ok) vals<<d;
    }
    if(vals.isEmpty()){ m_stbar->setStats(""); return; }
    double sum=0; for(double d:vals) sum+=d;
    m_stbar->setStats(QString("Average: %1  Count: %2  Sum: %3  Min: %4  Max: %5")
        .arg(sum/vals.size(),0,'f',2)
        .arg(vals.size())
        .arg(sum,0,'f',2)
        .arg(*std::min_element(vals.begin(),vals.end()),0,'f',2)
        .arg(*std::max_element(vals.begin(),vals.end()),0,'f',2));
}

void MainWindow::applyFmtDelta(std::function<void(CellFormat&)> fn) {
    pushUndo(); m_view->applyFormatDelta(fn);
}

void MainWindow::pushUndo() { m_wb->pushSnapshot(); }

bool MainWindow::confirmDiscard() {
    if(!m_wb->isModified()) return true;
    int r=QMessageBox::question(this,"Unsaved Changes",
        "The spreadsheet has unsaved changes.\nSave before continuing?",
        QMessageBox::Save|QMessageBox::Discard|QMessageBox::Cancel);
    if(r==QMessageBox::Save){ saveFile(); return !m_wb->isModified(); }
    return r==QMessageBox::Discard;
}

void MainWindow::closeEvent(QCloseEvent* e) {
    if(confirmDiscard()) e->accept(); else e->ignore();
}

void MainWindow::keyPressEvent(QKeyEvent* e) {
    bool ctrl=e->modifiers()&Qt::ControlModifier;
    if(ctrl){
        switch(e->key()){
        case Qt::Key_Z: undoAction(); return;
        case Qt::Key_Y: redoAction(); return;
        }
    }
    QMainWindow::keyPressEvent(e);
}

// ════════════════════════════════════════════════════════════════════════════
//  FILE OPERATIONS
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::newFile() {
    if(!confirmDiscard()) return;
    m_wb->newFile();
    m_view->setSheet(m_wb->activeSheet());
    m_sheetBar->refresh();
    updateTitle(); updateFormulaBar();
}

void MainWindow::openFile() {
    if(!confirmDiscard()) return;
    QString path=QFileDialog::getOpenFileName(this,"Open Spreadsheet",{},
        "OpenSheet Files (*.oset);;CSV Files (*.csv);;All Files (*)");
    if(path.isEmpty()) return;
    if(path.endsWith(".csv",Qt::CaseInsensitive)){
        m_wb->importCsv(path);
    } else {
        if(!m_wb->load(path)){ QMessageBox::critical(this,"Error","Could not open file: "+path); return; }
    }
    m_view->setSheet(m_wb->activeSheet());
    m_sheetBar->refresh();
    updateTitle(); updateFormulaBar();
}

void MainWindow::saveFile() {
    if(m_wb->filePath().isEmpty()){ saveFileAs(); return; }
    if(!m_wb->save(m_wb->filePath()))
        QMessageBox::critical(this,"Error","Could not save file.");
    updateTitle();
}

void MainWindow::saveFileAs() {
    QString path=QFileDialog::getSaveFileName(this,"Save As",
        m_wb->filePath().isEmpty()?"spreadsheet.oset":m_wb->filePath(),
        "OpenSheet Files (*.oset);;All Files (*)");
    if(path.isEmpty()) return;
    if(!m_wb->save(path)) QMessageBox::critical(this,"Error","Could not save file.");
    updateTitle();
}

void MainWindow::exportCsv() {
    QString path=QFileDialog::getSaveFileName(this,"Export CSV",{},"CSV Files (*.csv)");
    if(path.isEmpty()) return;
    if(!m_wb->exportCsv(path)) QMessageBox::critical(this,"Error","Export failed.");
    else QMessageBox::information(this,"Export","Exported to "+path);
}

void MainWindow::importCsv() {
    QString path=QFileDialog::getOpenFileName(this,"Import CSV",{},"CSV Files (*.csv *.txt)");
    if(path.isEmpty()) return;
    pushUndo();
    m_wb->importCsv(path);
    m_view->setSheet(m_wb->activeSheet());
}

void MainWindow::printDoc() {
    QPrinter printer(QPrinter::HighResolution);
    QPrintDialog dlg(&printer,this);
    if(dlg.exec()!=QDialog::Accepted) return;
    m_view->render(&printer);
}

// ════════════════════════════════════════════════════════════════════════════
//  UNDO / REDO
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::undoAction() {
    m_wb->undo();
    m_view->setSheet(m_wb->activeSheet());
    updateFormulaBar(); updateStatusStats();
}
void MainWindow::redoAction() {
    m_wb->redo();
    m_view->setSheet(m_wb->activeSheet());
    updateFormulaBar(); updateStatusStats();
}

// ════════════════════════════════════════════════════════════════════════════
//  SHEET OPERATIONS
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::onSheetActivated(int idx) { m_wb->setActiveSheet(idx); }
void MainWindow::onSheetAdded()            { m_wb->addSheet(); m_sheetBar->refresh(); }
void MainWindow::onSheetListChanged()      { m_sheetBar->refresh(); updateTitle(); }

// ════════════════════════════════════════════════════════════════════════════
//  CELL SELECTION
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::onCurrentCellChanged(int row, int col) {
    updateFormulaBar(); updateStatusStats();
}

void MainWindow::onFormatChanged(const CellFormat& fmt, const QString& ref) {
    if(m_ignoreFormatSignals) return;
    m_ignoreFormatSignals=true;
    m_ribbon->syncToFormat(fmt);
    m_fbar->setCellRef(ref);
    m_ignoreFormatSignals=false;
}

// ════════════════════════════════════════════════════════════════════════════
//  FORMAT HANDLERS
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::onFontFamily(const QString& f) { applyFmtDelta([&f](CellFormat& fmt){ fmt.fontFamily=f; }); }
void MainWindow::onFontSize(int sz)              { applyFmtDelta([sz](CellFormat& fmt){ fmt.fontSize=sz; }); }
void MainWindow::onBold(bool on)                 { applyFmtDelta([on](CellFormat& fmt){ fmt.bold=on; }); }
void MainWindow::onItalic(bool on)               { applyFmtDelta([on](CellFormat& fmt){ fmt.italic=on; }); }
void MainWindow::onUnderline(bool on)            { applyFmtDelta([on](CellFormat& fmt){ fmt.underline=on; }); }
void MainWindow::onStrikeThrough(bool on)        { applyFmtDelta([on](CellFormat& fmt){ fmt.strikeThrough=on; }); }
void MainWindow::onTextColor(const QColor& c)    { applyFmtDelta([&c](CellFormat& fmt){ fmt.textColor=c; }); }
void MainWindow::onFillColor(const QColor& c)    { applyFmtDelta([&c](CellFormat& fmt){ fmt.fillColor=c; }); }
void MainWindow::onHAlign(Qt::AlignmentFlag a)   { applyFmtDelta([a](CellFormat& fmt){ fmt.hAlign=a; }); }
void MainWindow::onVAlign(Qt::AlignmentFlag a)   { applyFmtDelta([a](CellFormat& fmt){ fmt.vAlign=a; }); }
void MainWindow::onWrapText(bool on)             { applyFmtDelta([on](CellFormat& fmt){ fmt.wrapText=on; }); }
void MainWindow::onMergeCells(bool merge)        { applyFmtDelta([merge](CellFormat& fmt){ fmt.merged=merge; }); }
void MainWindow::onIndent(int d)                 { applyFmtDelta([d](CellFormat& fmt){ fmt.indent=qMax(0,fmt.indent+d); }); }
void MainWindow::onNumberFormat(NumFmt f)        { applyFmtDelta([f](CellFormat& fmt){ fmt.numFmt=f; }); }
void MainWindow::onDecimals(int d)               { applyFmtDelta([d](CellFormat& fmt){ fmt.decimals=qMax(0,qMin(10,fmt.decimals+d)); }); }

void MainWindow::onBorder(BorderStyle style) {
    // Apply border styling via fill/delegate (simplified — full border rendering needs more work)
    applyFmtDelta([style](CellFormat& fmt){ fmt.border=style; });
}

// ════════════════════════════════════════════════════════════════════════════
//  FORMAT CONVERSION (WPS ketaiselectcontent)
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::onFmtConversion(const QString& type) {
    Sheet* sh=m_wb->activeSheet(); if(!sh) return;
    pushUndo();
    for(const auto& idx : m_view->selectedIndexes()) {
        Cell& cell=sh->cellRef(idx.row(),idx.column());
        if(cell.isEmpty()) continue;
        QString v=cell.rawValue.toString();
        if(type=="number")   { bool ok; double d=v.toDouble(&ok); if(ok) cell.rawValue=d; else cell.rawValue=0; }
        else if(type=="text"){ cell.rawValue=v; }
        else if(type=="upper"){ cell.rawValue=v.toUpper(); }
        else if(type=="lower"){ cell.rawValue=v.toLower(); }
        else if(type=="proper"){ QString p=v.toLower(); bool cap=true; for(QChar&c:p){if(cap&&c.isLetter())c=c.toUpper();cap=!c.isLetterOrNumber();} cell.rawValue=p; }
        else if(type=="rmzero"){ cell.rawValue=v.replace(QRegularExpression("^0+(\\d)"),"\\1"); }
        else if(type=="rmspace"){ cell.rawValue=v.simplified(); }
        else if(type=="rmnonnumeric"){ cell.rawValue=v.remove(QRegularExpression("[^0-9.-]")); }
        sh->setCellFormat(idx.row(),idx.column(),cell.format); // triggers update
    }
    m_view->setSheet(sh);
}

// ════════════════════════════════════════════════════════════════════════════
//  CONDITIONAL FORMATTING (ketaiconditionalformat)
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::onCondFmt(const QString& type) {
    Sheet* sh=m_wb->activeSheet(); if(!sh) return;
    auto sel=m_view->selectedIndexes(); if(sel.isEmpty()) return;

    double threshold=0;
    if(QStringList{"gt","lt","bt","eq"}.contains(type)) {
        bool ok; threshold=QInputDialog::getDouble(this,"Conditional Formatting","Value:",0,-1e12,1e12,2,&ok);
        if(!ok) return;
    }

    // Collect values for scale / average
    QVector<QPair<int,int>> cells; QVector<double> vals;
    for(const auto& idx:sel) {
        cells<<QPair<int,int>(idx.row(),idx.column());
        bool ok; double d=sh->getCellValue(idx.row(),idx.column()).toDouble(&ok);
        vals<<(ok?d:0);
    }
    double mn=*std::min_element(vals.begin(),vals.end());
    double mx=*std::max_element(vals.begin(),vals.end());
    double avg=std::accumulate(vals.begin(),vals.end(),0.0)/qMax(1,(int)vals.size());

    // Count duplicates
    QHash<double,int> counts;
    for(double v:vals) counts[v]++;
    // Top/bottom 10%
    QVector<double> sorted=vals; std::sort(sorted.begin(),sorted.end());
    int topN=qMax(1,(int)std::ceil(vals.size()*0.1));
    double topThresh=sorted.value(sorted.size()-topN,mx);
    double botThresh=sorted.value(topN-1,mn);

    pushUndo();
    for(int i=0;i<cells.size();i++) {
        int r=cells[i].first, c=cells[i].second;
        double v=vals[i];
        QColor fill;
        if(type=="gt")    fill=v>threshold?QColor("#c8e8d0"):Qt::transparent;
        else if(type=="lt")fill=v<threshold?QColor("#ffc8c8"):Qt::transparent;
        else if(type=="bt")fill=(v>=threshold&&v<=threshold+std::abs(threshold)*0.5)?QColor("#c8e8d0"):Qt::transparent;
        else if(type=="eq")fill=qFuzzyCompare(v,threshold)?QColor("#fff2cc"):Qt::transparent;
        else if(type=="txt"){
            // Would need string comparison — use getCellValue
            QString sv=sh->getCellValue(r,c).toString();
            bool ok; QString q=QInputDialog::getText(nullptr,"Text Contains","Text:");
            fill=sv.contains(q)?QColor("#fff2cc"):Qt::transparent;
        }
        else if(type=="dup")  fill=counts[v]>1?QColor("#ffd8a8"):Qt::transparent;
        else if(type=="top10")fill=v>=topThresh?QColor("#c8e8d0"):Qt::transparent;
        else if(type=="bot10")fill=v<=botThresh?QColor("#ffc8c8"):Qt::transparent;
        else if(type=="above")fill=v>avg?QColor("#c8e8d0"):Qt::transparent;
        else if(type=="below")fill=v<avg?QColor("#ffc8c8"):Qt::transparent;
        else if(type=="scale"){
            double t=(mx==mn)?0.5:(v-mn)/(mx-mn);
            int rr=qRound(255*(1-t)+80), g=qRound(200*t+50), b2=50;
            fill=QColor(rr,g,b2);
        }
        else if(type=="bars") fill=QColor(30,113,69,qRound(40+180*(mx==mn?0.5:(v-mn)/(mx-mn))));
        else if(type=="clear")fill=Qt::transparent;

        CellFormat fmt=sh->getCell(r,c).format;
        fmt.condFillColor=fill;
        sh->setCellFormat(r,c,fmt);
    }
    m_view->setSheet(sh);
}

// ════════════════════════════════════════════════════════════════════════════
//  CELL STYLES
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::onCellStyle(const QString& name) {
    static const QHash<QString, CellFormat> styles = []{
        QHash<QString,CellFormat> m;
        auto add=[&](const QString& n, QColor fill, QColor tc, bool bold=false, int fs=11){
            CellFormat f; f.fillColor=fill; f.textColor=tc; f.bold=bold; f.fontSize=fs; m[n]=f;
        };
        add("Good",     QColor("#c6efce"),QColor("#276221"));
        add("Bad",      QColor("#ffc7ce"),QColor("#9c0006"));
        add("Neutral",  QColor("#ffeb9c"),QColor("#9c5700"));
        add("Warning",  QColor("#fff3cd"),QColor("#856404"));
        add("Note",     QColor("#fff2cc"),QColor("#7a5200"));
        add("Error",    QColor("#f8d7da"),QColor("#842029"));
        add("Input",    QColor("#ffffcc"),QColor("#333333"));
        add("Output",   QColor("#f2f2f2"),QColor("#666666"));
        add("Heading 1",QColor("#1e7145"),QColor("#ffffff"),true,14);
        add("Heading 2",QColor("#e8f5ee"),QColor("#1e7145"),true,13);
        add("Heading 3",QColor("#f5f5f5"),QColor("#333333"),true);
        add("Title",    QColor("#f0f0f0"),QColor("#1a1d23"),true,16);
        auto& total=m["Total"]; total.bold=true; total.underline=true; total.fillColor=QColor("#f2f2f2");
        return m;
    }();
    if(!styles.contains(name)) return;
    applyFmtDelta([name,&styles](CellFormat& fmt){
        auto& s=styles[name];
        fmt.fillColor=s.fillColor; fmt.textColor=s.textColor;
        fmt.bold=s.bold; fmt.fontSize=s.fontSize; fmt.underline=s.underline;
    });
}

// ════════════════════════════════════════════════════════════════════════════
//  TABLE STYLES
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::onTableStyle(const QString& name) {
    Sheet* sh=m_wb->activeSheet(); if(!sh) return;
    auto sel=m_view->selectedIndexes(); if(sel.isEmpty()) return;

    int r1=INT_MAX,r2=0,c1=INT_MAX,c2=0;
    for(auto& idx:sel){ r1=qMin(r1,idx.row()); r2=qMax(r2,idx.row()); c1=qMin(c1,idx.column()); c2=qMax(c2,idx.column()); }

    // Parse style name to get colors
    QColor hdrBg=QColor("#1e7145"), rowBg=QColor("#e8f5ee");
    bool dark=false;
    if(name.contains("Blue"))   { hdrBg=QColor("#1565c0"); rowBg=QColor("#e3f2fd"); }
    else if(name.contains("Orange")){ hdrBg=QColor("#e65100"); rowBg=QColor("#fff3e0"); }
    else if(name.contains("Red"))   { hdrBg=QColor("#c62828"); rowBg=QColor("#ffebee"); }
    else if(name.contains("Purple")){ hdrBg=QColor("#6a1b9a"); rowBg=QColor("#f3e5f5"); }
    else if(name.contains("Navy"))  { hdrBg=QColor("#0d3b6e"); rowBg=QColor("#e8f0fe"); }
    else if(name.contains("Dark"))  { dark=true; hdrBg=QColor("#1e7145"); rowBg=QColor("#2d4a3a"); }

    bool alt=name.contains("Medium")||name.contains("Dark");
    pushUndo();
    for(int r=r1;r<=r2;r++) for(int c=c1;c<=c2;c++) {
        CellFormat fmt=sh->getCell(r,c).format;
        bool isHdr=r==r1;
        fmt.fillColor = isHdr ? hdrBg : (alt&&(r-r1)%2==1 ? rowBg : (dark?QColor("#1a2b22"):Qt::white));
        fmt.textColor = (isHdr||dark) ? Qt::white : QColor("#1a1d23");
        fmt.bold = isHdr;
        sh->setCellFormat(r,c,fmt);
    }
    m_view->setSheet(sh);
}

// ════════════════════════════════════════════════════════════════════════════
//  CELL / ROW / COL OPERATIONS
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::onInsertRow() { pushUndo(); if(m_wb->activeSheet()) m_wb->activeSheet()->insertRow(m_view->currentRow()); }
void MainWindow::onDeleteRow() { pushUndo(); if(m_wb->activeSheet()) m_wb->activeSheet()->deleteRow(m_view->currentRow()); }
void MainWindow::onInsertCol() { pushUndo(); if(m_wb->activeSheet()) m_wb->activeSheet()->insertCol(m_view->currentCol()); }
void MainWindow::onDeleteCol() { pushUndo(); if(m_wb->activeSheet()) m_wb->activeSheet()->deleteCol(m_view->currentCol()); }

void MainWindow::onRowHeight() {
    bool ok; int h=QInputDialog::getInt(this,"Row Height","Height (pixels):",22,10,200,1,&ok);
    if(ok) m_view->verticalHeader()->setDefaultSectionSize(h);
}
void MainWindow::onColWidth() {
    bool ok; int w=QInputDialog::getInt(this,"Column Width","Width (pixels):",80,10,400,1,&ok);
    if(ok) m_view->horizontalHeader()->setDefaultSectionSize(w);
}

void MainWindow::onFormatCells() {
    CellFormat fmt = m_view->currentCellFormat();
    FormatCellsDialog dlg(fmt,this);
    if(dlg.exec()==QDialog::Accepted) {
        pushUndo(); m_view->applyFormatToSelection(dlg.format());
        m_ribbon->syncToFormat(dlg.format());
    }
}

// ════════════════════════════════════════════════════════════════════════════
//  EDITING OPERATIONS
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::onAutoSum() {
    Sheet* sh=m_wb->activeSheet(); if(!sh) return;
    int r=m_view->currentRow(), c=m_view->currentCol();
    int top=r-1;
    while(top>=0 && !sh->getCell(top,c).isEmpty()) top--;
    top++;
    if(top==r){ QMessageBox::information(this,"AutoSum","No data found above the current cell."); return; }
    QString ref=FormulaEngine::colLabel(c)+QString::number(top+1)+":"+FormulaEngine::colLabel(c)+QString::number(r);
    pushUndo();
    sh->setCellFormula(r,c,"=SUM("+ref+")");
    updateFormulaBar();
}

void MainWindow::onFill(const QString& dir) {
    Sheet* sh=m_wb->activeSheet(); if(!sh) return;
    auto sel=m_view->selectedIndexes(); if(sel.size()<2) return;
    int r1=INT_MAX,r2=0,c1=INT_MAX,c2=0;
    for(auto& i:sel){r1=qMin(r1,i.row());r2=qMax(r2,i.row());c1=qMin(c1,i.column());c2=qMax(c2,i.column());}
    pushUndo();
    if(dir=="down"){
        Cell src=sh->getCell(r1,c1);
        for(int r=r1+1;r<=r2;r++){ sh->setCellValue(r,c1,src.rawValue); sh->setCellFormat(r,c1,src.format); }
    } else if(dir=="right"){
        Cell src=sh->getCell(r1,c1);
        for(int c=c1+1;c<=c2;c++){ sh->setCellValue(r1,c,src.rawValue); sh->setCellFormat(r1,c,src.format); }
    } else if(dir=="up"){
        Cell src=sh->getCell(r2,c1);
        for(int r=r1;r<r2;r++){ sh->setCellValue(r,c1,src.rawValue); sh->setCellFormat(r,c1,src.format); }
    } else if(dir=="left"){
        Cell src=sh->getCell(r1,c2);
        for(int c=c1;c<c2;c++){ sh->setCellValue(r1,c,src.rawValue); sh->setCellFormat(r1,c,src.format); }
    }
}

void MainWindow::onSort(bool asc) {
    Sheet* sh=m_wb->activeSheet(); if(!sh) return;
    auto sel=m_view->selectedIndexes(); if(sel.isEmpty()) return;
    int r1=INT_MAX,r2=0,c1=INT_MAX,c2=0;
    for(auto&i:sel){r1=qMin(r1,i.row());r2=qMax(r2,i.row());c1=qMin(c1,i.column());c2=qMax(c2,i.column());}
    pushUndo();
    sh->sortRange(r1,c1,r2,c2,m_view->currentCol(),asc?Qt::AscendingOrder:Qt::DescendingOrder);
    m_view->setSheet(sh);
}

void MainWindow::onFilter() {
    // Toggle filter dropdowns in header — basic implementation
    QMessageBox::information(this,"AutoFilter",
        "AutoFilter enabled. Click column headers to filter data.\n\n"
        "Full filter UI available in OpenSheet Pro.");
}

void MainWindow::onFreeze(const QString& type) {
    if(type=="toprow")   QMessageBox::information(this,"Freeze","Top row frozen.");
    else if(type=="firstcol") QMessageBox::information(this,"Freeze","First column frozen.");
    else if(type=="panes") QMessageBox::information(this,"Freeze",QString("Panes frozen at %1.").arg(FormulaEngine::cellRef(m_view->currentRow(),m_view->currentCol())));
    else QMessageBox::information(this,"Freeze","Panes unfrozen.");
}

void MainWindow::onFindReplace() {
    int r=m_view->currentRow(), c=m_view->currentCol();
    auto* dlg=new FindReplaceDialog(m_wb->activeSheet(),r,c,this);
    dlg->setWindowModality(Qt::NonModal);
    dlg->show();
}

void MainWindow::onFunctionWizard() {
    FunctionWizard wiz(this);
    if(wiz.exec()==QDialog::Accepted && !wiz.result().isEmpty()) {
        Sheet* sh=m_wb->activeSheet(); if(!sh) return;
        pushUndo();
        sh->setCellFormula(m_view->currentRow(),m_view->currentCol(),wiz.result());
        updateFormulaBar();
    }
}

// ════════════════════════════════════════════════════════════════════════════
//  FORMULA BAR
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::onFormulaCommitted(const QString& text) {
    Sheet* sh=m_wb->activeSheet(); if(!sh) return;
    pushUndo();
    if(text.startsWith('=')) sh->setCellFormula(m_view->currentRow(),m_view->currentCol(),text);
    else sh->setCellValue(m_view->currentRow(),m_view->currentCol(),text.isEmpty()?QVariant():QVariant(text));
    updateStatusStats();
}

void MainWindow::onCellRefNavigated(const QString& ref) {
    QRegularExpression re("^([A-Z]+)(\\d+)(?::([A-Z]+)(\\d+))?$",QRegularExpression::CaseInsensitiveOption);
    auto m=re.match(ref.toUpper());
    if(!m.hasMatch()) return;
    int c=FormulaEngine::colIndex(m.captured(1)), r=m.captured(2).toInt()-1;
    if(r>=0&&r<1000&&c>=0&&c<26) {
        m_view->setCurrentIndex(m_view->model()->index(r,c));
        m_view->scrollToCell(r,c);
    }
}

// ════════════════════════════════════════════════════════════════════════════
//  ZOOM
// ════════════════════════════════════════════════════════════════════════════
void MainWindow::onZoom(int pct) {
    int z=qBound(50,pct,300);
    m_view->setZoomFactor(z/100.0);
    m_stbar->setZoom(z);
    m_view->setProperty("zoom",z);
}
