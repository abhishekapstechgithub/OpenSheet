// ═══════════════════════════════════════════════════════════════════════════════
//  SpreadsheetView.cpp + CellDelegate — WPS ET style grid
// ═══════════════════════════════════════════════════════════════════════════════
#include "SpreadsheetView.h"
#include "FormulaEngine.h"
#include <QHeaderView>
#include <QScrollBar>
#include <QApplication>
#include <QClipboard>
#include <QKeyEvent>
#include <QWheelEvent>
#include <QContextMenuEvent>
#include <QMouseEvent>
#include <QMenu>
#include <QLineEdit>
#include <QPainter>
#include <QBrush>
#include <QFont>
#include <QColor>

// ════════════════════════════════════════════════════════════════════════════
//  SheetModel
// ════════════════════════════════════════════════════════════════════════════
SheetModel::SheetModel(Sheet* sheet, QObject* parent)
    : QAbstractTableModel(parent), m_sheet(sheet)
{
    connect(sheet, &Sheet::cellChanged, this, [this](int r, int c){
        auto idx = index(r, c);
        emit dataChanged(idx, idx);
    });
    connect(sheet, &Sheet::sheetChanged, this, [this]{
        beginResetModel(); endResetModel();
    });
}

void SheetModel::setSheet(Sheet* sheet) {
    beginResetModel();
    if(m_sheet) disconnect(m_sheet, nullptr, this, nullptr);
    m_sheet = sheet;
    if(m_sheet) {
        connect(m_sheet, &Sheet::cellChanged, this, [this](int r, int c){
            auto idx=index(r,c); emit dataChanged(idx,idx);
        });
        connect(m_sheet, &Sheet::sheetChanged, this, [this]{ beginResetModel(); endResetModel(); });
    }
    endResetModel();
}

int SheetModel::rowCount(const QModelIndex& parent) const {
    if(parent.isValid()) return 0;
    return m_sheet ? qMax(200, m_sheet->usedRowCount()+50) : 200;
}

int SheetModel::columnCount(const QModelIndex& parent) const {
    if(parent.isValid()) return 0;
    return m_sheet ? qMax(26, m_sheet->usedColCount()+5) : 26;
}

QVariant SheetModel::data(const QModelIndex& idx, int role) const {
    if(!idx.isValid() || !m_sheet) return {};
    int r=idx.row(), c=idx.column();
    const Cell& cell = m_sheet->getCell(r,c);

    switch(role) {
    case Qt::DisplayRole: {
        if(cell.isEmpty()) return {};
        QVariant val = m_sheet->getCellValue(r,c);
        return FormulaEngine::formatValue(val, (int)cell.format.numFmt, cell.format.decimals);
    }
    case Qt::EditRole:
        return cell.rawValue;
    case Qt::FontRole: {
        QFont f(cell.format.fontFamily, cell.format.fontSize);
        f.setBold(cell.format.bold); f.setItalic(cell.format.italic);
        f.setUnderline(cell.format.underline); f.setStrikeOut(cell.format.strikeThrough);
        return f;
    }
    case Qt::ForegroundRole:
        return cell.format.textColor.isValid() ? QBrush(cell.format.textColor) : QBrush(QColor("#1a1d23"));
    case Qt::BackgroundRole: {
        if(cell.format.hasCondFill()) return QBrush(cell.format.condFillColor);
        if(cell.format.fillColor.isValid() && cell.format.fillColor != Qt::transparent)
            return QBrush(cell.format.fillColor);
        return {};
    }
    case Qt::TextAlignmentRole: {
        int align = (int)cell.format.hAlign | (int)cell.format.vAlign;
        if(align==0) {
            // Auto: numbers right, text left
            if(!cell.isEmpty()){
                bool ok; m_sheet->getCellValue(r,c).toDouble(&ok);
                align = (ok?(int)Qt::AlignRight:(int)Qt::AlignLeft) | (int)Qt::AlignVCenter;
            } else align=(int)Qt::AlignLeft|(int)Qt::AlignVCenter;
        }
        return align;
    }
    case Qt::UserRole:  // raw formula
        return cell.rawValue;
    case Qt::UserRole+1: // is formula?
        return cell.isFormula();
    }
    return {};
}

bool SheetModel::setData(const QModelIndex& idx, const QVariant& value, int role) {
    if(!idx.isValid()||!m_sheet||role!=Qt::EditRole) return false;
    QString s=value.toString();
    if(s.startsWith('='))
        m_sheet->setCellFormula(idx.row(), idx.column(), s);
    else
        m_sheet->setCellValue(idx.row(), idx.column(), s.isEmpty()?QVariant():QVariant(s));
    return true;
}

Qt::ItemFlags SheetModel::flags(const QModelIndex& idx) const {
    if(!idx.isValid()) return Qt::NoItemFlags;
    return Qt::ItemIsEnabled|Qt::ItemIsSelectable|Qt::ItemIsEditable;
}

QVariant SheetModel::headerData(int section, Qt::Orientation orientation, int role) const {
    if(role==Qt::DisplayRole) {
        if(orientation==Qt::Horizontal) return columnLabel(section);
        return QString::number(section+1);
    }
    if(role==Qt::TextAlignmentRole) {
        if(orientation==Qt::Horizontal) return (int)(Qt::AlignHCenter|Qt::AlignVCenter);
        return (int)(Qt::AlignRight|Qt::AlignVCenter);
    }
    return {};
}

// ════════════════════════════════════════════════════════════════════════════
//  CellDelegate
// ════════════════════════════════════════════════════════════════════════════
CellDelegate::CellDelegate(QObject* parent) : QStyledItemDelegate(parent) {}

void CellDelegate::paint(QPainter* p, const QStyleOptionViewItem& opt, const QModelIndex& idx) const {
    p->save();

    // Background
    QColor bg = idx.data(Qt::BackgroundRole).value<QBrush>().color();
    if(!bg.isValid() || bg.alpha()==0) {
        bg = (opt.state & QStyle::State_Selected) ? QColor("#b8d9c8") : Qt::white;
    } else if(opt.state & QStyle::State_Selected) {
        bg = bg.lighter(115);
    }
    p->fillRect(opt.rect, bg);

    // Grid lines — WPS style: 0.7px #d8dce3
    p->setPen(QPen(QColor("#d8dce3"), 0.7));
    p->drawLine(opt.rect.bottomLeft(), opt.rect.bottomRight());
    p->drawLine(opt.rect.topRight(), opt.rect.bottomRight());

    // Text
    QString text = idx.data(Qt::DisplayRole).toString();
    if(!text.isEmpty()) {
        QFont font = idx.data(Qt::FontRole).value<QFont>();
        p->setFont(font);
        QColor fg = idx.data(Qt::ForegroundRole).value<QBrush>().color();
        if(!fg.isValid()) fg = QColor("#1a1d23");
        // Red for error values
        if(text.startsWith('#')) fg = QColor("#d32f2f");
        // Green for formulas
        if(idx.data(Qt::UserRole+1).toBool() && !text.startsWith('#')) fg = QColor("#0a5c35");
        p->setPen(fg);
        int align = idx.data(Qt::TextAlignmentRole).toInt();
        QRect textRect = opt.rect.adjusted(4,1,-4,-1);
        p->drawText(textRect, align, text);
    }

    // Active cell: draw inner border (handled by view selection highlight)
    p->restore();
}

QSize CellDelegate::sizeHint(const QStyleOptionViewItem&, const QModelIndex&) const {
    return QSize(80, 22);
}

QWidget* CellDelegate::createEditor(QWidget* parent, const QStyleOptionViewItem&, const QModelIndex&) const {
    auto* ed = new QLineEdit(parent);
    ed->setFrame(false);
    ed->setStyleSheet("QLineEdit{background:white;font-family:'Courier New';font-size:13px;padding:0 3px;}");
    return ed;
}

void CellDelegate::setEditorData(QWidget* editor, const QModelIndex& index) const {
    auto* ed = qobject_cast<QLineEdit*>(editor);
    if(ed) ed->setText(index.data(Qt::EditRole).toString());
}

void CellDelegate::setModelData(QWidget* editor, QAbstractItemModel* model, const QModelIndex& index) const {
    auto* ed = qobject_cast<QLineEdit*>(editor);
    if(ed) model->setData(index, ed->text(), Qt::EditRole);
}

// ════════════════════════════════════════════════════════════════════════════
//  SpreadsheetView
// ════════════════════════════════════════════════════════════════════════════
SpreadsheetView::SpreadsheetView(Sheet* sheet, QWidget* parent)
    : QTableView(parent)
{
    m_model = new SheetModel(sheet, this);
    setModel(m_model);
    setItemDelegate(new CellDelegate(this));
    setupAppearance();
    connect(selectionModel(), &QItemSelectionModel::currentChanged,
            this, &SpreadsheetView::onCurrentChanged);
}

void SpreadsheetView::setupAppearance() {
    // WPS ET grid appearance
    setShowGrid(false);        // We draw our own grid in the delegate
    setAlternatingRowColors(false);
    setSelectionMode(ExtendedSelection);
    setSelectionBehavior(SelectItems);
    setEditTriggers(DoubleClicked | AnyKeyPressed);
    setTabKeyNavigation(false);
    setHorizontalScrollMode(ScrollPerPixel);
    setVerticalScrollMode(ScrollPerPixel);
    setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);
    setMinimumSize(200,150);

    setStyleSheet(
        // Main table
        "QTableView{"
        "  background:white; selection-background-color:rgba(30,113,69,0.12);"
        "  selection-color:#1a1d23; border:none; outline:none;"
        "}"
        // Column headers — WPS #f2f4f8
        "QHeaderView::section:horizontal{"
        "  background:#f2f4f8; color:#666; font-weight:600; font-size:11px;"
        "  font-family:'Segoe UI',Arial; border:none;"
        "  border-right:1px solid #d0d5de; border-bottom:2px solid #c8cdd6;"
        "  padding:0 4px; min-height:22px;"
        "}"
        "QHeaderView::section:horizontal:checked{background:#1e7145;color:white;}"
        // Row headers
        "QHeaderView::section:vertical{"
        "  background:#f2f4f8; color:#999; font-size:11px;"
        "  font-family:'Segoe UI',Arial; border:none;"
        "  border-right:2px solid #c8cdd6; border-bottom:1px solid #d4d8e0;"
        "  padding:0 5px 0 0; min-height:22px;"
        "}"
        "QHeaderView::section:vertical:checked{background:#1e7145;color:white;}"
        // Corner button
        "QAbstractButton{background:#f2f4f8;border-right:2px solid #c8cdd6;border-bottom:2px solid #c8cdd6;}"
        // Scrollbars
        "QScrollBar:horizontal,QScrollBar:vertical{background:#f5f6f7;border:none;}"
        "QScrollBar:horizontal{height:8px;} QScrollBar:vertical{width:8px;}"
        "QScrollBar::handle{background:#c4c9d4;border-radius:4px;margin:1px;}"
        "QScrollBar::handle:hover{background:#7a8494;}"
        "QScrollBar::add-line,QScrollBar::sub-line{width:0;height:0;}"
    );

    // Header setup
    horizontalHeader()->setDefaultSectionSize(m_baseColW);
    horizontalHeader()->setMinimumSectionSize(8);
    horizontalHeader()->setSectionsMovable(false);
    horizontalHeader()->setHighlightSections(true);
    horizontalHeader()->setDefaultAlignment(Qt::AlignHCenter|Qt::AlignVCenter);
    horizontalHeader()->setFixedHeight(24);

    verticalHeader()->setDefaultSectionSize(m_baseRowH);
    verticalHeader()->setFixedWidth(46);
    verticalHeader()->setMinimumSectionSize(8);
    verticalHeader()->setHighlightSections(true);
    verticalHeader()->setDefaultAlignment(Qt::AlignRight|Qt::AlignVCenter);
}

void SpreadsheetView::setSheet(Sheet* sheet) {
    m_model->setSheet(sheet);
    scrollTo(m_model->index(0,0));
    setCurrentIndex(m_model->index(0,0));
}

Sheet* SpreadsheetView::sheet() const { return m_model->sheet(); }
int SpreadsheetView::currentRow() const { return currentIndex().row(); }
int SpreadsheetView::currentCol() const { return currentIndex().column(); }

void SpreadsheetView::scrollToCell(int r, int c) {
    scrollTo(m_model->index(r,c), EnsureVisible);
}

void SpreadsheetView::setZoomFactor(qreal f) {
    m_zoom = qBound(0.5, f, 3.0);
    horizontalHeader()->setDefaultSectionSize(qRound(m_baseColW * m_zoom));
    verticalHeader()->setDefaultSectionSize(qRound(m_baseRowH * m_zoom));
    QFont fnt; fnt.setFamily("Segoe UI"); fnt.setPixelSize(qMax(8,qRound(13*m_zoom)));
    setFont(fnt); horizontalHeader()->setFont(fnt); verticalHeader()->setFont(fnt);
    emit zoomChanged(qRound(m_zoom*100));
}

CellFormat SpreadsheetView::currentCellFormat() const {
    if(!currentIndex().isValid()||!m_model->sheet()) return {};
    return m_model->sheet()->getCell(currentRow(),currentCol()).format;
}

void SpreadsheetView::applyFormatToSelection(const CellFormat& fmt) {
    if(!m_model->sheet()) return;
    for(const auto& idx : selectedIndexes()) {
        Cell& cell = m_model->sheet()->cellRef(idx.row(), idx.column());
        cell.format = fmt;
        m_model->sheet()->setCellFormat(idx.row(), idx.column(), fmt);
    }
}

void SpreadsheetView::applyFormatDelta(std::function<void(CellFormat&)> fn) {
    if(!m_model->sheet()) return;
    for(const auto& idx : selectedIndexes()) {
        CellFormat fmt = m_model->sheet()->getCell(idx.row(),idx.column()).format;
        fn(fmt);
        m_model->sheet()->setCellFormat(idx.row(),idx.column(),fmt);
    }
}

void SpreadsheetView::onCurrentChanged(const QModelIndex& cur, const QModelIndex&) {
    if(!cur.isValid()||!m_model->sheet()) return;
    CellFormat fmt = m_model->sheet()->getCell(cur.row(),cur.column()).format;
    QString ref = FormulaEngine::cellRef(cur.row(), cur.column());
    emit formatChanged(fmt, ref);
    emit currentCellChanged(cur.row(), cur.column());
}

void SpreadsheetView::keyPressEvent(QKeyEvent* e) {
    bool ctrl = e->modifiers() & Qt::ControlModifier;
    if(ctrl) {
        switch(e->key()) {
        case Qt::Key_C: {
            // Copy to clipboard
            QStringList rows; int lr=-1; QStringList row;
            for(const auto& idx : selectedIndexes()) {
                if(lr!=-1 && idx.row()!=lr){ rows<<row.join('\t'); row.clear(); }
                row<<m_model->data(idx,Qt::DisplayRole).toString();
                lr=idx.row();
            }
            if(!row.isEmpty()) rows<<row.join('\t');
            QApplication::clipboard()->setText(rows.join('\n'));
            return;
        }
        case Qt::Key_V: {
            QString text=QApplication::clipboard()->text();
            QStringList lines=text.split('\n');
            int sr=currentRow(),sc=currentCol();
            if(m_model->sheet()){
                for(int r=0;r<lines.size();r++){
                    QStringList cells=lines[r].split('\t');
                    for(int c=0;c<cells.size();c++)
                        m_model->sheet()->setCellValue(sr+r,sc+c,cells[c]);
                }
            }
            return;
        }
        case Qt::Key_Z: // Undo handled in MainWindow
        case Qt::Key_Y: // Redo handled in MainWindow
            e->ignore(); return;
        }
    }
    if(e->key()==Qt::Key_Delete||e->key()==Qt::Key_Backspace) {
        if(m_model->sheet())
            for(const auto& idx : selectedIndexes())
                m_model->sheet()->clearCell(idx.row(),idx.column());
        return;
    }
    QTableView::keyPressEvent(e);
}

void SpreadsheetView::wheelEvent(QWheelEvent* e) {
    if(e->modifiers() & Qt::ControlModifier) {
        setZoomFactor(m_zoom + (e->angleDelta().y()>0?0.1:-0.1));
        e->accept(); return;
    }
    QTableView::wheelEvent(e);
}

void SpreadsheetView::contextMenuEvent(QContextMenuEvent* e) {
    QMenu menu(this);
    menu.setStyleSheet(
        "QMenu{background:white;border:1px solid #d8dce3;border-radius:4px;padding:3px 0;"
        "font-family:'Segoe UI';font-size:12px;}"
        "QMenu::item{padding:6px 20px 6px 14px;}"
        "QMenu::item:selected{background:#e8f5ee;color:#1e7145;}"
        "QMenu::separator{height:1px;background:#d8dce3;margin:3px 0;}");
    menu.addAction("Cut",    [this]{ keyPressEvent(new QKeyEvent(QEvent::KeyPress,Qt::Key_X,Qt::ControlModifier)); });
    menu.addAction("Copy",   [this]{ keyPressEvent(new QKeyEvent(QEvent::KeyPress,Qt::Key_C,Qt::ControlModifier)); });
    menu.addAction("Paste",  [this]{ keyPressEvent(new QKeyEvent(QEvent::KeyPress,Qt::Key_V,Qt::ControlModifier)); });
    menu.addSeparator();
    menu.addAction("Insert Row",    [this]{ if(m_model->sheet()) m_model->sheet()->insertRow(currentRow()); });
    menu.addAction("Delete Row",    [this]{ if(m_model->sheet()) m_model->sheet()->deleteRow(currentRow()); });
    menu.addAction("Insert Column", [this]{ if(m_model->sheet()) m_model->sheet()->insertCol(currentCol()); });
    menu.addAction("Delete Column", [this]{ if(m_model->sheet()) m_model->sheet()->deleteCol(currentCol()); });
    menu.addSeparator();
    menu.addAction("Clear Contents",[this]{
        if(m_model->sheet()) for(const auto& idx:selectedIndexes()) m_model->sheet()->clearCell(idx.row(),idx.column());
    });
    menu.exec(e->globalPos());
}

void SpreadsheetView::mousePressEvent(QMouseEvent* e) {
    QTableView::mousePressEvent(e);
}
