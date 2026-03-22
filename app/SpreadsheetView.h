#pragma once
#include <QTableView>
#include <QAbstractTableModel>
#include <QStyledItemDelegate>
#include "Sheet.h"
#include "FormulaEngine.h"

// ─── Table Model ──────────────────────────────────────────────────────────────
class SheetModel : public QAbstractTableModel {
    Q_OBJECT
public:
    explicit SheetModel(Sheet* sheet, QObject* parent=nullptr);
    void setSheet(Sheet* sheet);

    int rowCount(const QModelIndex& parent={}) const override;
    int columnCount(const QModelIndex& parent={}) const override;
    QVariant data(const QModelIndex& index, int role=Qt::DisplayRole) const override;
    bool setData(const QModelIndex& index, const QVariant& value, int role=Qt::EditRole) override;
    Qt::ItemFlags flags(const QModelIndex& index) const override;
    QVariant headerData(int section, Qt::Orientation orientation, int role=Qt::DisplayRole) const override;

    static QString columnLabel(int col) { return FormulaEngine::colLabel(col); }
    Sheet* sheet() const { return m_sheet; }

private:
    Sheet* m_sheet;
};

// ─── Cell Delegate ────────────────────────────────────────────────────────────
class CellDelegate : public QStyledItemDelegate {
    Q_OBJECT
public:
    explicit CellDelegate(QObject* parent=nullptr);
    void paint(QPainter* painter, const QStyleOptionViewItem& option, const QModelIndex& index) const override;
    QSize sizeHint(const QStyleOptionViewItem& option, const QModelIndex& index) const override;
    QWidget* createEditor(QWidget* parent, const QStyleOptionViewItem& option, const QModelIndex& index) const override;
    void setEditorData(QWidget* editor, const QModelIndex& index) const override;
    void setModelData(QWidget* editor, QAbstractItemModel* model, const QModelIndex& index) const override;
};

// ─── SpreadsheetView ──────────────────────────────────────────────────────────
class SpreadsheetView : public QTableView {
    Q_OBJECT
public:
    explicit SpreadsheetView(Sheet* sheet, QWidget* parent=nullptr);

    void   setSheet(Sheet* sheet);
    Sheet* sheet() const;
    SheetModel* sheetModel() const { return m_model; }

    int  currentRow() const;
    int  currentCol() const;
    void selectAll();
    void selectRow(int row);
    void selectCol(int col);
    void scrollToCell(int row, int col);
    void setZoomFactor(qreal factor);

    CellFormat currentCellFormat() const;
    void applyFormatToSelection(const CellFormat& fmt);
    void applyFormatDelta(std::function<void(CellFormat&)> fn);

signals:
    void currentCellChanged(int row, int col);
    void formatChanged(const CellFormat& fmt, const QString& ref);
    void zoomChanged(int percent);

protected:
    void keyPressEvent(QKeyEvent* e) override;
    void wheelEvent(QWheelEvent* e) override;
    void contextMenuEvent(QContextMenuEvent* e) override;
    void mousePressEvent(QMouseEvent* e) override;

private slots:
    void onCurrentChanged(const QModelIndex& cur, const QModelIndex& prev);

private:
    SheetModel* m_model;
    int         m_baseRowH { 22 };
    int         m_baseColW { 80 };
    qreal       m_zoom     { 1.0 };
    void setupAppearance();
};
