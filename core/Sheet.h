#pragma once
#include "Cell.h"
#include "FormulaEngine.h"
#include <QString>
#include <QHash>
#include <QVariant>

class Sheet : public QObject {
    Q_OBJECT
public:
    explicit Sheet(const QString& name = "Sheet1", QObject* parent = nullptr);

    // ── Identity ──────────────────────────────────────────────────────────────
    QString name() const { return m_name; }
    void    setName(const QString& n) { m_name = n; }
    bool    hidden()  const { return m_hidden; }
    void    setHidden(bool h) { m_hidden = h; }

    // ── Cell access ───────────────────────────────────────────────────────────
    const Cell& getCell(int row, int col) const;
    Cell&       cellRef(int row, int col);
    void        setCellValue(int row, int col, const QVariant& value);
    void        setCellFormula(int row, int col, const QString& formula);
    void        setCellFormat(int row, int col, const CellFormat& fmt);
    void        clearCell(int row, int col);
    void        clearCells(int r1,int c1,int r2,int c2);

    // Evaluated value (formula resolved)
    QVariant getCellValue(int row, int col) const;

    // ── Dimensions ────────────────────────────────────────────────────────────
    int usedRowCount() const;
    int usedColCount() const;

    // ── Row / column ops ─────────────────────────────────────────────────────
    void insertRow(int row);
    void deleteRow(int row);
    void insertCol(int col);
    void deleteCol(int col);

    // ── Sort ─────────────────────────────────────────────────────────────────
    void sortRange(int r1,int c1,int r2,int c2,int keyCol,Qt::SortOrder order);

    // ── Serialisation ────────────────────────────────────────────────────────
    QByteArray  toJson() const;
    static Sheet* fromJson(const QByteArray& data, QObject* parent = nullptr);

signals:
    void cellChanged(int row, int col);
    void sheetChanged();

private:
    QString m_name;
    bool    m_hidden { false };
    QHash<quint64, Cell> m_cells;  // key = row<<20 | col
    mutable Cell m_empty;

    static quint64 key(int r,int c){ return ((quint64)r << 20) | (quint64)c; }
};
