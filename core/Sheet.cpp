#include "Sheet.h"
#include <QJsonObject>
#include <QJsonArray>
#include <QJsonDocument>
#include <algorithm>

Sheet::Sheet(const QString& name, QObject* parent)
    : QObject(parent), m_name(name) {}

const Cell& Sheet::getCell(int row, int col) const {
    auto it = m_cells.constFind(key(row,col));
    return it != m_cells.constEnd() ? it.value() : m_empty;
}

Cell& Sheet::cellRef(int row, int col) {
    return m_cells[key(row,col)];
}

QVariant Sheet::getCellValue(int row, int col) const {
    const Cell& cell = getCell(row,col);
    if (cell.isEmpty()) return QVariant();
    if (cell.isFormula()) {
        return FormulaEngine::evaluate(cell.rawValue.toString(),
                                        const_cast<Sheet*>(this), row, col);
    }
    return cell.rawValue;
}

void Sheet::setCellValue(int row, int col, const QVariant& value) {
    auto& cell = cellRef(row, col);
    cell.rawValue = value;
    cell.dirty    = true;
    emit cellChanged(row, col);
}

void Sheet::setCellFormula(int row, int col, const QString& formula) {
    auto& cell = cellRef(row, col);
    cell.rawValue = formula;
    cell.dirty    = true;
    emit cellChanged(row, col);
}

void Sheet::setCellFormat(int row, int col, const CellFormat& fmt) {
    auto& cell = cellRef(row, col);
    cell.format = fmt;
    emit cellChanged(row, col);
}

void Sheet::clearCell(int row, int col) {
    m_cells.remove(key(row,col));
    emit cellChanged(row, col);
}

void Sheet::clearCells(int r1,int c1,int r2,int c2) {
    for (int r=r1;r<=r2;r++) for (int c=c1;c<=c2;c++) m_cells.remove(key(r,c));
    emit sheetChanged();
}

int Sheet::usedRowCount() const {
    int max = 0;
    for (auto it=m_cells.constBegin();it!=m_cells.constEnd();++it)
        max = qMax(max, (int)(it.key()>>20)+1);
    return qMax(200, max+50);
}

int Sheet::usedColCount() const {
    int max = 0;
    for (auto it=m_cells.constBegin();it!=m_cells.constEnd();++it)
        max = qMax(max, (int)(it.key()&0xFFFFF)+1);
    return qMax(26, max+5);
}

void Sheet::insertRow(int row) {
    QHash<quint64,Cell> newCells;
    for (auto it=m_cells.begin();it!=m_cells.end();++it) {
        int r=(int)(it.key()>>20), c=(int)(it.key()&0xFFFFF);
        newCells[key(r>=row?r+1:r, c)] = it.value();
    }
    m_cells = newCells;
    emit sheetChanged();
}

void Sheet::deleteRow(int row) {
    QHash<quint64,Cell> newCells;
    for (auto it=m_cells.begin();it!=m_cells.end();++it) {
        int r=(int)(it.key()>>20), c=(int)(it.key()&0xFFFFF);
        if (r==row) continue;
        newCells[key(r>row?r-1:r, c)] = it.value();
    }
    m_cells = newCells;
    emit sheetChanged();
}

void Sheet::insertCol(int col) {
    QHash<quint64,Cell> newCells;
    for (auto it=m_cells.begin();it!=m_cells.end();++it) {
        int r=(int)(it.key()>>20), c=(int)(it.key()&0xFFFFF);
        newCells[key(r, c>=col?c+1:c)] = it.value();
    }
    m_cells = newCells;
    emit sheetChanged();
}

void Sheet::deleteCol(int col) {
    QHash<quint64,Cell> newCells;
    for (auto it=m_cells.begin();it!=m_cells.end();++it) {
        int r=(int)(it.key()>>20), c=(int)(it.key()&0xFFFFF);
        if (c==col) continue;
        newCells[key(r, c>col?c-1:c)] = it.value();
    }
    m_cells = newCells;
    emit sheetChanged();
}

void Sheet::sortRange(int r1,int c1,int r2,int c2,int keyCol,Qt::SortOrder order) {
    // Collect rows
    struct Row { QHash<int,Cell> cells; QVariant key; };
    QVector<Row> rows;
    for (int r=r1;r<=r2;r++) {
        Row row;
        for (int c=c1;c<=c2;c++) {
            auto it=m_cells.constFind(key(r,c));
            if (it!=m_cells.constEnd()) row.cells[c]=it.value();
        }
        row.key = getCellValue(r,keyCol);
        rows << row;
    }
    std::stable_sort(rows.begin(),rows.end(),[&](const Row&a,const Row&b){
        bool aOk,bOk;
        double an=a.key.toDouble(&aOk), bn=b.key.toDouble(&bOk);
        int cmp = (aOk&&bOk) ? (an<bn?-1:an>bn?1:0) : a.key.toString().compare(b.key.toString());
        return order==Qt::AscendingOrder ? cmp<0 : cmp>0;
    });
    for (int i=0;i<rows.size();i++) {
        for (int c=c1;c<=c2;c++) m_cells.remove(key(r1+i,c));
        for (auto& [c,cell] : rows[i].cells.asKeyValueRange())
            m_cells[key(r1+i,c)] = cell;
    }
    emit sheetChanged();
}

QByteArray Sheet::toJson() const {
    QJsonArray arr;
    for (auto it=m_cells.constBegin();it!=m_cells.constEnd();++it) {
        int r=(int)(it.key()>>20), c=(int)(it.key()&0xFFFFF);
        const Cell& cell=it.value();
        if (cell.isEmpty() && cell.format.bold==false) continue;
        QJsonObject obj;
        obj["r"]=r; obj["c"]=c;
        obj["v"]=cell.rawValue.toString();
        QJsonObject fmt;
        fmt["ff"]=cell.format.fontFamily; fmt["fs"]=cell.format.fontSize;
        fmt["b"]=(int)cell.format.bold; fmt["i"]=(int)cell.format.italic;
        fmt["u"]=(int)cell.format.underline; fmt["s"]=(int)cell.format.strikeThrough;
        fmt["w"]=(int)cell.format.wrapText; fmt["in"]=cell.format.indent;
        fmt["tc"]=cell.format.textColor.name(); fmt["fc"]=cell.format.fillColor.name();
        fmt["ha"]=(int)cell.format.hAlign; fmt["va"]=(int)cell.format.vAlign;
        fmt["nf"]=(int)cell.format.numFmt; fmt["dc"]=cell.format.decimals;
        obj["f"]=fmt;
        arr.append(obj);
    }
    QJsonObject root; root["name"]=m_name; root["hidden"]=(int)m_hidden; root["cells"]=arr;
    return QJsonDocument(root).toJson(QJsonDocument::Compact);
}

Sheet* Sheet::fromJson(const QByteArray& data, QObject* parent) {
    auto doc = QJsonDocument::fromJson(data);
    if (!doc.isObject()) return nullptr;
    QJsonObject root = doc.object();
    auto* sh = new Sheet(root["name"].toString("Sheet1"), parent);
    sh->m_hidden = root["hidden"].toInt()!=0;
    for (const QJsonValue& v : root["cells"].toArray()) {
        QJsonObject o=v.toObject();
        int r=o["r"].toInt(), c=o["c"].toInt();
        sh->m_cells[key(r,c)].rawValue = o["v"].toString();
        QJsonObject fmt=o["f"].toObject();
        auto& cf=sh->m_cells[key(r,c)].format;
        cf.fontFamily   = fmt["ff"].toString("Calibri");
        cf.fontSize     = fmt["fs"].toInt(11);
        cf.bold         = fmt["b"].toInt()!=0;
        cf.italic       = fmt["i"].toInt()!=0;
        cf.underline    = fmt["u"].toInt()!=0;
        cf.strikeThrough= fmt["s"].toInt()!=0;
        cf.wrapText     = fmt["w"].toInt()!=0;
        cf.indent       = fmt["in"].toInt(0);
        cf.textColor    = QColor(fmt["tc"].toString("#000000"));
        cf.fillColor    = QColor(fmt["fc"].toString("transparent"));
        cf.hAlign       = (Qt::AlignmentFlag)fmt["ha"].toInt(Qt::AlignLeft);
        cf.vAlign       = (Qt::AlignmentFlag)fmt["va"].toInt(Qt::AlignVCenter);
        cf.numFmt       = (NumFmt)fmt["nf"].toInt(0);
        cf.decimals     = fmt["dc"].toInt(2);
    }
    return sh;
}
