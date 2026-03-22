#include "Workbook.h"
#include <QFile>
#include <QTextStream>
#include <QJsonObject>
#include <QJsonArray>
#include <QJsonDocument>
#include <QFileInfo>

Workbook::Workbook(QObject* parent) : QObject(parent) {
    newFile();
}
Workbook::~Workbook() { qDeleteAll(m_sheets); }

void Workbook::newFile() {
    qDeleteAll(m_sheets); m_sheets.clear();
    m_undoStack.clear(); m_redoStack.clear();
    m_filePath.clear(); m_modified=false;
    m_activeIdx=0;
    addSheet("Sheet1"); addSheet("Sheet2"); addSheet("Sheet3");
    emit sheetListChanged();
    emit activeSheetChanged(0);
}

void Workbook::setActiveSheet(int idx) {
    if(idx>=0&&idx<m_sheets.size()&&idx!=m_activeIdx){
        m_activeIdx=idx; emit activeSheetChanged(idx);
    }
}

Sheet* Workbook::addSheet(const QString& name) {
    QString n = name.isEmpty() ? QString("Sheet%1").arg(m_sheets.size()+1) : name;
    auto* sh = new Sheet(n, this);
    connect(sh,&Sheet::cellChanged,this,[this](int,int){ setModified(true); });
    connect(sh,&Sheet::sheetChanged,this,[this]{ setModified(true); });
    m_sheets << sh;
    emit sheetListChanged();
    return sh;
}

void Workbook::removeSheet(int idx) {
    if(m_sheets.size()<=1) return;
    delete m_sheets.takeAt(idx);
    if(m_activeIdx>=m_sheets.size()) m_activeIdx=m_sheets.size()-1;
    emit sheetListChanged();
    emit activeSheetChanged(m_activeIdx);
}

void Workbook::renameSheet(int idx, const QString& name) {
    if(idx>=0&&idx<m_sheets.size()){ m_sheets[idx]->setName(name); emit sheetListChanged(); }
}

void Workbook::moveSheet(int from, int to) {
    if(from<0||from>=m_sheets.size()||to<0||to>=m_sheets.size()) return;
    m_sheets.move(from,to);
    m_activeIdx = to;
    emit sheetListChanged();
}

void Workbook::duplicateSheet(int idx) {
    if(idx<0||idx>=m_sheets.size()) return;
    auto* orig = m_sheets[idx];
    auto* dup  = Sheet::fromJson(orig->toJson(), this);
    if(!dup) return;
    dup->setName(orig->name() + " (2)");
    connect(dup,&Sheet::cellChanged,this,[this](int,int){ setModified(true); });
    connect(dup,&Sheet::sheetChanged,this,[this]{ setModified(true); });
    m_sheets.insert(idx+1, dup);
    m_activeIdx = idx+1;
    emit sheetListChanged();
    emit activeSheetChanged(m_activeIdx);
}

// ── Serialisation ──────────────────────────────────────────────────────────
QString Workbook::serialise() const {
    QJsonArray shArr;
    for(auto* sh : m_sheets) shArr << QJsonDocument::fromJson(sh->toJson()).object();
    QJsonObject root;
    root["version"] = "1.0";
    root["active"]  = m_activeIdx;
    root["sheets"]  = shArr;
    return QJsonDocument(root).toJson();
}

bool Workbook::deserialise(const QString& json) {
    auto doc = QJsonDocument::fromJson(json.toUtf8());
    if(!doc.isObject()) return false;
    QJsonObject root = doc.object();
    qDeleteAll(m_sheets); m_sheets.clear();
    for(const auto& v : root["sheets"].toArray()) {
        auto* sh = Sheet::fromJson(QJsonDocument(v.toObject()).toJson(), this);
        if(!sh) sh = new Sheet("Sheet",this);
        connect(sh,&Sheet::cellChanged,this,[this](int,int){ setModified(true); });
        connect(sh,&Sheet::sheetChanged,this,[this]{ setModified(true); });
        m_sheets << sh;
    }
    if(m_sheets.isEmpty()) addSheet("Sheet1");
    m_activeIdx = qBound(0,root["active"].toInt(0),m_sheets.size()-1);
    return true;
}

bool Workbook::save(const QString& path) {
    QFile f(path); if(!f.open(QIODevice::WriteOnly)) return false;
    f.write(serialise().toUtf8());
    m_filePath = path; setModified(false); return true;
}

bool Workbook::load(const QString& path) {
    QFile f(path); if(!f.open(QIODevice::ReadOnly)) return false;
    if(!deserialise(f.readAll())) return false;
    m_filePath = path; setModified(false);
    emit sheetListChanged(); emit activeSheetChanged(m_activeIdx); return true;
}

bool Workbook::exportCsv(const QString& path) {
    Sheet* sh = activeSheet(); if(!sh) return false;
    QFile f(path); if(!f.open(QIODevice::WriteOnly|QIODevice::Text)) return false;
    QTextStream out(&f);
    int rows=sh->usedRowCount(), cols=sh->usedColCount();
    for(int r=0;r<rows;r++){
        QStringList row;
        for(int c=0;c<cols;c++){
            QString v=sh->getCellValue(r,c).toString();
            if(v.contains(',')) v="\""+v.replace('"','\"')+"\"";
            row<<v;
        }
        out<<row.join(',')<<"\n";
    }
    return true;
}

bool Workbook::importCsv(const QString& path) {
    QFile f(path); if(!f.open(QIODevice::ReadOnly|QIODevice::Text)) return false;
    pushSnapshot();
    QTextStream in(&f); int row=0;
    while(!in.atEnd()){
        QString line=in.readLine();
        QStringList cols; QString cur; bool inQ=false;
        for(QChar ch:line){
            if(ch=='"'){if(inQ&&line.size()>0&&line[0]=='"'){cur+='"';}else inQ=!inQ;}
            else if(ch==','&&!inQ){cols<<cur;cur.clear();}
            else cur+=ch;
        }
        cols<<cur;
        for(int c=0;c<cols.size();c++) if(!cols[c].trimmed().isEmpty()) activeSheet()->setCellValue(row,c,cols[c].trimmed());
        row++;
    }
    setModified(true); return true;
}

// ── Undo ──────────────────────────────────────────────────────────────────
void Workbook::pushSnapshot() {
    m_undoStack << serialise(); if(m_undoStack.size()>100) m_undoStack.removeFirst();
    m_redoStack.clear();
}
void Workbook::undo() {
    if(m_undoStack.isEmpty()) return;
    m_redoStack << serialise();
    deserialise(m_undoStack.takeLast());
    emit sheetListChanged(); emit activeSheetChanged(m_activeIdx);
}
void Workbook::redo() {
    if(m_redoStack.isEmpty()) return;
    m_undoStack << serialise();
    deserialise(m_redoStack.takeLast());
    emit sheetListChanged(); emit activeSheetChanged(m_activeIdx);
}
bool Workbook::canUndo() const { return !m_undoStack.isEmpty(); }
bool Workbook::canRedo() const { return !m_redoStack.isEmpty(); }
