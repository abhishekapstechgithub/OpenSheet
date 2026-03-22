#pragma once
#include "Sheet.h"
#include <QObject>
#include <QVector>
#include <QString>

class UndoStack;

class Workbook : public QObject {
    Q_OBJECT
public:
    explicit Workbook(QObject* parent = nullptr);
    ~Workbook();

    // ── Sheets ────────────────────────────────────────────────────────────────
    int          sheetCount()              const { return m_sheets.size(); }
    Sheet*       sheet(int idx)            const { return m_sheets.value(idx); }
    Sheet*       activeSheet()             const { return m_sheets.value(m_activeIdx); }
    int          activeIndex()             const { return m_activeIdx; }
    void         setActiveSheet(int idx);
    Sheet*       addSheet(const QString& name = QString());
    void         removeSheet(int idx);
    void         renameSheet(int idx, const QString& name);
    void         moveSheet(int from, int to);
    void         duplicateSheet(int idx);

    // ── File ops ─────────────────────────────────────────────────────────────
    bool         save(const QString& path);
    bool         load(const QString& path);
    bool         exportCsv(const QString& path);
    bool         importCsv(const QString& path);
    void         newFile();

    QString      filePath()   const { return m_filePath; }
    bool         isModified() const { return m_modified; }
    void         setModified(bool m){ m_modified=m; emit modifiedChanged(m); }

    // ── Undo ─────────────────────────────────────────────────────────────────
    void pushSnapshot();
    void undo();
    void redo();
    bool canUndo() const;
    bool canRedo() const;

signals:
    void sheetListChanged();
    void activeSheetChanged(int idx);
    void modifiedChanged(bool modified);

private:
    QVector<Sheet*> m_sheets;
    int             m_activeIdx   { 0 };
    QString         m_filePath;
    bool            m_modified    { false };
    QStringList     m_undoStack;
    QStringList     m_redoStack;

    QString serialise() const;
    bool    deserialise(const QString& json);
};
