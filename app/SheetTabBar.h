// ─── SheetTabBar.h ────────────────────────────────────────────────────────────
#pragma once
#include <QWidget>
#include <QTabBar>
class Workbook;

class SheetTabBar : public QWidget {
    Q_OBJECT
public:
    explicit SheetTabBar(Workbook* wb, QWidget* parent=nullptr);
    void refresh();
signals:
    void sheetActivated(int idx);
    void sheetRenamed(int idx, const QString& name);
    void sheetAdded();
    void sheetDeleted(int idx);
    void sheetDuplicated(int idx);
private:
    Workbook* m_wb;
    QTabBar*  m_tabs;
};
