#include "SheetTabBar.h"
#include "Workbook.h"
#include <QHBoxLayout>
#include <QPushButton>
#include <QToolButton>
#include <QInputDialog>
#include <QMenu>
#include <QContextMenuEvent>

SheetTabBar::SheetTabBar(Workbook* wb, QWidget* parent) : QWidget(parent), m_wb(wb) {
    setFixedHeight(30);
    setStyleSheet("background:#f0f2f5;border-top:1px solid #d8dce3;");
    auto* hl = new QHBoxLayout(this); hl->setContentsMargins(6,2,6,2); hl->setSpacing(3);

    // Navigation buttons
    for(auto& [text,tip,delta] : QList<std::tuple<QString,QString,int>>{
        {"⏮","First",-99},{"◀","Previous",-1},{"▶","Next",1},{"⏭","Last",99}})
    {
        auto* btn=new QToolButton; btn->setText(text); btn->setToolTip(tip);
        btn->setFixedSize(20,20); btn->setAutoRaise(true);
        btn->setStyleSheet("QToolButton{border:1px solid #d8dce3;border-radius:2px;background:white;font-size:10px;}"
                           "QToolButton:hover{background:#e8f5ee;border-color:#1e7145;}");
        connect(btn,&QToolButton::clicked,this,[this,delta]{
            int idx=qBound(0,(delta<-10?0:delta>10?m_wb->sheetCount()-1:m_wb->activeIndex()+delta),m_wb->sheetCount()-1);
            emit sheetActivated(idx);
        });
        hl->addWidget(btn);
    }

    m_tabs = new QTabBar; m_tabs->setExpanding(false); m_tabs->setMovable(true);
    m_tabs->setStyleSheet(
        "QTabBar::tab{padding:0 14px;height:24px;font-size:12px;background:#eef0f3;"
        "border:1px solid #d8dce3;border-bottom:none;border-radius:3px 3px 0 0;color:#454d5c;margin-right:1px;}"
        "QTabBar::tab:selected{background:white;color:#1e7145;font-weight:600;border-color:#1e7145;border-bottom:2px solid white;}"
        "QTabBar::tab:hover:!selected{background:white;color:#1a1d23;}");
    connect(m_tabs, &QTabBar::currentChanged, this, &SheetTabBar::sheetActivated);
    connect(m_tabs, &QTabBar::tabMoved, this, [this](int from, int to){ m_wb->moveSheet(from,to); });
    m_tabs->installEventFilter(this);
    hl->addWidget(m_tabs, 1);

    auto* addBtn = new QToolButton; addBtn->setText("+"); addBtn->setToolTip("Add Sheet");
    addBtn->setFixedSize(22,22); addBtn->setAutoRaise(true);
    addBtn->setStyleSheet("QToolButton{border:1px solid #d8dce3;border-radius:11px;background:white;font-size:15px;line-height:1;}"
                          "QToolButton:hover{background:#1e7145;color:white;border-color:#1e7145;}");
    connect(addBtn,&QToolButton::clicked,this,&SheetTabBar::sheetAdded);
    hl->addWidget(addBtn);

    refresh();
}

void SheetTabBar::refresh() {
    QSignalBlocker b(m_tabs);
    while(m_tabs->count()>0) m_tabs->removeTab(0);
    for(int i=0;i<m_wb->sheetCount();i++) m_tabs->addTab(m_wb->sheet(i)->name());
    m_tabs->setCurrentIndex(m_wb->activeIndex());
}

bool eventFilter(QObject* obj, QEvent* event) { return false; }
