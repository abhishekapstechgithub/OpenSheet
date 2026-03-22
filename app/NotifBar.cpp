#include "NotifBar.h"
#include <QHBoxLayout>
#include <QLabel>
#include <QPushButton>
#include <QToolButton>

NotifBar::NotifBar(QWidget* parent) : QWidget(parent) {
    setFixedHeight(28);
    setStyleSheet("background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #e5f6ec,stop:1 #f0fbf4);"
                  "border-bottom:1px solid #b5d9c5;");
    auto* hl=new QHBoxLayout(this); hl->setContentsMargins(10,2,8,2); hl->setSpacing(7);

    auto* star=new QLabel("★"); star->setStyleSheet("color:#e09000;font-size:14px;");
    auto* msg=new QLabel("<b>OpenSheet Pro</b> — Unlock pivot tables · AI formulas · Macros · Real-time collaboration · 500GB files · Cloud sync");
    msg->setStyleSheet("font-size:11.5px;color:#155c34;");

    auto* upgradeBtn=new QPushButton("Upgrade Now");
    upgradeBtn->setFixedHeight(22);
    upgradeBtn->setStyleSheet(
        "QPushButton{background:#1e7145;color:white;border:none;border-radius:11px;"
        "padding:0 14px;font-size:11px;font-weight:600;}"
        "QPushButton:hover{background:#155c34;}");

    auto* closeBtn=new QToolButton; closeBtn->setText("✕"); closeBtn->setFixedSize(18,18);
    closeBtn->setStyleSheet("QToolButton{border:none;background:transparent;color:#7a8494;font-size:13px;}"
                             "QToolButton:hover{background:rgba(220,0,0,.1);color:#d32f2f;border-radius:9px;}");

    hl->addWidget(star); hl->addWidget(msg,1); hl->addWidget(upgradeBtn); hl->addWidget(closeBtn);

    connect(upgradeBtn,&QPushButton::clicked,this,&NotifBar::upgradeClicked);
    connect(closeBtn,  &QToolButton::clicked,this,[this]{ hide(); });
}
