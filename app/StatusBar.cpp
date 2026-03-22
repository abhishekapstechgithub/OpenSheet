#include "StatusBar.h"
#include <QHBoxLayout>
#include <QLabel>
#include <QSlider>
#include <QToolButton>
#include <QFrame>

OSStatusBar::OSStatusBar(QWidget* parent) : QWidget(parent) {
    setFixedHeight(22);
    setStyleSheet("background:#f0f2f5;border-top:1px solid #d8dce3;");
    auto* hl=new QHBoxLayout(this); hl->setContentsMargins(10,0,10,0); hl->setSpacing(7);

    // Ready indicator
    auto* dot=new QLabel; dot->setFixedSize(7,7);
    dot->setStyleSheet("background:#1e7145;border-radius:3px;");
    auto* ready=new QLabel("Ready"); ready->setStyleSheet("font-size:11px;color:#454d5c;");

    auto* sep1=new QFrame; sep1->setFrameShape(QFrame::VLine); sep1->setFixedWidth(1);
    sep1->setStyleSheet("color:#d8dce3;"); sep1->setFixedHeight(12);

    m_mode=new QLabel("Normal"); m_mode->setStyleSheet("font-size:11px;color:#454d5c;");

    auto* sep2=new QFrame; sep2->setFrameShape(QFrame::VLine); sep2->setFixedWidth(1);
    sep2->setStyleSheet("color:#d8dce3;"); sep2->setFixedHeight(12);

    m_stats=new QLabel; m_stats->setStyleSheet("font-size:11px;color:#454d5c;");

    hl->addWidget(dot); hl->addWidget(ready); hl->addWidget(sep1); hl->addWidget(m_mode);
    hl->addWidget(sep2); hl->addWidget(m_stats,1);

    // View mode buttons
    for(auto& [text,tip] : QList<QPair<QString,QString>>{{"⊞","Normal"},{"📄","Page Layout"},{"⊟","Page Break"}}) {
        auto* btn=new QToolButton; btn->setText(text); btn->setToolTip(tip);
        btn->setFixedSize(17,17); btn->setAutoRaise(true); btn->setCheckable(true);
        btn->setStyleSheet("QToolButton{border:1px solid #d8dce3;border-radius:2px;background:transparent;font-size:10px;}"
                           "QToolButton:checked{background:#c8e8d8;border-color:#1e7145;}"
                           "QToolButton:hover{background:#e8f5ee;}");
        if(text=="⊞") btn->setChecked(true);
        hl->addWidget(btn);
    }

    auto* sep3=new QFrame; sep3->setFrameShape(QFrame::VLine); sep3->setFixedWidth(1);
    sep3->setStyleSheet("color:#d8dce3;"); sep3->setFixedHeight(12);
    hl->addWidget(sep3);

    // Zoom controls
    auto* zMinus=new QToolButton; zMinus->setText("−"); zMinus->setFixedSize(16,16);
    zMinus->setStyleSheet("QToolButton{border:1px solid #d8dce3;border-radius:2px;background:white;"
                          "font-size:13px;font-weight:700;color:#454d5c;}"
                          "QToolButton:hover{background:#e8f5ee;border-color:#1e7145;color:#1e7145;}");

    m_zoomSlider=new QSlider(Qt::Horizontal); m_zoomSlider->setRange(50,300);
    m_zoomSlider->setValue(100); m_zoomSlider->setFixedWidth(76); m_zoomSlider->setFixedHeight(14);
    m_zoomSlider->setStyleSheet(
        "QSlider::groove:horizontal{height:3px;background:#d8dce3;border-radius:1px;}"
        "QSlider::handle:horizontal{width:12px;height:12px;border-radius:6px;background:white;"
        "border:2px solid #1e7145;margin:-4px 0;}"
        "QSlider::sub-page:horizontal{background:#1e7145;border-radius:1px;}");

    auto* zPlus=new QToolButton; zPlus->setText("+"); zPlus->setFixedSize(16,16);
    zPlus->setStyleSheet(zMinus->styleSheet());

    m_zoomLabel=new QLabel("100%"); m_zoomLabel->setFixedWidth(34);
    m_zoomLabel->setAlignment(Qt::AlignRight|Qt::AlignVCenter);
    m_zoomLabel->setStyleSheet("font-size:11px;color:#454d5c;font-family:'Courier New';");

    hl->addWidget(zMinus); hl->addWidget(m_zoomSlider); hl->addWidget(zPlus); hl->addWidget(m_zoomLabel);

    connect(m_zoomSlider,&QSlider::valueChanged,this,[this](int v){ setZoom(v); emit zoomChanged(v); });
    connect(zMinus,&QToolButton::clicked,this,[this]{ setZoom(qMax(50,m_zoomSlider->value()-10)); emit zoomChanged(m_zoomSlider->value()); });
    connect(zPlus, &QToolButton::clicked,this,[this]{ setZoom(qMin(300,m_zoomSlider->value()+10));emit zoomChanged(m_zoomSlider->value()); });
}

void OSStatusBar::setMode(const QString& m)      { m_mode->setText(m); }
void OSStatusBar::setStats(const QString& stats)  { m_stats->setText(stats); }
void OSStatusBar::setZoom(int pct) {
    QSignalBlocker b(m_zoomSlider); m_zoomSlider->setValue(pct);
    m_zoomLabel->setText(QString::number(pct)+"%");
}
