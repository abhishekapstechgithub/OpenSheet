// ═══════════════════════════════════════════════════════════════════════════════
//  RibbonWidget.cpp — WPS ET 2018white_dark style ribbon
//  All 7 groups: Clipboard | Font | Alignment | Number | Styles | Cells | Editing
// ═══════════════════════════════════════════════════════════════════════════════
#include "RibbonWidget.h"
#include "ColorButton.h"
#include "Theme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QFrame>
#include <QLabel>
#include <QTabBar>
#include <QStackedWidget>
#include <QFontComboBox>
#include <QMenu>
#include <QAction>
#include <QPainter>
#include <QPainterPath>
#include <QPixmap>
#include <QPolygon>
#include <QApplication>
#include <functional>

// ════════════════════════════════════════════════════════════════════════════
//  ICON FACTORY
// ════════════════════════════════════════════════════════════════════════════
QIcon RibbonWidget::makeIcon(std::function<void(QPainter&,int)> fn, int sz) {
    QPixmap pm(sz*2,sz*2); pm.fill(Qt::transparent);
    QPainter p(&pm); p.setRenderHint(QPainter::Antialiasing,true);
    p.scale(2,2); fn(p,sz); p.end();
    pm.setDevicePixelRatio(2.0);
    return QIcon(pm);
}

static QPen dp(qreal w=1.6){ return QPen(QColor("#444"),w,Qt::SolidLine,Qt::RoundCap,Qt::RoundJoin); }
static QPen gp(qreal w=1.8){ return QPen(QColor("#1e7145"),w,Qt::SolidLine,Qt::RoundCap,Qt::RoundJoin); }
static QPen rp(qreal w=1.5){ return QPen(QColor("#d32f2f"),w,Qt::SolidLine,Qt::RoundCap,Qt::RoundJoin); }

namespace Icons {
static QIcon paste() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.2)); p.setBrush(QColor("#dde1e7")); p.drawRoundedRect(3,7,14,12,1,1);
    p.setBrush(Qt::white); p.drawRoundedRect(6,2,8,6,1,1); p.setPen(gp(1.4)); p.drawRoundedRect(6,2,8,6,1,1);
    p.setPen(dp(1.0)); p.drawLine(5,12,14,12); p.drawLine(5,15,14,15);
});}
static QIcon cut() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.6)); p.drawEllipse(1,13,5,5); p.drawEllipse(13,13,5,5);
    p.drawLine(4,13,10,7); p.drawLine(16,13,10,7); p.drawLine(5,2,10,7); p.drawLine(15,2,10,7);
});}
static QIcon copy() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp()); p.setBrush(QColor("#e8e8e8")); p.drawRoundedRect(5,5,13,13,1,1);
    p.setBrush(Qt::white); p.setPen(dp(1.3)); p.drawRoundedRect(2,2,13,13,1,1);
});}
static QIcon fmtPaint() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setBrush(QColor("#1e7145")); p.setPen(gp()); p.drawRoundedRect(1,1,9,12,1,1);
    p.setPen(dp(1.5)); p.setBrush(Qt::NoBrush); p.drawLine(10,11,17,17); p.drawLine(12,9,17,6);
});}
static QIcon bold() { return RibbonWidget::makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",(int)(s*0.65),QFont::Black); p.setFont(f); p.setPen(QColor("#222"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"B");
});}
static QIcon italic() { return RibbonWidget::makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",(int)(s*0.65),QFont::Normal,true); p.setFont(f); p.setPen(QColor("#222"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"I");
});}
static QIcon underline() { return RibbonWidget::makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",(int)(s*0.55),QFont::Bold); p.setFont(f); p.setPen(QColor("#222"));
    p.drawText(QRect(0,-1,s,s),Qt::AlignCenter,"U");
    p.setPen(QPen(QColor("#1e7145"),2)); p.drawLine(2,s-2,s-2,s-2);
});}
static QIcon strike() { return RibbonWidget::makeIcon([](QPainter&p,int s){
    QFont f("Segoe UI",(int)(s*0.5),QFont::Normal); p.setFont(f); p.setPen(QColor("#222"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"S");
    p.setPen(QPen(QColor("#d32f2f"),1.5)); p.drawLine(2,s/2,s-2,s/2);
});}
static QIcon borders() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.3)); p.drawRect(2,2,16,16);
    p.setPen(QPen(QColor("#aaa"),0.7)); p.drawLine(10,2,10,18); p.drawLine(2,10,18,10);
});}
static QIcon alignL() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.5)); p.drawLine(2,5,18,5); p.drawLine(2,9,13,9); p.drawLine(2,13,18,13); p.drawLine(2,17,11,17);
});}
static QIcon alignC() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.5)); p.drawLine(2,5,18,5); p.drawLine(5,9,15,9); p.drawLine(2,13,18,13); p.drawLine(5,17,15,17);
});}
static QIcon alignR() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.5)); p.drawLine(2,5,18,5); p.drawLine(7,9,18,9); p.drawLine(2,13,18,13); p.drawLine(9,17,18,17);
});}
static QIcon vTop() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(gp(2)); p.drawLine(2,2,18,2);
    p.setPen(dp(1.5)); p.drawLine(6,5,10,15); p.drawLine(14,5,10,15); p.drawLine(6,5,14,5);
});}
static QIcon vMid() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(gp(2)); p.drawLine(2,10,18,10); p.setPen(dp(1.5)); p.drawRect(5,3,10,14);
});}
static QIcon vBot() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(gp(2)); p.drawLine(2,18,18,18);
    p.setPen(dp(1.5)); p.drawLine(6,15,10,5); p.drawLine(14,15,10,5); p.drawLine(6,15,14,15);
});}
static QIcon wrap() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.4)); p.drawLine(2,5,16,5); p.drawLine(2,9,18,9);
    QPainterPath path; path.moveTo(18,9); path.cubicTo(20,9,20,14,18,14); path.lineTo(14,14);
    p.drawPath(path);
    p.setBrush(QColor("#444")); p.setPen(Qt::NoPen);
    QPolygon a; a<<QPoint(14,11)<<QPoint(14,17)<<QPoint(10,14); p.drawPolygon(a);
    p.setPen(dp(1.4)); p.drawLine(2,13,10,13);
});}
static QIcon merge() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.1)); p.drawRect(2,2,7,7); p.drawRect(11,2,7,7); p.drawRect(2,11,7,7); p.drawRect(11,11,7,7);
    p.setPen(gp(2)); p.drawLine(10,2,10,18); p.drawLine(2,10,18,10);
});}
static QIcon condFmt() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.0)); p.setBrush(QColor("#ffe8e8")); p.drawRoundedRect(2,2,16,7,1,1);
    p.setBrush(QColor("#e8ffe8")); p.drawRoundedRect(2,11,16,7,1,1);
    p.setPen(rp(1.2)); p.drawLine(4,5,7,5); p.drawLine(9,5,14,5);
    p.setPen(gp(1.2)); p.drawLine(4,14,14,14);
});}
static QIcon cellStyle() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.0)); p.setBrush(QColor("#c8e8d8")); p.drawRect(2,2,7,7);
    p.setBrush(QColor("#b8d4f0")); p.drawRect(11,2,7,7);
    p.setBrush(QColor("#ffd8a8")); p.drawRect(2,11,7,7);
    p.setBrush(QColor("#f0c8c8")); p.drawRect(11,11,7,7);
});}
static QIcon tableStyle() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(0.5)); p.setBrush(QColor("#1e7145")); p.drawRect(2,2,16,5);
    p.setBrush(QColor("#e8f5ee")); p.drawRect(2,7,8,4); p.setBrush(Qt::white); p.drawRect(10,7,8,4);
    p.setBrush(QColor("#e8f5ee")); p.drawRect(2,11,8,4); p.setBrush(Qt::white); p.drawRect(10,11,8,4);
    p.setBrush(QColor("#e8f5ee")); p.drawRect(2,15,8,3); p.setBrush(Qt::white); p.drawRect(10,15,8,3);
});}
static QIcon insRow() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.3)); p.drawLine(2,6,18,6); p.drawLine(2,14,18,14);
    p.setPen(gp(2.2)); p.drawLine(10,7.5,10,12.5); p.drawLine(7,10,13,10);
});}
static QIcon delRow() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.3)); p.drawLine(2,6,18,6); p.drawLine(2,14,18,14);
    p.setPen(rp(2.2)); p.drawLine(7,10,13,10);
});}
static QIcon insCol() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.3)); p.drawLine(6,2,6,18); p.drawLine(14,2,14,18);
    p.setPen(gp(2.2)); p.drawLine(7.5,10,12.5,10); p.drawLine(10,7,10,13);
});}
static QIcon delCol() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.3)); p.drawLine(6,2,6,18); p.drawLine(14,2,14,18);
    p.setPen(rp(2.2)); p.drawLine(7.5,10,12.5,10);
});}
static QIcon fmtCell() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp()); p.setBrush(Qt::white); p.drawRoundedRect(2,2,16,16,2,2);
    p.setPen(gp(1.5)); p.drawLine(5,7,15,7); p.drawLine(5,11,15,11); p.drawLine(5,15,11,15);
});}
static QIcon rowsCols() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.3)); p.drawLine(2,5,18,5); p.drawLine(2,10,18,10); p.drawLine(2,15,18,15);
    p.setPen(gp(1.8)); p.drawLine(10,2,10,18);
    p.setBrush(QColor("#1e7145")); p.setPen(Qt::NoPen);
    QPolygon u; u<<QPoint(7,3)<<QPoint(10,0)<<QPoint(13,3); p.drawPolygon(u);
    QPolygon d2; d2<<QPoint(7,17)<<QPoint(10,20)<<QPoint(13,17); p.drawPolygon(d2);
});}
static QIcon sheets() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp()); p.setBrush(QColor("#e8e8e8")); p.drawRoundedRect(2,6,16,13,1,1);
    p.setBrush(QColor("#e8f5ee")); p.drawRoundedRect(4,4,7,5,1,1);
    p.setPen(gp(1.5)); p.drawLine(14,3,14,8); p.drawLine(11,5.5,17,5.5);
});}
static QIcon autoSum() { return RibbonWidget::makeIcon([](QPainter&p,int s){
    QFont f("Georgia",(int)(s*0.75),QFont::Bold); p.setFont(f); p.setPen(QColor("#1e7145"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"\u03A3");
});}
static QIcon fill() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp()); p.setBrush(QColor("#e8f5ee")); p.drawRoundedRect(2,2,16,16,2,2);
    p.setPen(gp(2.2)); p.drawLine(10,5,10,15); p.drawLine(5,10,15,10);
});}
static QIcon sortAsc() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.5)); p.drawLine(3,16,8,4); p.drawLine(8,4,13,16); p.drawLine(5,12,11,12);
    p.setPen(dp(1.1)); p.drawLine(15,5,19,5); p.drawLine(15,9,19,9); p.drawLine(15,13,19,13);
});}
static QIcon sortDesc() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.5)); p.drawLine(3,4,8,16); p.drawLine(8,16,13,4); p.drawLine(5,8,11,8);
    p.setPen(dp(1.1)); p.drawLine(15,5,19,5); p.drawLine(15,9,19,9); p.drawLine(15,13,19,13);
});}
static QIcon filter() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(gp(2)); p.drawLine(2,4,18,4); p.drawLine(5,9,15,9); p.drawLine(8,14,12,14);
    p.setBrush(QColor("#1e7145")); p.setPen(Qt::NoPen);
    QPolygon a; a<<QPoint(5,18)<<QPoint(8,15)<<QPoint(8,9)<<QPoint(5,9); p.drawPolygon(a);
});}
static QIcon findIcon() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp(1.8)); p.drawEllipse(2,2,12,12); p.drawLine(11,11,18,18);
});}
static QIcon freeze() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(gp(2)); p.drawLine(9,2,9,18); p.drawLine(2,9,18,9);
    p.setPen(dp(1.0)); p.drawRect(2,2,7,7); p.drawRect(9,2,9,7); p.drawRect(2,9,7,9); p.drawRect(9,9,9,9);
});}
static QIcon fxIcon() { return RibbonWidget::makeIcon([](QPainter&p,int s){
    QFont f("Georgia",(int)(s*0.65),QFont::Bold,true); p.setFont(f); p.setPen(QColor("#1e7145"));
    p.drawText(QRect(0,0,s,s),Qt::AlignCenter,"f");
    QFont f2("Segoe UI",(int)(s*0.4)); p.setFont(f2); p.setPen(QColor("#444"));
    p.drawText(QRect(s/2,s/3,s/2,s*2/3),Qt::AlignVCenter,"x");
});}
static QIcon fmtConv() { return RibbonWidget::makeIcon([](QPainter&p,int){
    p.setPen(dp()); p.setBrush(QColor("#e8f5ee")); p.drawRoundedRect(1,2,11,16,2,2);
    p.setPen(gp(2)); p.drawLine(13,10,19,10);
    p.setBrush(QColor("#1e7145")); p.setPen(Qt::NoPen);
    QPolygon a; a<<QPoint(16,7)<<QPoint(19,10)<<QPoint(16,13); p.drawPolygon(a);
});}
} // namespace Icons

// ════════════════════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════════════════════
QFrame* RibbonWidget::vSep() {
    auto* f=new QFrame; f->setFrameShape(QFrame::VLine); f->setFixedWidth(1);
    f->setFixedHeight(72); f->setStyleSheet("QFrame{color:#e0e3e8;}"); return f;
}

QWidget* RibbonWidget::groupLabel(const QString& text) {
    auto* l=new QLabel(text); l->setAlignment(Qt::AlignCenter);
    l->setStyleSheet("QLabel{color:#9aa0ac;font-size:10px;font-weight:500;}");
    l->setFixedHeight(14); return l;
}

QToolButton* RibbonWidget::mkXLBtn(const QIcon& icon,const QString& label,const QString& tip) {
    auto* btn=new QToolButton; btn->setIcon(icon); btn->setIconSize({26,26});
    btn->setText(label); btn->setToolTip(tip);
    btn->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);
    btn->setAutoRaise(true); btn->setFixedSize(52,72);
    btn->setStyleSheet(
        "QToolButton{border:1px solid transparent;border-radius:3px;background:transparent;"
        "font-size:10px;font-weight:500;color:#454d5c;padding:3px 1px 2px;}"
        "QToolButton:hover{background:#e8f5ee;border-color:#c8e8d8;}"
        "QToolButton:pressed,QToolButton:checked{background:#c8e8d8;border-color:#1e7145;}");
    return btn;
}

QToolButton* RibbonWidget::mkSmBtn(const QIcon& icon,const QString& tip,bool checkable) {
    auto* btn=new QToolButton; btn->setIcon(icon); btn->setIconSize({16,16});
    btn->setToolTip(tip); btn->setAutoRaise(true); btn->setCheckable(checkable);
    btn->setFixedSize(27,27);
    btn->setStyleSheet(
        "QToolButton{border:1px solid transparent;border-radius:3px;background:transparent;}"
        "QToolButton:hover{background:#e8f5ee;border-color:#c8e8d8;}"
        "QToolButton:pressed,QToolButton:checked{background:#c8e8d8;border-color:#1e7145;}");
    return btn;
}

QToolButton* RibbonWidget::mkMdBtn(const QIcon& icon,const QString& label,const QString& tip) {
    auto* btn=new QToolButton; btn->setIcon(icon); btn->setIconSize({18,18});
    btn->setText(label); btn->setToolTip(tip);
    btn->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
    btn->setAutoRaise(true); btn->setFixedHeight(36); btn->setMinimumWidth(80);
    btn->setPopupMode(QToolButton::InstantPopup);
    btn->setStyleSheet(
        "QToolButton{border:1px solid transparent;border-radius:3px;background:transparent;"
        "font-size:10px;font-weight:500;color:#454d5c;padding:2px 6px;}"
        "QToolButton:hover{background:#e8f5ee;border-color:#c8e8d8;}"
        "QToolButton::menu-indicator{width:8px;image:none;}"
        "QToolButton:pressed{background:#c8e8d8;border-color:#1e7145;}");
    return btn;
}

static QMenu* styledMenu(QWidget* parent=nullptr) {
    auto* m=new QMenu(parent);
    m->setStyleSheet(
        "QMenu{background:white;border:1px solid #d8dce3;border-radius:4px;padding:3px 0;"
        "font-family:'Segoe UI',Arial;font-size:12px;}"
        "QMenu::item{padding:6px 20px 6px 14px;color:#1a1d23;}"
        "QMenu::item:selected{background:#e8f5ee;color:#1e7145;}"
        "QMenu::separator{height:1px;background:#d8dce3;margin:3px 0;}");
    return m;
}

// ════════════════════════════════════════════════════════════════════════════
//  CONSTRUCTION
// ════════════════════════════════════════════════════════════════════════════
RibbonWidget::RibbonWidget(QWidget* parent) : QWidget(parent) {
    setFixedHeight(134);
    setAttribute(Qt::WA_StyledBackground,true);
    setStyleSheet(
        "RibbonWidget{background:#ffffff;border-bottom:1px solid #d8dce3;}"
        "QTabBar::tab{padding:0 14px;height:30px;font-size:12px;font-weight:400;"
        "color:#454d5c;background:transparent;border:none;border-bottom:2px solid transparent;}"
        "QTabBar::tab:selected{color:#1e7145;border-bottom:2px solid #1e7145;font-weight:600;}"
        "QTabBar::tab:hover:!selected{color:#1e7145;background:#f0f8f4;}"
        "QComboBox{border:1px solid #d8dce3;border-radius:3px;background:white;padding:2px 6px;"
        "font-size:12px;selection-background-color:#c8e8d8;}"
        "QComboBox:hover,QComboBox:focus{border-color:#1e7145;}"
        "QComboBox::drop-down{border:none;width:16px;}"
        "QSpinBox{border:1px solid #d8dce3;border-radius:3px;background:white;font-size:12px;"
        "min-width:44px;}"
        "QSpinBox:hover,QSpinBox:focus{border-color:#1e7145;}");
    buildLayout();
}

void RibbonWidget::buildLayout() {
    auto* vl=new QVBoxLayout(this);
    vl->setContentsMargins(0,0,0,0); vl->setSpacing(0);

    // Tab bar row (bg: #f2f3f5)
    auto* tabRow=new QWidget; tabRow->setFixedHeight(30);
    tabRow->setStyleSheet("background:#f2f3f5;border-bottom:1px solid #d8dce3;");
    auto* tabHl=new QHBoxLayout(tabRow); tabHl->setContentsMargins(4,0,0,0); tabHl->setSpacing(0);

    m_tabs=new QTabBar; m_tabs->setExpanding(false); m_tabs->setDrawBase(false);
    m_tabs->addTab("Home"); m_tabs->addTab("Insert"); m_tabs->addTab("Page Layout");
    m_tabs->addTab("Formulas"); m_tabs->addTab("Data"); m_tabs->addTab("Review");
    m_tabs->addTab("View"); m_tabs->addTab("Tools");
    tabHl->addWidget(m_tabs); tabHl->addStretch();
    connect(m_tabs,&QTabBar::currentChanged,this,[this](int i){
        if(i!=0) { QTimer::singleShot(0,this,[this]{ m_tabs->setCurrentIndex(0); }); }
    });

    vl->addWidget(tabRow);

    // Home ribbon body
    m_homePanel=buildHomeTab();
    vl->addWidget(m_homePanel);
}

QWidget* RibbonWidget::buildHomeTab() {
    auto* panel=new QWidget; panel->setFixedHeight(104);
    panel->setStyleSheet("background:#ffffff;");
    auto* hl=new QHBoxLayout(panel);
    hl->setContentsMargins(6,2,6,0); hl->setSpacing(0);

    auto addGroup=[&](QWidget* g){ hl->addWidget(g); hl->addWidget(vSep()); hl->addSpacing(2); };

    addGroup(buildClipboardGroup());
    addGroup(buildFontGroup());
    addGroup(buildAlignGroup());
    addGroup(buildNumberGroup());
    addGroup(buildStylesGroup());
    addGroup(buildCellsGroup());
    hl->addWidget(buildEditingGroup());
    hl->addStretch();
    return panel;
}

// ── CLIPBOARD ───────────────────────────────────────────────────────────────
QWidget* RibbonWidget::buildClipboardGroup() {
    auto* g=new QWidget; auto* col=new QVBoxLayout(g); col->setContentsMargins(2,0,2,0); col->setSpacing(0);
    auto* row=new QHBoxLayout; row->setSpacing(2);

    auto* btnPaste=mkXLBtn(Icons::paste(),"Paste","Paste (Ctrl+V)");
    auto* stack=new QVBoxLayout; stack->setSpacing(2);
    auto* btnCut =mkSmBtn(Icons::cut(),  "Cut (Ctrl+X)");
    auto* btnCopy=mkSmBtn(Icons::copy(), "Copy (Ctrl+C)");
    auto* btnFmt =mkSmBtn(Icons::fmtPaint(),"Format Painter");
    stack->addWidget(btnCut); stack->addWidget(btnCopy); stack->addWidget(btnFmt);

    row->addWidget(btnPaste); row->addSpacing(2); row->addLayout(stack);
    col->addLayout(row); col->addWidget(groupLabel("Clipboard"));

    connect(btnPaste,&QToolButton::clicked,this,&RibbonWidget::pasteRequested);
    connect(btnCut,  &QToolButton::clicked,this,&RibbonWidget::cutRequested);
    connect(btnCopy, &QToolButton::clicked,this,&RibbonWidget::copyRequested);
    connect(btnFmt,  &QToolButton::clicked,this,&RibbonWidget::formatPainterRequested);
    return g;
}

// ── FONT ─────────────────────────────────────────────────────────────────────
QWidget* RibbonWidget::buildFontGroup() {
    auto* g=new QWidget; auto* col=new QVBoxLayout(g); col->setContentsMargins(2,0,2,0); col->setSpacing(3);

    // Row 1: font family | size | grow | shrink
    auto* r1=new QHBoxLayout; r1->setSpacing(2);
    m_fontFamily=new QComboBox;
    const QStringList fonts={"Calibri","Arial","Helvetica","Times New Roman","Georgia","Verdana","Courier New","Tahoma","Cambria","Trebuchet MS","Palatino","Impact"};
    m_fontFamily->addItems(fonts); m_fontFamily->setFixedHeight(24); m_fontFamily->setMaximumWidth(148);
    m_fontSize=new QSpinBox; m_fontSize->setRange(5,96); m_fontSize->setValue(11); m_fontSize->setFixedHeight(24); m_fontSize->setFixedWidth(48);
    auto* btnGrow  =mkSmBtn(makeIcon([](QPainter&p,int s){ QFont f("Segoe UI",(int)(s*.55),QFont::Bold);p.setFont(f);p.setPen(QColor("#333"));p.drawText(QRect(0,0,s-4,s),Qt::AlignCenter,"A");p.setPen(gp(1.5));p.drawText(QRect(s/2-1,s/2-4,s/2+2,s/2+3),Qt::AlignCenter,"+");}), "Increase Font Size");
    auto* btnShrink=mkSmBtn(makeIcon([](QPainter&p,int s){ QFont f("Segoe UI",(int)(s*.48));p.setFont(f);p.setPen(QColor("#333"));p.drawText(QRect(0,0,s-4,s),Qt::AlignCenter,"A");p.setPen(rp(1.5));p.drawText(QRect(s/2-1,s/2-4,s/2+2,s/2+3),Qt::AlignCenter,"-");}), "Decrease Font Size");
    r1->addWidget(m_fontFamily,1); r1->addWidget(m_fontSize); r1->addWidget(btnGrow); r1->addWidget(btnShrink);

    // Row 2: B I U S | text-color | fill-color | borders
    auto* r2=new QHBoxLayout; r2->setSpacing(2);
    m_btnBold     =mkSmBtn(Icons::bold(),     "Bold (Ctrl+B)",true);
    m_btnItalic   =mkSmBtn(Icons::italic(),   "Italic (Ctrl+I)",true);
    m_btnUnderline=mkSmBtn(Icons::underline(),"Underline (Ctrl+U)",true);
    m_btnStrike   =mkSmBtn(Icons::strike(),   "Strikethrough",true);

    m_textColorBtn=new ColorButton(makeIcon([](QPainter&p,int s){QFont f("Segoe UI",(int)(s*.65),QFont::Bold);p.setFont(f);p.setPen(QColor("#1a1d23"));p.drawText(QRect(0,-1,s,s),Qt::AlignCenter,"A");}),"Font Color",g);
    m_fillColorBtn=new ColorButton(makeIcon([](QPainter&p,int s){
        QPolygon pg; pg<<QPoint(4,s-3)<<QPoint(s/2,2)<<QPoint(s-3,s-3);
        p.setBrush(QColor("#1e7145")); p.setPen(Qt::NoPen); p.drawPolygon(pg);
        p.setBrush(QColor("#888")); p.setPen(Qt::NoPen); p.drawRect(1,s-3,s-2,3);
    }),"Fill Color",g);
    m_fillColorBtn->setColor(Qt::yellow);

    auto* btnBord=mkSmBtn(Icons::borders(),"Borders");
    // Borders menu
    auto* bordMenu=styledMenu(btnBord);
    const QList<QPair<QString,BorderStyle>> borderItems={
        {"All Borders",BorderStyle::All},{"Outside Borders",BorderStyle::Outside},
        {"Thick Box Border",BorderStyle::Thick},{},
        {"Top Border",BorderStyle::Top},{"Bottom Border",BorderStyle::Bottom},
        {"Left Border",BorderStyle::Left},{"Right Border",BorderStyle::Right},{},
        {"No Border",BorderStyle::None}
    };
    for(auto& [label,style]:borderItems){
        if(label.isEmpty()){bordMenu->addSeparator();continue;}
        auto* act=bordMenu->addAction(label);
        connect(act,&QAction::triggered,this,[this,style]{ emit borderRequested(style); });
    }
    btnBord->setMenu(bordMenu); btnBord->setPopupMode(QToolButton::InstantPopup);

    r2->addWidget(m_btnBold); r2->addWidget(m_btnItalic);
    r2->addWidget(m_btnUnderline); r2->addWidget(m_btnStrike);
    r2->addSpacing(3);
    r2->addWidget(m_textColorBtn); r2->addWidget(m_fillColorBtn); r2->addWidget(btnBord);
    r2->addStretch();

    col->addLayout(r1); col->addLayout(r2); col->addWidget(groupLabel("Font"));

    connect(m_fontFamily,qOverload<int>(&QComboBox::currentIndexChanged),this,
        [this]{ emit fontFamilyChanged(m_fontFamily->currentText()); });
    connect(m_fontSize,qOverload<int>(&QSpinBox::valueChanged),this,&RibbonWidget::fontSizeChanged);
    connect(m_btnBold,      &QToolButton::toggled,this,&RibbonWidget::boldToggled);
    connect(m_btnItalic,    &QToolButton::toggled,this,&RibbonWidget::italicToggled);
    connect(m_btnUnderline, &QToolButton::toggled,this,&RibbonWidget::underlineToggled);
    connect(m_btnStrike,    &QToolButton::toggled,this,&RibbonWidget::strikeThroughToggled);
    connect(m_textColorBtn, &ColorButton::colorChanged,this,&RibbonWidget::textColorChanged);
    connect(m_fillColorBtn, &ColorButton::colorChanged,this,&RibbonWidget::fillColorChanged);
    connect(btnGrow,  &QToolButton::clicked,this,[this]{ m_fontSize->setValue(qMin(96,m_fontSize->value()+1)); });
    connect(btnShrink,&QToolButton::clicked,this,[this]{ m_fontSize->setValue(qMax(5,m_fontSize->value()-1));  });
    return g;
}

// ── ALIGNMENT ────────────────────────────────────────────────────────────────
QWidget* RibbonWidget::buildAlignGroup() {
    auto* g=new QWidget; auto* col=new QVBoxLayout(g); col->setContentsMargins(2,0,2,0); col->setSpacing(3);

    // Row 1: VTop VMid VBot | Ori | Wrap | Merge
    auto* r1=new QHBoxLayout; r1->setSpacing(2);
    m_btnVTop=mkSmBtn(Icons::vTop(),"Align Top",true);
    m_btnVMid=mkSmBtn(Icons::vMid(),"Align Middle",true);
    m_btnVBot=mkSmBtn(Icons::vBot(),"Align Bottom",true);

    auto* btnOri=mkSmBtn(makeIcon([](QPainter&p,int s){
        p.save(); p.translate(s/2,s/2); p.rotate(-45);
        QFont f("Segoe UI",(int)(s*.35)); p.setFont(f); p.setPen(QColor("#1e7145")); p.drawText(-s/2,s/8,QString("abc"));
        p.restore(); p.setPen(QPen(QColor("#1e7145"),1.4)); p.drawLine(2,s-2,s-2,s-2);
    }),"Text Orientation");
    auto* oriMenu=styledMenu(btnOri);
    for(auto& label:QStringList{"Angle Counterclockwise","Angle Clockwise","Vertical Text","Rotate Text Up","Rotate Text Down"})
        oriMenu->addAction(label);
    btnOri->setMenu(oriMenu); btnOri->setPopupMode(QToolButton::InstantPopup);

    m_btnWrap=mkSmBtn(Icons::wrap(),"Wrap Text",true);

    auto* btnMerge=mkSmBtn(Icons::merge(),"Merge Cells");
    auto* mergeMenu=styledMenu(btnMerge);
    auto* actMC=mergeMenu->addAction("Merge and Center");
    auto* actMA=mergeMenu->addAction("Merge Across");
    mergeMenu->addSeparator();
    auto* actUM=mergeMenu->addAction("Unmerge Cells");
    connect(actMC,&QAction::triggered,this,[this]{ emit mergeCellsRequested(true); });
    connect(actMA,&QAction::triggered,this,[this]{ emit mergeCellsRequested(true); });
    connect(actUM,&QAction::triggered,this,[this]{ emit mergeCellsRequested(false); });
    btnMerge->setMenu(mergeMenu); btnMerge->setPopupMode(QToolButton::MenuButtonPopup);

    r1->addWidget(m_btnVTop); r1->addWidget(m_btnVMid); r1->addWidget(m_btnVBot);
    r1->addSpacing(3); r1->addWidget(btnOri); r1->addWidget(m_btnWrap);
    r1->addSpacing(2); r1->addWidget(btnMerge); r1->addStretch();

    // Row 2: AlL AlC AlR | IndentDec IndentInc
    auto* r2=new QHBoxLayout; r2->setSpacing(2);
    m_btnAlL=mkSmBtn(Icons::alignL(),"Align Left",true);
    m_btnAlC=mkSmBtn(Icons::alignC(),"Center",true);
    m_btnAlR=mkSmBtn(Icons::alignR(),"Align Right",true);
    auto* btnIndDec=mkSmBtn(makeIcon([](QPainter&p,int){ p.setPen(dp(1.4)); p.drawLine(6,4,14,4);p.drawLine(6,8,14,8);p.drawLine(6,12,14,12);p.drawLine(6,16,14,16); p.setPen(gp(2)); QPolygon a;a<<QPoint(1,6)<<QPoint(1,14)<<QPoint(4,10);p.drawPolygon(a); }),"Decrease Indent");
    auto* btnIndInc=mkSmBtn(makeIcon([](QPainter&p,int){ p.setPen(dp(1.4)); p.drawLine(6,4,14,4);p.drawLine(6,8,14,8);p.drawLine(6,12,14,12);p.drawLine(6,16,14,16); p.setPen(gp(2)); QPolygon a;a<<QPoint(5,6)<<QPoint(5,14)<<QPoint(8,10);p.drawPolygon(a); }),"Increase Indent");

    r2->addWidget(m_btnAlL); r2->addWidget(m_btnAlC); r2->addWidget(m_btnAlR);
    r2->addSpacing(3); r2->addWidget(btnIndDec); r2->addWidget(btnIndInc); r2->addStretch();

    col->addLayout(r1); col->addLayout(r2); col->addWidget(groupLabel("Alignment"));

    connect(m_btnAlL, &QToolButton::clicked,this,[this]{ emit hAlignChanged(Qt::AlignLeft);    });
    connect(m_btnAlC, &QToolButton::clicked,this,[this]{ emit hAlignChanged(Qt::AlignHCenter); });
    connect(m_btnAlR, &QToolButton::clicked,this,[this]{ emit hAlignChanged(Qt::AlignRight);   });
    connect(m_btnVTop,&QToolButton::clicked,this,[this]{ emit vAlignChanged(Qt::AlignTop);     });
    connect(m_btnVMid,&QToolButton::clicked,this,[this]{ emit vAlignChanged(Qt::AlignVCenter); });
    connect(m_btnVBot,&QToolButton::clicked,this,[this]{ emit vAlignChanged(Qt::AlignBottom);  });
    connect(m_btnWrap,&QToolButton::toggled,this,&RibbonWidget::wrapTextToggled);
    connect(btnIndInc,&QToolButton::clicked,this,[this]{ emit indentChanged(+1); });
    connect(btnIndDec,&QToolButton::clicked,this,[this]{ emit indentChanged(-1); });
    return g;
}

// ── NUMBER FORMAT ────────────────────────────────────────────────────────────
QWidget* RibbonWidget::buildNumberGroup() {
    auto* g=new QWidget; auto* col=new QVBoxLayout(g); col->setContentsMargins(2,0,2,0); col->setSpacing(3);

    m_numFmtCombo=new QComboBox; m_numFmtCombo->setFixedHeight(24); m_numFmtCombo->setMaximumWidth(150);
    m_numFmtCombo->addItems({"General","Number","Currency ($)","Accounting","Percentage (%)","Scientific","Short Date","Long Date","Time","Fraction","Text"});

    // WPS Format Conversion pill button
    auto* btnFmtConv=new QToolButton;
    btnFmtConv->setText(" Format Conversion ▾");
    btnFmtConv->setIcon(Icons::fmtConv()); btnFmtConv->setIconSize({14,14});
    btnFmtConv->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
    btnFmtConv->setFixedHeight(22); btnFmtConv->setMinimumWidth(148);
    btnFmtConv->setStyleSheet(
        "QToolButton{border:1px solid #d8dce3;border-radius:3px;background:#f0f2f5;"
        "font-size:11px;font-weight:500;color:#1e7145;text-align:left;padding:0 6px;}"
        "QToolButton:hover{border-color:#1e7145;background:#e8f5ee;}"
        "QToolButton::menu-indicator{width:0;image:none;}");
    auto* fcMenu=styledMenu(btnFmtConv);
    for(auto& [label,type] : QList<QPair<QString,QString>>{
        {"Convert to Number","number"},{"Convert to Text","text"},{"Convert to Date","date"},
        {"UPPERCASE","upper"},{"lowercase","lower"},{"Proper Case","proper"},
        {"",""}, {"Remove Leading Zeros","rmzero"},{"Remove Extra Spaces","rmspace"},
        {"Remove Non-Numeric Chars","rmnonnumeric"}})
    {
        if(label.isEmpty()){fcMenu->addSeparator();continue;}
        auto* act=fcMenu->addAction(label);
        connect(act,&QAction::triggered,this,[this,type]{ emit formatConversionRequested(type); });
    }
    btnFmtConv->setMenu(fcMenu); btnFmtConv->setPopupMode(QToolButton::InstantPopup);

    // Number symbol row
    auto* r2=new QHBoxLayout; r2->setSpacing(2);
    auto mkSym=[&](const QString& text, const QString& tip, NumFmt fmt)->QToolButton*{
        auto* b=new QToolButton; b->setText(text); b->setToolTip(tip); b->setAutoRaise(true);
        b->setFixedSize(25,25);
        b->setStyleSheet("QToolButton{border:1px solid #d8dce3;border-radius:3px;background:white;"
                          "font-size:13px;font-weight:700;color:#1e7145;}"
                          "QToolButton:hover{background:#e8f5ee;border-color:#1e7145;}");
        connect(b,&QToolButton::clicked,this,[this,fmt]{ emit numberFormatChanged(fmt); });
        return b;
    };
    r2->addWidget(mkSym("$","Currency",NumFmt::Currency));
    r2->addWidget(mkSym("%","Percent", NumFmt::Percent));
    r2->addWidget(mkSym(",","Thousands",NumFmt::Thousands));
    r2->addSpacing(3);
    auto* btnDecInc=new QToolButton; btnDecInc->setText(".0+"); btnDecInc->setToolTip("Increase Decimal");
    btnDecInc->setFixedSize(28,25); btnDecInc->setAutoRaise(true);
    btnDecInc->setStyleSheet("QToolButton{border:1px solid #d8dce3;border-radius:3px;background:white;font-size:10px;font-weight:600;color:#1e7145;}"
                              "QToolButton:hover{background:#e8f5ee;border-color:#1e7145;}");
    auto* btnDecDec=new QToolButton; btnDecDec->setText(".0-"); btnDecDec->setToolTip("Decrease Decimal");
    btnDecDec->setFixedSize(28,25); btnDecDec->setAutoRaise(true);
    btnDecDec->setStyleSheet(btnDecInc->styleSheet());
    r2->addWidget(btnDecInc); r2->addWidget(btnDecDec); r2->addStretch();

    col->addWidget(m_numFmtCombo); col->addWidget(btnFmtConv); col->addLayout(r2);
    col->addWidget(groupLabel("Number"));

    connect(m_numFmtCombo,qOverload<int>(&QComboBox::currentIndexChanged),this,
        [this](int i){ emit numberFormatChanged((NumFmt)i); });
    connect(btnDecInc,&QToolButton::clicked,this,[this]{ emit decimalsChanged(+1); });
    connect(btnDecDec,&QToolButton::clicked,this,[this]{ emit decimalsChanged(-1); });
    return g;
}

// ── STYLES ───────────────────────────────────────────────────────────────────
QWidget* RibbonWidget::buildStylesGroup() {
    auto* g=new QWidget; auto* col=new QVBoxLayout(g); col->setContentsMargins(2,0,2,0); col->setSpacing(2);

    // Conditional Formatting button
    auto* btnCF=mkMdBtn(Icons::condFmt(),"Conditional\nFormatting","Conditional Formatting");
    btnCF->setFixedWidth(136);
    auto* cfMenu=styledMenu(btnCF);
    cfMenu->addSection("Highlight Cell Rules");
    for(auto&[l,t]:QList<QPair<QString,QString>>{{"Greater Than…","gt"},{"Less Than…","lt"},{"Between…","bt"},{"Equal To…","eq"},{"Text Contains…","txt"},{"Duplicate Values","dup"}})
        connect(cfMenu->addAction(l),&QAction::triggered,this,[this,t]{ emit conditionalFormatRequested(t); });
    cfMenu->addSeparator(); cfMenu->addSection("Top/Bottom Rules");
    for(auto&[l,t]:QList<QPair<QString,QString>>{{"Top 10 Items","top10"},{"Bottom 10 Items","bot10"},{"Above Average","above"},{"Below Average","below"}})
        connect(cfMenu->addAction(l),&QAction::triggered,this,[this,t]{ emit conditionalFormatRequested(t); });
    cfMenu->addSeparator();
    for(auto&[l,t]:QList<QPair<QString,QString>>{{"Color Scale (G-W-R)","scale"},{"Data Bars","bars"},{"Icon Sets","icons"}})
        connect(cfMenu->addAction(l),&QAction::triggered,this,[this,t]{ emit conditionalFormatRequested(t); });
    cfMenu->addSeparator();
    connect(cfMenu->addAction("Clear Rules"),&QAction::triggered,this,[this]{ emit conditionalFormatRequested("clear"); });
    btnCF->setMenu(cfMenu);

    // Cell Styles + Table Styles row
    auto* r2=new QHBoxLayout; r2->setSpacing(3);
    auto* btnCS=mkMdBtn(Icons::cellStyle(),"Cell\nStyles","Cell Styles"); btnCS->setFixedWidth(66);
    auto* csMenu=styledMenu(btnCS);
    for(auto&name:QStringList{"Good","Bad","Neutral","Warning","Heading 1","Heading 2","Heading 3","Title","Total","Input","Output","Note","Error"})
        connect(csMenu->addAction(name),&QAction::triggered,this,[this,name]{ emit cellStyleRequested(name); });
    btnCS->setMenu(csMenu);

    auto* btnTS=mkMdBtn(Icons::tableStyle(),"Table\nStyles","Table Styles"); btnTS->setFixedWidth(66);
    auto* tsMenu=styledMenu(btnTS);
    for(auto&name:QStringList{"Light 1 — Green","Light 2 — Blue","Light 3 — Orange","Light 4 — Red","Medium 1 — Green","Medium 2 — Blue","Dark 1 — Green","Dark 2 — Navy"})
        connect(tsMenu->addAction(name),&QAction::triggered,this,[this,name]{ emit tableStyleRequested(name); });
    btnTS->setMenu(tsMenu);

    r2->addWidget(btnCS); r2->addWidget(btnTS);
    col->addWidget(btnCF); col->addLayout(r2); col->addWidget(groupLabel("Styles"));
    return g;
}

// ── CELLS ────────────────────────────────────────────────────────────────────
QWidget* RibbonWidget::buildCellsGroup() {
    auto* g=new QWidget; auto* col=new QVBoxLayout(g); col->setContentsMargins(2,0,2,0); col->setSpacing(2);
    auto* r1=new QHBoxLayout; r1->setSpacing(2);

    auto* btnIR=mkSmBtn(Icons::insRow(),"Insert Row");
    auto* btnDR=mkSmBtn(Icons::delRow(),"Delete Row");
    auto* btnIC=mkSmBtn(Icons::insCol(),"Insert Column");
    auto* btnDC=mkSmBtn(Icons::delCol(),"Delete Column");
    auto* btnFC=mkSmBtn(Icons::fmtCell(),"Format Cells…");
    r1->addWidget(btnIR); r1->addWidget(btnDR); r1->addWidget(btnIC); r1->addWidget(btnDC); r1->addWidget(btnFC);

    auto* btnRC=mkMdBtn(Icons::rowsCols(),"Rows and\nColumns","Row/Column Operations"); btnRC->setFixedWidth(136);
    auto* rcMenu=styledMenu(btnRC);
    for(auto&[l,s]:QList<QPair<QString,QString>>{{"Insert Row Above","insrow"},{"Insert Row Below","insrowb"},{"Delete Row","delrow"},{"Insert Column Left","inscol"},{"Insert Column Right","inscolr"},{"Delete Column","delcol"}}) {
        if(l.isEmpty()){rcMenu->addSeparator();continue;}
        auto* act=rcMenu->addAction(l);
        if(l=="Insert Row Above")   connect(act,&QAction::triggered,this,&RibbonWidget::insertRowRequested);
        else if(l=="Delete Row")    connect(act,&QAction::triggered,this,&RibbonWidget::deleteRowRequested);
        else if(l=="Insert Column Left") connect(act,&QAction::triggered,this,&RibbonWidget::insertColRequested);
        else if(l=="Delete Column") connect(act,&QAction::triggered,this,&RibbonWidget::deleteColRequested);
    }
    rcMenu->addSeparator();
    connect(rcMenu->addAction("Row Height…"),     &QAction::triggered,this,&RibbonWidget::rowHeightRequested);
    connect(rcMenu->addAction("Column Width…"),   &QAction::triggered,this,&RibbonWidget::colWidthRequested);
    rcMenu->addAction("AutoFit Row Height");
    rcMenu->addAction("AutoFit Column Width");
    btnRC->setMenu(rcMenu);

    auto* btnSH=mkMdBtn(Icons::sheets(),"Sheets","Sheet Operations"); btnSH->setFixedWidth(136);
    auto* shMenu=styledMenu(btnSH);
    shMenu->addAction("Insert Sheet"); shMenu->addAction("Delete Sheet"); shMenu->addAction("Rename Sheet");
    shMenu->addAction("Duplicate Sheet"); shMenu->addSeparator();
    shMenu->addAction("Move Left"); shMenu->addAction("Move Right"); shMenu->addSeparator();
    shMenu->addAction("Protect Sheet (Pro)"); shMenu->addAction("Hide Sheet");
    btnSH->setMenu(shMenu);

    col->addLayout(r1); col->addWidget(btnRC); col->addWidget(btnSH); col->addWidget(groupLabel("Cells"));

    connect(btnIR,&QToolButton::clicked,this,&RibbonWidget::insertRowRequested);
    connect(btnDR,&QToolButton::clicked,this,&RibbonWidget::deleteRowRequested);
    connect(btnIC,&QToolButton::clicked,this,&RibbonWidget::insertColRequested);
    connect(btnDC,&QToolButton::clicked,this,&RibbonWidget::deleteColRequested);
    connect(btnFC,&QToolButton::clicked,this,&RibbonWidget::formatCellsRequested);
    return g;
}

// ── EDITING ──────────────────────────────────────────────────────────────────
QWidget* RibbonWidget::buildEditingGroup() {
    auto* g=new QWidget; auto* col=new QVBoxLayout(g); col->setContentsMargins(2,0,2,0); col->setSpacing(0);
    auto* r=new QHBoxLayout; r->setSpacing(2);

    auto* btnSum=mkXLBtn(Icons::autoSum(),"AutoSum","AutoSum (Alt+=)");
    auto* sumMenu=styledMenu(btnSum);
    for(auto&fn:QStringList{"Sum","Average","Count Numbers","Max","Min","More Functions…"})
        sumMenu->addAction(fn);
    btnSum->setMenu(sumMenu); btnSum->setPopupMode(QToolButton::MenuButtonPopup);

    auto* btnFill=mkXLBtn(Icons::fill(),"Fill","Fill");
    auto* fillMenu=styledMenu(btnFill);
    for(auto&[l,d]:QList<QPair<QString,QString>>{{"Fill Down  Ctrl+D","down"},{"Fill Right  Ctrl+R","right"},{"Fill Up","up"},{"Fill Left","left"}})
        connect(fillMenu->addAction(l),&QAction::triggered,this,[this,d]{ emit fillRequested(d); });
    fillMenu->addSeparator();
    fillMenu->addAction("Flash Fill  Ctrl+E");
    fillMenu->addAction("Fill Series…");
    btnFill->setMenu(fillMenu); btnFill->setPopupMode(QToolButton::MenuButtonPopup);

    // Stack of small buttons
    auto* vs=new QVBoxLayout; vs->setSpacing(2);
    auto* vr1=new QHBoxLayout; vr1->setSpacing(2);
    auto* btnSA =mkSmBtn(Icons::sortAsc(), "Sort A to Z");
    auto* btnSD =mkSmBtn(Icons::sortDesc(),"Sort Z to A");
    m_btnFilter =mkSmBtn(Icons::filter(),  "Filter",true);
    vr1->addWidget(btnSA); vr1->addWidget(btnSD); vr1->addWidget(m_btnFilter);

    auto* vr2=new QHBoxLayout; vr2->setSpacing(2);
    auto* btnFrz=mkSmBtn(Icons::freeze(),"Freeze Panes");
    auto* frzMenu=styledMenu(btnFrz);
    for(auto&[l,t]:QList<QPair<QString,QString>>{{"Freeze Top Row","toprow"},{"Freeze First Column","firstcol"},{"Freeze Panes","panes"},{"",""},{"Unfreeze Panes","unfreeze"}})
    {
        if(l.isEmpty()){frzMenu->addSeparator();continue;}
        connect(frzMenu->addAction(l),&QAction::triggered,this,[this,t]{ emit freezeRequested(t); });
    }
    btnFrz->setMenu(frzMenu); btnFrz->setPopupMode(QToolButton::MenuButtonPopup);

    auto* btnFR =mkSmBtn(Icons::findIcon(),"Find & Replace (Ctrl+H)");
    auto* btnFn =mkSmBtn(Icons::fxIcon(),  "Insert Function");
    vr2->addWidget(btnFrz); vr2->addWidget(btnFR); vr2->addWidget(btnFn);

    vs->addLayout(vr1); vs->addLayout(vr2);
    r->addWidget(btnSum); r->addWidget(btnFill); r->addSpacing(3); r->addLayout(vs);
    col->addLayout(r); col->addWidget(groupLabel("Editing"));

    connect(btnSum, &QToolButton::clicked,this,&RibbonWidget::autoSumRequested);
    connect(btnSA,  &QToolButton::clicked,this,[this]{ emit sortRequested(true);  });
    connect(btnSD,  &QToolButton::clicked,this,[this]{ emit sortRequested(false); });
    connect(m_btnFilter,&QToolButton::clicked,this,&RibbonWidget::filterToggled);
    connect(btnFR,  &QToolButton::clicked,this,&RibbonWidget::findReplaceRequested);
    connect(btnFn,  &QToolButton::clicked,this,&RibbonWidget::functionWizardRequested);
    return g;
}

// ════════════════════════════════════════════════════════════════════════════
//  SYNC FROM FORMAT
// ════════════════════════════════════════════════════════════════════════════
void RibbonWidget::syncToFormat(const CellFormat& fmt) {
    QSignalBlocker b1(m_fontFamily), b2(m_fontSize), b3(m_btnBold), b4(m_btnItalic),
                   b5(m_btnUnderline), b6(m_btnStrike), b7(m_numFmtCombo);

    int fi = m_fontFamily->findText(fmt.fontFamily);
    if(fi>=0) m_fontFamily->setCurrentIndex(fi);
    m_fontSize->setValue(fmt.fontSize);
    m_btnBold->setChecked(fmt.bold);
    m_btnItalic->setChecked(fmt.italic);
    m_btnUnderline->setChecked(fmt.underline);
    m_btnStrike->setChecked(fmt.strikeThrough);
    m_btnWrap->setChecked(fmt.wrapText);
    m_textColorBtn->setColor(fmt.textColor.isValid() ? fmt.textColor : Qt::black);
    m_fillColorBtn->setColor(fmt.fillColor.isValid() ? fmt.fillColor : Qt::transparent);
    m_numFmtCombo->setCurrentIndex((int)fmt.numFmt);

    for(auto* b : {m_btnAlL,m_btnAlC,m_btnAlR}) b->setChecked(false);
    if     (fmt.hAlign==Qt::AlignLeft)    m_btnAlL->setChecked(true);
    else if(fmt.hAlign==Qt::AlignHCenter) m_btnAlC->setChecked(true);
    else if(fmt.hAlign==Qt::AlignRight)   m_btnAlR->setChecked(true);

    for(auto* b : {m_btnVTop,m_btnVMid,m_btnVBot}) b->setChecked(false);
    if     (fmt.vAlign==Qt::AlignTop)     m_btnVTop->setChecked(true);
    else if(fmt.vAlign==Qt::AlignVCenter) m_btnVMid->setChecked(true);
    else if(fmt.vAlign==Qt::AlignBottom)  m_btnVBot->setChecked(true);
}
