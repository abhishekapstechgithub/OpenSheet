#include "FormatCellsDialog.h"
#include "Theme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QGroupBox>
#include <QLabel>
#include <QComboBox>
#include <QSpinBox>
#include <QCheckBox>
#include <QDialogButtonBox>
#include <QPushButton>
#include <QColorDialog>

FormatCellsDialog::FormatCellsDialog(const CellFormat& fmt, QWidget* parent)
    : QDialog(parent), m_fmt(fmt), m_textColor(fmt.textColor), m_fillColor(fmt.fillColor)
{
    setWindowTitle("Format Cells");
    setMinimumWidth(480);
    setStyleSheet(Theme::appStyle());

    auto* main=new QVBoxLayout(this);

    // ── Font row ──────────────────────────────────────────────────────────
    auto* fontGroup=new QGroupBox("Font");
    auto* fg=new QGridLayout(fontGroup);
    fg->addWidget(new QLabel("Font Family:"),0,0);
    m_fontFam=new QComboBox;
    m_fontFam->addItems({"Calibri","Arial","Helvetica","Times New Roman","Georgia","Verdana","Courier New","Tahoma","Cambria"});
    m_fontFam->setCurrentText(fmt.fontFamily); fg->addWidget(m_fontFam,0,1);

    fg->addWidget(new QLabel("Font Size:"),0,2);
    m_fontSize=new QSpinBox; m_fontSize->setRange(5,96); m_fontSize->setValue(fmt.fontSize); fg->addWidget(m_fontSize,0,3);

    m_bold      =new QCheckBox("Bold");       m_bold->setChecked(fmt.bold);
    m_italic    =new QCheckBox("Italic");     m_italic->setChecked(fmt.italic);
    m_underline =new QCheckBox("Underline");  m_underline->setChecked(fmt.underline);
    m_strike    =new QCheckBox("Strikethrough"); m_strike->setChecked(fmt.strikeThrough);
    m_wrap      =new QCheckBox("Wrap Text");  m_wrap->setChecked(fmt.wrapText);
    auto* checkRow=new QHBoxLayout;
    for(auto*c:{m_bold,m_italic,m_underline,m_strike,m_wrap}) checkRow->addWidget(c);
    checkRow->addStretch();
    fg->addLayout(checkRow,1,0,1,4);
    main->addWidget(fontGroup);

    // ── Format + Alignment row ───────────────────────────────────────────
    auto* fmtGroup=new QGroupBox("Format & Alignment");
    auto* fag=new QGridLayout(fmtGroup);

    fag->addWidget(new QLabel("Number Format:"),0,0);
    m_numFmt=new QComboBox;
    m_numFmt->addItems({"General","Number","Currency ($)","Accounting","Percentage (%)","Scientific","Short Date","Long Date","Time","Fraction","Text"});
    m_numFmt->setCurrentIndex((int)fmt.numFmt); fag->addWidget(m_numFmt,0,1);

    fag->addWidget(new QLabel("Decimals:"),0,2);
    m_decimals=new QSpinBox; m_decimals->setRange(0,10); m_decimals->setValue(fmt.decimals); fag->addWidget(m_decimals,0,3);

    fag->addWidget(new QLabel("H-Align:"),1,0);
    m_hAlign=new QComboBox; m_hAlign->addItems({"Left","Center","Right"});
    if(fmt.hAlign==Qt::AlignHCenter)m_hAlign->setCurrentIndex(1);
    else if(fmt.hAlign==Qt::AlignRight)m_hAlign->setCurrentIndex(2);
    fag->addWidget(m_hAlign,1,1);

    fag->addWidget(new QLabel("V-Align:"),1,2);
    m_vAlign=new QComboBox; m_vAlign->addItems({"Top","Middle","Bottom"});
    if(fmt.vAlign==Qt::AlignTop)m_vAlign->setCurrentIndex(0);
    else if(fmt.vAlign==Qt::AlignBottom)m_vAlign->setCurrentIndex(2);
    else m_vAlign->setCurrentIndex(1);
    fag->addWidget(m_vAlign,1,3);
    main->addWidget(fmtGroup);

    // ── Colors ────────────────────────────────────────────────────────────
    auto* colGroup=new QGroupBox("Colors");
    auto* cg=new QHBoxLayout(colGroup);

    auto mkColorRow=[&](const QString& label, QLabel*& sw, QColor& col){
        auto* vl2=new QVBoxLayout; vl2->addWidget(new QLabel(label));
        sw=new QLabel; sw->setFixedSize(80,24);
        sw->setStyleSheet(QString("background:%1;border:1px solid #ccc;border-radius:3px;").arg(col.name()));
        auto* btn=new QPushButton("Change…"); btn->setFixedHeight(24);
        connect(btn,&QPushButton::clicked,this,[this,&col,sw](){
            QColor c=QColorDialog::getColor(col,this,"Select Color");
            if(c.isValid()){col=c;sw->setStyleSheet(QString("background:%1;border:1px solid #ccc;border-radius:3px;").arg(c.name()));}
        });
        auto* hl2=new QHBoxLayout; hl2->addWidget(sw); hl2->addWidget(btn);
        vl2->addLayout(hl2); cg->addLayout(vl2);
    };
    mkColorRow("Text Color:", m_textColorSw, m_textColor);
    cg->addSpacing(16);
    mkColorRow("Fill Color:", m_fillColorSw, m_fillColor);
    cg->addStretch();
    main->addWidget(colGroup);

    // ── Buttons ───────────────────────────────────────────────────────────
    auto* bb=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel);
    bb->button(QDialogButtonBox::Ok)->setProperty("primary","true");
    connect(bb,&QDialogButtonBox::accepted,this,[this]{ apply(); accept(); });
    connect(bb,&QDialogButtonBox::rejected,this,&QDialog::reject);
    main->addWidget(bb);
}

void FormatCellsDialog::apply() {
    m_fmt.fontFamily    = m_fontFam->currentText();
    m_fmt.fontSize      = m_fontSize->value();
    m_fmt.bold          = m_bold->isChecked();
    m_fmt.italic        = m_italic->isChecked();
    m_fmt.underline     = m_underline->isChecked();
    m_fmt.strikeThrough = m_strike->isChecked();
    m_fmt.wrapText      = m_wrap->isChecked();
    m_fmt.numFmt        = (NumFmt)m_numFmt->currentIndex();
    m_fmt.decimals      = m_decimals->value();
    static const Qt::AlignmentFlag ha[]={Qt::AlignLeft,Qt::AlignHCenter,Qt::AlignRight};
    static const Qt::AlignmentFlag va[]={Qt::AlignTop,Qt::AlignVCenter,Qt::AlignBottom};
    m_fmt.hAlign        = ha[m_hAlign->currentIndex()];
    m_fmt.vAlign        = va[m_vAlign->currentIndex()];
    m_fmt.textColor     = m_textColor;
    m_fmt.fillColor     = m_fillColor;
}
