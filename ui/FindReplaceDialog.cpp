#include "FindReplaceDialog.h"
#include "Sheet.h"
#include "FormulaEngine.h"
#include "Theme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QCheckBox>
#include <QPushButton>

FindReplaceDialog::FindReplaceDialog(Sheet* sheet, int& curRow, int& curCol, QWidget* parent)
    : QDialog(parent), m_sheet(sheet), m_row(curRow), m_col(curCol)
{
    setWindowTitle("Find & Replace");
    setMinimumWidth(380);
    setStyleSheet(Theme::appStyle());
    auto* vl=new QVBoxLayout(this);

    auto mkRow=[&](const QString& label, QLineEdit*& edit, const QString& ph){
        auto* hl=new QHBoxLayout; auto* lbl=new QLabel(label); lbl->setFixedWidth(90);
        edit=new QLineEdit; edit->setPlaceholderText(ph);
        hl->addWidget(lbl); hl->addWidget(edit); vl->addLayout(hl);
    };
    mkRow("Find:", m_find, "Search text…");
    mkRow("Replace with:", m_replace, "Replace with…");

    auto* cRow=new QHBoxLayout;
    m_matchCase  =new QCheckBox("Match case");
    m_entireCell =new QCheckBox("Entire cell contents");
    cRow->addWidget(m_matchCase); cRow->addWidget(m_entireCell); cRow->addStretch();
    vl->addLayout(cRow);

    m_msg=new QLabel; m_msg->setStyleSheet("color:#454d5c;font-size:11px;min-height:20px;padding:4px;background:#f5f6f7;border-radius:3px;");
    vl->addWidget(m_msg);

    auto* btnRow=new QHBoxLayout;
    auto mk=[&](const QString& label, bool primary, auto fn){
        auto* btn=new QPushButton(label);
        if(primary) btn->setProperty("primary","true");
        connect(btn,&QPushButton::clicked,this,fn);
        btnRow->addWidget(btn);
    };
    mk("Find Next",    false,[this]{ findNext(); });
    mk("Replace",      false,[this]{ replaceOne(); });
    mk("Replace All",  true, [this]{ replaceAll(); });
    mk("Close",        false,[this]{ close(); });
    vl->addLayout(btnRow);
}

static int g_fr=-1, g_fc=-1;
void FindReplaceDialog::findNext() {
    const QString val=m_find->text(); if(val.isEmpty())return;
    bool mc=m_matchCase->isChecked(), wc=m_entireCell->isChecked();
    int rows=m_sheet->usedRowCount(), cols=m_sheet->usedColCount();
    for(int r=0;r<rows;r++) for(int c=0;c<cols;c++) {
        if(r<g_fr||(r==g_fr&&c<=g_fc)) continue;
        QString v=m_sheet->getCellValue(r,c).toString();
        bool match=wc?(mc?v==val:v.toLower()==val.toLower()):(mc?v.contains(val):v.toLower().contains(val.toLower()));
        if(match){ m_row=r; m_col=c; g_fr=r; g_fc=c; m_msg->setText("Found at "+FormulaEngine::cellRef(r,c)); return; }
    }
    m_msg->setText("Not found — wrapping to beginning"); g_fr=-1; g_fc=-1;
}

void FindReplaceDialog::replaceOne() {
    const QString fv=m_find->text(), rv=m_replace->text();
    const Cell& cell=m_sheet->getCell(m_row,m_col);
    QString v=cell.rawValue.toString();
    if(v.contains(fv)){ m_sheet->setCellValue(m_row,m_col,v.replace(fv,rv)); m_msg->setText("Replaced 1 occurrence"); }
    findNext();
}

void FindReplaceDialog::replaceAll() {
    const QString fv=m_find->text(), rv=m_replace->text(); int count=0;
    int rows=m_sheet->usedRowCount(), cols=m_sheet->usedColCount();
    for(int r=0;r<rows;r++) for(int c=0;c<cols;c++) {
        QString v=m_sheet->getCell(r,c).rawValue.toString();
        if(v.contains(fv)){ m_sheet->setCellValue(r,c,v.replace(fv,rv)); count++; }
    }
    m_msg->setText(QString("Replaced %1 occurrence(s)").arg(count));
}
