#include "FormulaBar.h"
#include <QHBoxLayout>
#include <QLineEdit>
#include <QLabel>
#include <QToolButton>
#include <QKeyEvent>

FormulaBar::FormulaBar(QWidget* parent) : QWidget(parent) {
    setFixedHeight(28);
    setStyleSheet("background:white;border-bottom:1px solid #d8dce3;");
    auto* hl=new QHBoxLayout(this); hl->setContentsMargins(6,2,6,2); hl->setSpacing(4);

    m_refBox=new QLineEdit; m_refBox->setFixedWidth(70); m_refBox->setFixedHeight(22);
    m_refBox->setAlignment(Qt::AlignHCenter);
    m_refBox->setStyleSheet("QLineEdit{border:1px solid #c4c9d4;border-radius:3px;background:#f5f6f8;"
                             "font-family:'Courier New';font-size:12px;font-weight:600;color:#1e7145;padding:0 4px;}"
                             "QLineEdit:focus{border-color:#1e7145;background:white;}");

    auto* fxBtn=new QToolButton; fxBtn->setText("ƒx"); fxBtn->setFixedSize(30,22);
    fxBtn->setStyleSheet("QToolButton{border:1px solid #d8dce3;border-radius:3px;background:#f5f6f8;"
                          "font-size:13px;font-style:italic;font-weight:800;color:#1e7145;font-family:Georgia,serif;}"
                          "QToolButton:hover{background:#e8f5ee;border-color:#1e7145;}");

    auto* div=new QFrame; div->setFrameShape(QFrame::VLine); div->setFixedWidth(1);
    div->setStyleSheet("color:#d8dce3;");

    m_formulaEdit=new QLineEdit; m_formulaEdit->setFixedHeight(22);
    m_formulaEdit->setPlaceholderText("Enter a value or formula  (e.g.  =SUM(A1:A10))");
    m_formulaEdit->setStyleSheet("QLineEdit{border:1px solid #c4c9d4;border-radius:3px;background:white;"
                                  "font-family:'Courier New';font-size:12px;padding:0 8px;}"
                                  "QLineEdit:focus{border-color:#1e7145;box-shadow:0 0 0 2px rgba(30,113,69,.1);}");

    hl->addWidget(m_refBox); hl->addWidget(fxBtn); hl->addWidget(div); hl->addWidget(m_formulaEdit,1);

    connect(m_refBox,&QLineEdit::returnPressed,this,[this]{ emit cellRefNavigated(m_refBox->text().trimmed().toUpper()); });
    connect(m_formulaEdit,&QLineEdit::returnPressed,this,[this]{ emit formulaCommitted(m_formulaEdit->text()); });
    connect(fxBtn,&QToolButton::clicked,this,&FormulaBar::functionWizardRequested);
}

void FormulaBar::setCellRef(const QString& ref)    { m_refBox->setText(ref); }
void FormulaBar::setFormulaText(const QString& t)  { QSignalBlocker b(m_formulaEdit); m_formulaEdit->setText(t); }
QString FormulaBar::formulaText() const             { return m_formulaEdit->text(); }
