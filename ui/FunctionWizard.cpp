#include "FunctionWizard.h"
#include "Theme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QListWidget>
#include <QLabel>
#include <QLineEdit>
#include <QPushButton>
#include <QDialogButtonBox>

FunctionWizard::FunctionWizard(QWidget* parent) : QDialog(parent) {
    setWindowTitle("Insert Function"); setMinimumSize(480,360);
    setStyleSheet(Theme::appStyle());
    auto* vl=new QVBoxLayout(this);
    auto* hl=new QHBoxLayout;

    // Left: search + list
    auto* left=new QVBoxLayout;
    auto* srch=new QLineEdit; srch->setPlaceholderText("Search functions…");
    left->addWidget(srch);
    auto* list=new QListWidget;
    const QStringList fns={"SUM","AVERAGE","MIN","MAX","COUNT","COUNTA","IF","IFS","AND","OR","NOT","IFERROR",
        "ROUND","ROUNDUP","ROUNDDOWN","ABS","SQRT","POWER","MOD","INT","EXP","LN","LOG","PI",
        "SIN","COS","TAN","RAND","RANDBETWEEN","PRODUCT","STDEV","STDEVP","VAR","MEDIAN","MODE",
        "LARGE","SMALL","RANK","LEN","LEFT","RIGHT","MID","UPPER","LOWER","PROPER","TRIM",
        "CONCATENATE","TEXTJOIN","SUBSTITUTE","FIND","SEARCH","VALUE","TEXT","REPT","CHAR","CODE",
        "TODAY","NOW","YEAR","MONTH","DAY","HOUR","DATEDIF","EDATE","EOMONTH",
        "ISBLANK","ISNUMBER","ISTEXT","ISERROR","CHOOSE","VLOOKUP","HLOOKUP"};
    list->addItems(fns); list->setCurrentRow(0);
    left->addWidget(list,1);

    // Right: description + args
    auto* right=new QVBoxLayout;
    auto* desc=new QLabel("Select a function to see its description."); desc->setWordWrap(true);
    desc->setStyleSheet("padding:8px;background:#f5f6f7;border-radius:4px;font-size:12px;color:#454d5c;min-height:80px;");
    right->addWidget(desc);
    auto* argsLabel=new QLabel("Arguments (comma-separated):"); right->addWidget(argsLabel);
    auto* argsEdit=new QLineEdit; argsEdit->setPlaceholderText("e.g. A1:A10"); right->addWidget(argsEdit);
    auto* prevLabel=new QLabel; prevLabel->setStyleSheet("padding:6px;background:#e8f5ee;border:1px solid #c8e8d8;border-radius:3px;font-family:'Courier New';color:#1e7145;");
    right->addWidget(prevLabel); right->addStretch();

    hl->addLayout(left,1); hl->addSpacing(12); hl->addLayout(right,2);
    vl->addLayout(hl);

    auto* bb=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel);
    bb->button(QDialogButtonBox::Ok)->setProperty("primary","true");
    vl->addWidget(bb);

    static const QHash<QString,QString> DESCS{
        {"SUM","Adds all numbers in a range. =SUM(A1:A10)"},
        {"AVERAGE","Returns the average value. =AVERAGE(B1:B20)"},
        {"IF","Conditional: =IF(condition,value_if_true,value_if_false)"},
        {"VLOOKUP","Vertical lookup: =VLOOKUP(lookup,range,col_index,exact)"},
        {"TODAY","Returns today's date. =TODAY()"},
        {"NOW","Returns current date and time. =NOW()"},
        {"COUNTIF","Counts cells matching criteria: =COUNTIF(range,criteria)"},
        {"SUMIF","Sums cells matching criteria: =SUMIF(range,criteria,sum_range)"},
        {"TEXTJOIN","Joins text with delimiter: =TEXTJOIN(\",\",TRUE,A1:A5)"},
        {"STDEV","Sample standard deviation of a dataset"},
        {"MEDIAN","Middle value in a sorted range"},
    };

    auto updatePreview=[=](){
        QString fn=list->currentItem()?list->currentItem()->text():"";
        QString args=argsEdit->text();
        prevLabel->setText("=" + fn + "(" + (args.isEmpty()?"…":args) + ")");
    };

    connect(list,&QListWidget::currentTextChanged,this,[=](const QString& fn){
        desc->setText(DESCS.value(fn, fn+"(...)"));
        updatePreview();
    });
    connect(argsEdit,&QLineEdit::textChanged,this,updatePreview);
    connect(srch,&QLineEdit::textChanged,this,[list](const QString& q){
        for(int i=0;i<list->count();i++)
            list->item(i)->setHidden(!list->item(i)->text().toLower().contains(q.toLower()));
    });
    connect(bb,&QDialogButtonBox::accepted,this,[this,list,argsEdit]{
        if(list->currentItem())
            m_result="="+list->currentItem()->text()+"("+(argsEdit->text().isEmpty()?"":argsEdit->text())+")";
        accept();
    });
    connect(bb,&QDialogButtonBox::rejected,this,&QDialog::reject);

    // trigger initial preview
    if(list->currentItem()) emit list->currentTextChanged(list->currentItem()->text());
}
