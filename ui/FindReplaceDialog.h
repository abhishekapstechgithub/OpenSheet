#pragma once
#include <QDialog>
class QLineEdit; class QCheckBox; class QLabel;
class Sheet;

class FindReplaceDialog : public QDialog {
    Q_OBJECT
public:
    explicit FindReplaceDialog(Sheet* sheet, int& curRow, int& curCol, QWidget* parent=nullptr);
private:
    Sheet*      m_sheet; int& m_row; int& m_col;
    QLineEdit  *m_find, *m_replace;
    QCheckBox  *m_matchCase, *m_entireCell;
    QLabel*     m_msg;
    void findNext(); void replaceOne(); void replaceAll();
};
