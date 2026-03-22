#pragma once
#include <QDialog>
#include <QString>

class FunctionWizard : public QDialog {
    Q_OBJECT
public:
    explicit FunctionWizard(QWidget* parent=nullptr);
    QString result() const { return m_result; }
private:
    QString m_result;
};
