#pragma once
#include <QWidget>
class QLineEdit; class QLabel;

class FormulaBar : public QWidget {
    Q_OBJECT
public:
    explicit FormulaBar(QWidget* parent=nullptr);
    void setCellRef(const QString& ref);
    void setFormulaText(const QString& text);
    QString formulaText() const;
signals:
    void formulaCommitted(const QString& text);
    void cellRefNavigated(const QString& ref);
    void functionWizardRequested();
private:
    QLineEdit* m_refBox;
    QLineEdit* m_formulaEdit;
};
