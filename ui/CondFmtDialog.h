#pragma once
#include <QDialog>
// Conditional formatting and cell styles handled inline in MainWindow
class CondFmtDialog : public QDialog { Q_OBJECT public: explicit CondFmtDialog(QWidget*p=nullptr):QDialog(p){} };
class CellStyleDialog : public QDialog { Q_OBJECT public: explicit CellStyleDialog(QWidget*p=nullptr):QDialog(p){} };
