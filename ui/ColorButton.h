#pragma once
#include <QToolButton>
#include <QColor>

// A button with a colored indicator bar underneath (WPS style)
class ColorButton : public QToolButton {
    Q_OBJECT
    Q_PROPERTY(QColor color READ color WRITE setColor NOTIFY colorChanged)
public:
    explicit ColorButton(const QIcon& icon, const QString& tip, QWidget* parent=nullptr);
    QColor color() const { return m_color; }
    void   setColor(const QColor& c);
signals:
    void colorChanged(const QColor& c);
protected:
    void paintEvent(QPaintEvent* e) override;
    void mousePressEvent(QMouseEvent* e) override;
private:
    QColor m_color { Qt::black };
    void showPicker();
};
