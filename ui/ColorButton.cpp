#include "ColorButton.h"
#include <QPainter>
#include <QColorDialog>
#include <QMouseEvent>

ColorButton::ColorButton(const QIcon& icon, const QString& tip, QWidget* parent)
    : QToolButton(parent)
{
    setIcon(icon); setToolTip(tip);
    setFixedSize(28,28); setAutoRaise(true);
    setStyleSheet(
        "QToolButton{border:1px solid transparent;border-radius:3px;background:transparent;}"
        "QToolButton:hover{background:#e8f5ee;border-color:#c8e8d8;}"
        "QToolButton:pressed{background:#c8e8d8;border-color:#1e7145;}");
}

void ColorButton::setColor(const QColor& c) {
    if(m_color==c) return;
    m_color=c; update(); emit colorChanged(c);
}

void ColorButton::paintEvent(QPaintEvent* e) {
    QToolButton::paintEvent(e);
    QPainter p(this);
    p.setRenderHint(QPainter::Antialiasing);
    QRect sw(5, height()-6, width()-10, 4);
    p.fillRect(sw, m_color);
    p.setPen(QPen(QColor("#ccc"),0.5));
    p.drawRect(sw);
}

void ColorButton::mousePressEvent(QMouseEvent* e) {
    QToolButton::mousePressEvent(e);
    showPicker();
}

void ColorButton::showPicker() {
    QColor c = QColorDialog::getColor(m_color, this, "Select Color",
                                       QColorDialog::ShowAlphaChannel);
    if(c.isValid()) setColor(c);
}
