#pragma once
#include <QWidget>
class QLabel; class QSlider;
class OSStatusBar : public QWidget {
    Q_OBJECT
public:
    explicit OSStatusBar(QWidget* parent=nullptr);
    void setMode(const QString& m);
    void setStats(const QString& stats);
    void setZoom(int pct);
signals:
    void zoomChanged(int pct);
private:
    QLabel*  m_mode; QLabel* m_stats; QLabel* m_zoomLabel; QSlider* m_zoomSlider;
};
