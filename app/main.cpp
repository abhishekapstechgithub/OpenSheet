#include <QApplication>
#include <QSplashScreen>
#include <QTimer>
#include <QPainter>
#include "MainWindow.h"
#include "Theme.h"

int main(int argc, char* argv[]) {
    QApplication app(argc, argv);

    app.setApplicationName("OpenSheet ET");
    app.setApplicationVersion("1.0.0");
    app.setOrganizationName("OpenSheet");
    app.setOrganizationDomain("opensheet.app");
    app.setStyleSheet(Theme::appStyle());

    // Splash screen
    QPixmap splash(480, 280);
    splash.fill(QColor(Theme::GREEN_DARK));
    QPainter sp(&splash);
    sp.setRenderHint(QPainter::Antialiasing);
    // Logo icon
    sp.setBrush(QColor(255,255,255,40)); sp.setPen(Qt::NoPen);
    sp.drawRoundedRect(30,30,56,56,8,8);
    sp.setPen(QPen(Qt::white,2.5)); sp.setBrush(Qt::NoBrush);
    sp.drawLine(58,42,58,74); sp.drawLine(42,58,74,58);
    sp.drawRect(38,38,20,20); sp.drawRect(60,38,20,20);
    sp.drawRect(38,60,20,20); sp.drawRect(60,60,20,20);
    // Title
    QFont titleFont("Segoe UI",24,QFont::Light); sp.setFont(titleFont);
    sp.setPen(Qt::white); sp.drawText(100,65,"OpenSheet ET");
    QFont subFont("Segoe UI",11); sp.setFont(subFont);
    sp.setPen(QColor(255,255,255,180)); sp.drawText(100,90,"Professional Spreadsheet");
    // Version
    sp.setPen(QColor(255,255,255,120)); QFont vf("Segoe UI",9); sp.setFont(vf);
    sp.drawText(30,260,"Version 1.0  |  Powered by Qt 6 & C++20");
    sp.end();

    QSplashScreen* ss = new QSplashScreen(splash);
    ss->show(); app.processEvents();
    ss->showMessage("Loading formula engine…",Qt::AlignBottom|Qt::AlignHCenter,QColor(255,255,255,200));
    app.processEvents();

    MainWindow win;

    QTimer::singleShot(1200, &win, [ss,&win]{
        ss->finish(&win); delete ss;
        win.show();
    });

    return app.exec();
}
