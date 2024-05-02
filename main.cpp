#include "canalysedata.h"
#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    CAnalyseData w;

    QSettings settings(QString(QApplication::applicationDirPath()+"\\Setting.ini"), QSettings::IniFormat);
    settings.beginGroup("PathSetting");
    w.resize(settings.value("WindowWidth").toInt(), settings.value("WindowHeight").toInt());

    w.show();

    return a.exec();
}
