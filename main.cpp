#include "mainwindow.h"
#include "ExcelFile.h"
#include <QApplication>
#include <QAxWidget>
#include <QAxObject>
#include <QDebug>
#include <QVariant>
#include <QDateTime>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    MainWindow w;
    w.show();

    /*
    ExcelFile file("C:\\zzzz\\Personal_Plan_Shi_JW_201607.xlsx");
    qDebug() << file.getProperty("G3");
    qDebug() << file.getProperty(3, 7);
    file.close();
    */

    return a.exec();
}
