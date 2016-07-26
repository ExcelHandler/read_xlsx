#include "mainwindow.h"
#include <QApplication>
#include <QAxWidget>
#include <QAxObject>
#include <QDebug>
#include <QVariant>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    MainWindow w;
    w.show();

    QAxObject *excel = new QAxObject("Excel.Application");
    excel->setProperty("Visible", false);
    // obj->setProperty("Caption", "Hello world");
    QAxObject *workBooks = excel->querySubObject("Workbooks");
    QAxObject *workBook = workBooks->querySubObject("Open(QString)",
         "C:\\zzzz\\Personal_Plan_Shi_JW_201607.xlsx");

    //QAxObject *workBook = workBooks->querySubObject("Item(const int)", 1);
    QAxObject *sheets = workBook->querySubObject("Sheets");
    QAxObject *sheet = sheets->querySubObject("Item(int)", 1);

    QAxObject *range = sheet->querySubObject("Range(const QVariant&)", QVariant(QString("A1:A1")));
    qDebug() << range->property("Value");
    range = sheet->querySubObject("Range(const QVariant&)", QVariant(QString("D7:D7")));
    qDebug() << range->property("Value");

    workBook->dynamicCall("Close (Boolean)", false);
    excel->dynamicCall("Quit (void)");

    // range->dynamicCall("Clear()");
    // range->dynamicCall("SetValue(const QVariant&)", QVariant(5));
    // obj->dynamicCall("SetScreenUpdating(bool)", true);
    return a.exec();
}
