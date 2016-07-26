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

    QAxObject *excel = new QAxObject("Excel.Application");
    excel->setProperty("Visible", true);
    // obj->setProperty("Caption", "Hello world");
    QAxObject *workBooks = excel->querySubObject("Workbooks");
    QAxObject *workBook = workBooks->querySubObject("Open(QString)",
         "C:\\zzzz\\Personal_Plan_Shi_JW_201607.xlsx");

    //QAxObject *workBook = workBooks->querySubObject("Item(const int)", 1);
    QAxObject *sheets = workBook->querySubObject("Sheets");
    QAxObject *sheet = sheets->querySubObject("Item(int)", 1);

    // Test title
    QAxObject *range = sheet->querySubObject("Range(const QVariant&)", QVariant(QString("A1:A1")));
    qDebug() << range->property("Value");
    // Test first date
    range = sheet->querySubObject("Range(const QVariant&)", QVariant(QString("D7:D7")));
    qDebug() << range->property("Value");
    // QDateTime time = QDateTime((QDateTime)(range->property("Value")).date(), (range->property("Value")).time());
    QDateTime time = (range->property("Value")).toDateTime();
    qDebug() << time.date();
    // Test second date
    range = sheet->querySubObject("Range(const QVariant&)", QVariant(QString("D119:D119")));
    qDebug() << (range->property("Value")).toDateTime().date();
    // Test empty date
    range = sheet->querySubObject("Range(const QVariant&)", QVariant(QString("D3:D3")));
    qDebug() << range->property("Value");

    workBook->dynamicCall("Close (Boolean)", false);
    excel->dynamicCall("Quit (void)");

    ExcelFile file("C:\\zzzz\\Personal_Plan_Shi_JW_201607.xlsx");
    qDebug() << file.getProperty("D3");
    file.close();


    // range->dynamicCall("Clear()");
    // range->dynamicCall("SetValue(const QVariant&)", QVariant(5));
    // obj->dynamicCall("SetScreenUpdating(bool)", true);
    return a.exec();
}
