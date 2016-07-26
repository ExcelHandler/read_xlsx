#include "mainwindow.h"
#include <QApplication>
#include <QAxWidget>
#include <QAxObject>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    MainWindow w;
    w.show();

    QAxObject *obj = new QAxObject("Excel.Application");
    obj->setProperty("Visible", true);
    obj->setProperty("Caption", "Hello world");
    QAxObject *workBooks = obj->querySubObject("Workbooks");
    //打开已存的excel文件
    QAxObject *workBook =
         workBooks->querySubObject("Open(QString)",
         "C:\\Users\\lenovo\\Documents\\m_Internship\\Personal_Plan_Shi_JW_201607.xlsx");

    workBooks->dynamicCall("Add");
    //QAxObject *workBook = workBooks->querySubObject("Item(const int)", 1);
    QAxObject *sheets = workBook->querySubObject("Sheets");
    QAxObject *sheet = sheets->querySubObject("Item(int)", 1);
    QAxObject *range = sheet->querySubObject("Range(const QVariant&)", QVariant(QString("A1:A1")));
    range->dynamicCall("Clear()");
    range->dynamicCall("SetValue(const QVariant&)", QVariant(5));
    obj->dynamicCall("SetScreenUpdating(bool)", true);
    return a.exec();
}
