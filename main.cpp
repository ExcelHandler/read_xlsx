#include <QtGui>
#include <QApplication>
#include <QAxObject>
#include <QAxWidget>
#include <qaxselect.h>
int main(int argc, char **argv)
{
    QApplication a(argc, argv);
    QAxSelect select;
    select.show();
    QAxWidget excel("Excel.Application");
    excel.setProperty("Visible", true);
    QAxObject * workbooks = excel.querySubObject("WorkBooks");
    workbooks->dynamicCall("Open (const QString&)", QString("c:/test.xls"));
    QAxObject * workbook = excel.querySubObject("ActiveWorkBook");
    QAxObject * worksheets = workbook->querySubObject("WorkSheets");
    int intCount = worksheets->property("Count").toInt();
    for (int i = 1; i <= intCount; i++)
    {
        int intVal;
        QAxObject * worksheet = workbook->querySubObject("Worksheets(int)", i);
        qDebug() << i << worksheet->property("Name").toString();
        QAxObject * range = worksheet->querySubObject("Cells(1,1)");
        intVal = range->property("Value").toInt();
        range->setProperty("Value", QVariant(intVal+1));
        QAxObject * range2 = worksheet->querySubObject("Range(C1)");
        intVal = range2->property("Value").toInt();
        range2->setProperty("Value", QVariant(intVal+1));
    }
    QAxObject * worksheet = workbook->querySubObject("Worksheets(int)", 1);
    QAxObject * usedrange = worksheet->querySubObject("UsedRange");
    QAxObject * rows = usedrange->querySubObject("Rows");
    QAxObject * columns = usedrange->querySubObject("Columns");
    int intRowStart = usedrange->property("Row").toInt();
    int intColStart = usedrange->property("Column").toInt();
    int intCols = columns->property("Count").toInt();
    int intRows = rows->property("Count").toInt();
    for (int i = intRowStart; i < intRowStart + intRows; i++)
    {
        for (int j = intColStart; j <= intColStart + intCols; j++)
        {
            QAxObject * range = worksheet->querySubObject("Cells(int,int)", i, j );
            qDebug() << i << j << range->property("Value");
        }
    }
    excel.setProperty("DisplayAlerts", 0);
    workbook->dynamicCall("SaveAs (const QString&)", QString("c:/xlsbyqt.xls"));
    excel.setProperty("DisplayAlerts", 1);
    workbook->dynamicCall("Close (Boolean)", false);
    excel.dynamicCall("Quit (void)");
    return a.exec();
}
