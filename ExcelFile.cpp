#include "ExcelFile.h"

ExcelFile::ExcelFile(const QString& filename)
{
    excel = new QAxObject("Excel.Application");
    excel->setProperty("Visible", true);

    wordBooks = excel -> querySubObject("Workbooks");
    wordBook = wordBooks -> querySubObject("Open(QString)",filename);
    sheets = wordBook->querySubObject("Sheets");
    sheet = sheets->querySubObject("Item(int)", 1);
}


void ExcelFile::close()
{
    wordBook->dynamicCall("Close (Boolean)", false);
    excel->dynamicCall("Quit (void)");
}


QVariant ExcelFile::getProperty(const QString &index)
{
    QString cellRange = index + ":" + index;
    QAxObject *range = sheet -> querySubObject("Range(const QVariant&)", QVariant(cellRange));
    return range -> property("Value");
}

QVariant ExcelFile::getProperty(int row, int col)
{
    QChar colIndex = (QChar)('A' + col - 1);
    QString index = colIndex + QString::number(row);
    return getProperty(index);
}
