#ifndef EXCELFILE_H
#define EXCELFILE_H

#include <QString>
#include <QAxObject>


class ExcelFile
{
public:
    ExcelFile(const QString& filename);

    //关闭文件
    void close();

    //获得某一单元格的属性
    QVariant getProperty(const QString& index);
    QVariant getProperty(int row, int col);

private:
    QAxObject* excel;
    QAxObject* wordBooks;
    QAxObject* wordBook;
    QAxObject* sheets;
    QAxObject* sheet;
};

#endif // EXCELFILE_H
