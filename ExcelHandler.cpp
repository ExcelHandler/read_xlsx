#include "ExcelHandler.h"
#include "ExcelFile.h"
#include <QProgressDialog>
#include <QDebug>

ExcelHandler::ExcelHandler(MainWindow *root, QObject *parent) : _root(root), QObject(parent), isWorking(true)
{

}


void ExcelHandler::ExcelHandlerSlot(const QStringList &selected)
{
    isWorking = true;
    int sum = selected.length();
    for (int i = 0; i < sum; i++)
    {

        qDebug() << selected[i] << endl;

        if (!isWorking)
            break;
        emit update(i + 1);
    }
}

void ExcelHandler::readFromOneExcel(const QString& filePath)
{
    ExcelFile file(filePath);

}
