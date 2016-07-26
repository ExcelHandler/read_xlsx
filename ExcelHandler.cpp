#include "ExcelHandler.h"
#include <QProgressDialog>
#include <QDebug>

ExcelHandler::ExcelHandler(MainWindow *root, QObject *parent) : _root(root), QObject(parent)
{

}


void ExcelHandler::ExcelHandlerSlot(const QStringList &selected)
{
    int sum = selected.length();
    QProgressDialog progress(tr("Importing files..."), tr("Abort Import"),
                             0, sum);
    for (int i = 0; i < sum; i++)
    {
        progress.setValue(i);
        qDebug() << selected[i] << endl;

        if (progress.wasCanceled())
            break;
    }
}
