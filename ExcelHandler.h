#ifndef EXCELHANDLER_H
#define EXCELHANDLER_H

#include <QObject>

class MainWindow;

class ExcelHandler : public QObject
{
    Q_OBJECT
public:
    explicit ExcelHandler(MainWindow* root, QObject *parent = 0);

signals:
    void ExcelHandlerSignal();

public slots:
    void ExcelHandlerSlot(const QStringList&);


private:
    MainWindow* _root;
};

#endif // EXCELHANDLER_H
