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
    void update(int step);

public slots:
    void ExcelHandlerSlot(const QStringList&);

private slots:
    void stopWorking() { isWorking = false; }


private:
    void readFromOneExcel(const QString& filePath);

    MainWindow* _root;
    bool isWorking;
};

#endif // EXCELHANDLER_H
