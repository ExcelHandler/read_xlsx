#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QProgressDialog>
#include <QThread>
#include "ExcelHandler.h"

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();


signals:
    selectFinished(const QStringList& selected);
    void stopExcelHandler();


private slots:
    void on_actionImport_triggered();

public slots:
    void updateProgressBar(int step);


private:
    void initConnect();

private:
    Ui::MainWindow *ui;
    ExcelHandler handler;
    QThread importThread;
    QProgressDialog* progress;
};

#endif // MAINWINDOW_H
