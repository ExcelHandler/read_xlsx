#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
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

private slots:
    void on_actionImport_triggered();

private:
    Ui::MainWindow *ui;
    ExcelHandler handler;
    QThread importThread;
};

#endif // MAINWINDOW_H
