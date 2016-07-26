#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QProgressDialog>

namespace Ui {
class MainWindow;
}

class ExcelHandler;
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

private:
    void initConnect();

private:
    Ui::MainWindow *ui;
    ExcelHandler* handler;
};

#endif // MAINWINDOW_H
