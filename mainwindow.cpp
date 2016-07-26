#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "ExcelHandler.h"
#include <QFileDialog>
#include <QInputDialog>
#include <QDebug>
#include <QTreeView>
#include <QListView>
#include <QProgressDialog>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow), handler(NULL)
{
    ui->setupUi(this);
    initConnect();
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::initConnect()
{

}

void MainWindow::on_actionImport_triggered()
{
    handler = new ExcelHandler();
    handler -> show();
}
