#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include <QDebug>
#include <QTreeView>
#include <QListView>
#include <QProgressDialog>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow), handler(this)
{
    ui->setupUi(this);
    connect(this, SIGNAL(selectFinished(const QStringList&)), &handler, SLOT(ExcelHandlerSlot(const QStringList&)), Qt::QueuedConnection);
    handler.moveToThread(&importThread);
    importThread.start();
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_actionImport_triggered()
{
    QFileDialog *fileDialog = new QFileDialog(this);
    fileDialog -> setWindowTitle(tr("文件位置"));
    fileDialog -> setDirectory(".");
    fileDialog -> setNameFilter("*.xls *.xlsx *.csv");
    fileDialog -> setFileMode(QFileDialog::ExistingFiles);

    if(fileDialog -> exec() == QDialog::Accepted)
    {
        QStringList selected = fileDialog -> selectedFiles();
        emit selectFinished(selected);

        for (auto i = selected.begin(); i != selected.end(); ++i)
            qDebug() << *i << endl;
    }


}
