#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include <QDebug>
#include <QTreeView>
#include <QListView>
#include <QProgressDialog>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow), handler(this), progress(NULL)
{
    ui->setupUi(this);
    initConnect();
    handler.moveToThread(&importThread);
    importThread.start();
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::initConnect()
{
    connect(this, SIGNAL(selectFinished(const QStringList&)), &handler, SLOT(ExcelHandlerSlot(const QStringList&)), Qt::QueuedConnection);
    connect(&handler, SIGNAL(update(int)), this, SLOT(updateProgressBar(int)));
    connect(this, SIGNAL(stopExcelHandler()), &handler, SLOT(stopWorking()));
}

void MainWindow::updateProgressBar(int step)
{
    progress -> setValue(step);
    qDebug() << "Update progress bar " << step;
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
        int sum = selected.length();
        progress = new QProgressDialog(tr("Importing files..."), tr("Abort Import"),
                                 0, sum, this);
        progress -> show();

        for (auto i = selected.begin(); i != selected.end(); ++i)
            qDebug() << *i << endl;
    }


}
