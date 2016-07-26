#include "ExcelHandler.h"
#include "ExcelFile.h"
#include "ui_excelhandler.h"
#include <QProgressDialog>
#include <QFileDialog>
#include <QDate>
#include <QDebug>
#include <QApplication>

ExcelHandler::ExcelHandler(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::ExcelHandler)
{
    ui->setupUi(this);
    int defaultMonth = QDate::currentDate().month();
    ui -> month_box -> setRange(1, 12);
    ui -> month_box -> setValue(defaultMonth);
    connect(ui -> file_btn, SIGNAL(clicked()), this, SLOT(selectFiles()));
}

ExcelHandler::~ExcelHandler()
{
    delete ui;
}


void ExcelHandler::selectFiles()
{
    QFileDialog *fileDialog = new QFileDialog(this);
    fileDialog -> setWindowTitle(tr("文件位置"));
    fileDialog -> setDirectory(".");
    fileDialog -> setNameFilter("*.xls *.xlsx *.csv");
    fileDialog -> setFileMode(QFileDialog::ExistingFiles);

    if(fileDialog -> exec() == QDialog::Accepted)
    {
        QStringList selected = fileDialog -> selectedFiles();
        importFiles(selected);
    }
}

void ExcelHandler::importFiles(const QStringList &selected)
{
    int numFiles = selected.length();
    QProgressDialog progress("Importing files...", "Abort Import", 0, numFiles, this);
    for (int i = 0; i < numFiles; i++)
    {
        progress.setValue(i);
        if (progress.wasCanceled()) break;
        qDebug() << "MDZZ" << endl;
        //开始处理
        qApp -> processEvents();
        processFile(selected[i]);
    }
    progress.setValue(numFiles);
    this -> close();
}

void ExcelHandler::processFile(const QString &filename)
{
    ExcelFile file(filename);
    file.close();
}
