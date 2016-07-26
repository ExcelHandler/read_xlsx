#ifndef EXCELHANDLER_H
#define EXCELHANDLER_H

#include <QWidget>

namespace Ui {
class ExcelHandler;
}

class ExcelHandler : public QWidget
{
    Q_OBJECT

public:
    explicit ExcelHandler(QWidget *parent = 0);
    ~ExcelHandler();

private slots:
    void selectFiles();
    void importFiles(const QStringList&);
    void processFile(const QString& filename);

private:
    Ui::ExcelHandler *ui;
};

#endif // EXCELHANDLER_H
