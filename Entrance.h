#ifndef ENTRANCE_H
#define ENTRANCE_H

#include <QWidget>

namespace Ui {
class Entrance;
}

class Entrance : public QWidget
{
    Q_OBJECT

public:
    explicit Entrance(QWidget *parent = 0);
    ~Entrance();

private:
    Ui::Entrance *ui;
};

#endif // ENTRANCE_H
