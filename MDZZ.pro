#-------------------------------------------------
#
# Project created by QtCreator 2016-07-25T15:43:21
#
#-------------------------------------------------

QT       += core gui
QT       += sql
QT       += axcontainer

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = MDZZ
TEMPLATE = app

include(3rdparty/qtxlsx/src/xlsx/qtxlsx.pri)

SOURCES += main.cpp\
        mainwindow.cpp \
    Entrance.cpp

HEADERS  += mainwindow.h \
    Entrance.h

FORMS    += mainwindow.ui \
    entrance.ui
