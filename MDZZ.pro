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

SOURCES += main.cpp\
        mainwindow.cpp \
    Entrance.cpp \
    ExcelFile.cpp \
    ExcelHandler.cpp

HEADERS  += mainwindow.h \
    Entrance.h \
    ExcelFile.h \
    ExcelHandler.h

FORMS    += mainwindow.ui \
    entrance.ui
