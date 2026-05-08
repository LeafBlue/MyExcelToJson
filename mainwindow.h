#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QMap>
#include <QString>

#include "xlsxdocument.h"
#include "xlsxcell.h"
#include "xlsxcellrange.h"

#include "json.hpp"

using json = nlohmann::json;

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE


class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    void parseExcel(QXlsx::Document &xlsx);

    QStringList getRowData(QXlsx::Document &xlsx, int row);

    void handletext(QXlsx::Document &xlsx);

    // jsonObj 改为 json 引用
    void handlelinetype(json& jsonObj, int lineindex, QString linename, QString text);

    //辅助
    QString getnextlinename(QString input);

    bool isIntegerString(const QString& str);

private slots:
    void on_pushButton_clicked();
    void on_pushButton_2_clicked();

private:


    void loadSheetNames(const QString &fileName);

private:
    Ui::MainWindow *ui;

    QStringList firstline;

    QStringList secondline;

    int startline = 1;

    json jsonArray = json::array();

};

#endif // MAINWINDOW_H
