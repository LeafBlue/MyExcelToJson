#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QMessageBox>
#include <QDebug>
#include <QFileDialog>
#include <QFile>
#include <QTextStream>
#include <QCoreApplication>



using namespace QXlsx;

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    setWindowTitle("Excel转Json");
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::parseExcel(QXlsx::Document &xlsx)
{

    jsonArray.clear();

    firstline.clear();
    firstline = getRowData(xlsx,1);

    secondline.clear();
    secondline = getRowData(xlsx,2);

    handletext(xlsx);

}

QStringList MainWindow::getRowData(QXlsx::Document &xlsx, int row)
{
    QStringList rowData;
    int col = 1;

    while (true) {
        auto *cell = xlsx.cellAt(row, col);

        if (!cell || cell->value().toString().trimmed().isEmpty()) {
            break;
        }

        rowData << cell->value().toString();
        col++;

        if (col >5000) break;
    }

    return rowData;
}

void MainWindow::handletext(QXlsx::Document &xlsx)
{
    int colCount = firstline.size();
    if (colCount == 0) return;

    int currentRow = 4;
    QString lastId = "";

    // 使用索引访问，彻底避免指针/引用失效问题
    int currentObjIndex = -1;

    while (true) {
        // 1. 检查当前行是否有效
        bool hasContent = false;
        for (int c = 1; c <= colCount; ++c) {
            auto *cell = xlsx.cellAt(currentRow, c);
            if (cell && !cell->value().toString().trimmed().isEmpty()) {
                hasContent = true; break;
            }
        }
        if (!hasContent) break;

        // 2. 获取主键（A列）
        auto *idCell = xlsx.cellAt(currentRow, 1);
        QString currentId = idCell ? idCell->value().toString().trimmed() : "";

        // 3. 主键变化 → 创建新对象并记录索引
        if (currentId != lastId && !currentId.isEmpty()) {
            jsonArray.push_back(json::object());
            currentObjIndex = jsonArray.size() - 1; // 记录当前操作的对象索引
            lastId = currentId;
        }

        // 兜底：如果第一行就没有 ID，强制创建一个
        if (currentObjIndex == -1) {
            jsonArray.push_back(json::object());
            currentObjIndex = 0;
        }

        // 4. 处理当前行的所有列（通过索引安全访问）
        json& currentObj = jsonArray[currentObjIndex];
        for (int col = 1; col <= colCount; ++col) {
            auto *cell = xlsx.cellAt(currentRow, col);
            QString text = cell ? cell->value().toString() : "";

            if (col - 1 < secondline.size()) {
                handlelinetype(currentObj, col - 1, firstline[col-1], text);
            }
        }

        currentRow++;
    }
}

void MainWindow::handlelinetype(json& jsonObj, int lineindex, QString linename, QString text)
{
    if(text.isEmpty()){
        return;
    }

    int linetype = 1;//纯字段

    for (int i = 0; i < linename.size(); ++i) {

        //判断首次遇到
        //首次遇到点，展开为对象
        if (linename[i] == ".") {
            linetype = 2;

            break;
        }
        //首次遇到方框
        if (linename[i] == "[") {
            //结尾是方框，确认无嵌套
            if(i + 2 == linename.size()){
                linetype = 3;

                break;
            }
            else{
                //方框+点
                linetype = 4;

                break;
            }
        }
    }

    if (linetype == 1) {
        std::string keyStr = linename.toStdString();

        if (secondline[lineindex] == "int") {
            if (isIntegerString(text)) {
                jsonObj[keyStr] = text.toLongLong();
            } else {
                jsonObj[keyStr] = text.toDouble();
            }
        }
        else if (secondline[lineindex] == "str") {
            jsonObj[keyStr] = text.toStdString();
        }
        else if (secondline[lineindex] == "json") {
            try {
                jsonObj[keyStr] = json::parse(text.toStdString());
            } catch (...) {
                jsonObj[keyStr] = text.toStdString();
            }
        }
    }

    else if (linetype == 2) {
        QString reallinename = linename.split('.')[0];
        std::string keyStr = reallinename.toStdString();

        // contains
        if (jsonObj.contains(keyStr)) {
            if (jsonObj[keyStr].is_object()) {
                // 获取子对象
                json& lsjsonobj = jsonObj[keyStr];

                //截取剩下的拿去处理
                QString nextlinename = linename.mid(linename.indexOf('.') + 1);

                handlelinetype(lsjsonobj, lineindex, nextlinename, text);
            }
        }
        else{
            json lsjsonobj = json::object();

            //截取剩下的拿去处理
            QString nextlinename = linename.mid(linename.indexOf('.') + 1);
            handlelinetype(lsjsonobj, lineindex, nextlinename, text);

            jsonObj[keyStr] = lsjsonobj;
        }
    }

    else if (linetype == 3) {
        QStringList testlist = text.split(",");
        QString reallinename = linename.left(linename.length() - 2);
        std::string keyStr = reallinename.toStdString();

        json lsjsonarray = json::array();
        for (int i = 0; i < testlist.size(); ++i) {
            QString item = testlist[i].trimmed();
            if (item.isEmpty()) continue;

            if (secondline[lineindex] == "int") {
                if (isIntegerString(item)) {
                    lsjsonarray.push_back(item.toLongLong());
                } else {
                    lsjsonarray.push_back(item.toDouble());
                }
            }
            else if (secondline[lineindex] == "str") {
                lsjsonarray.push_back(item.toStdString());
            }
        }
        jsonObj[keyStr] = lsjsonarray;
    }

    else if (linetype == 4) {
        //获得数组键
        QString realname = linename.left(linename.indexOf("["));
        std::string realname_str = realname.toStdString();

        // 数组是否存在
        if (jsonObj.contains(realname_str) && jsonObj[realname_str].is_array()) {
            qDebug() << "[L4 ARR_EXIST]" << "array:" << realname << "size:" << jsonObj[realname_str].size();

            json& arr = jsonObj[realname_str];
            json& latestItem = arr[arr.size() - 1];

            // 获取剩余路径
            QString remaining = linename.mid(linename.indexOf("[]") + 2);
            if (remaining.startsWith(".")) remaining = remaining.mid(1);

            QString nextlinenamecheck = getnextlinename(linename);

            //是否需要新对象
            bool needNewObject = false;

            if (!remaining.isEmpty()) {
                QString fieldName = remaining;
                int dotPos = fieldName.indexOf('.');
                int bracketPos = fieldName.indexOf('[');
                int cutPos = -1;
                if (dotPos != -1) cutPos = dotPos;
                if (bracketPos != -1 && (cutPos == -1 || bracketPos < cutPos)) cutPos = bracketPos;

                QString rest = "";
                if (cutPos != -1) {
                    fieldName = fieldName.left(cutPos);
                    rest = remaining.mid(cutPos);
                    if (rest.startsWith(".")) rest = rest.mid(1);
                    if (rest.startsWith("[]")) {
                        fieldName += "[]";
                        rest = rest.mid(2);
                        if (rest.startsWith(".")) rest = rest.mid(1);
                    }
                }

                bool isLeaf = rest.isEmpty();

                if (isLeaf) {
                    if (latestItem.contains(fieldName.toStdString())) {
                        needNewObject = true;
                    }
                }
            }

            if (needNewObject) {
                json newObject = json::object();
                QString subPath = linename.mid(linename.indexOf("[]") + 3); // 跳过 [].
                handlelinetype(newObject, lineindex, subPath, text);
                arr.push_back(newObject);
            } else {
                json& targetItem = arr[arr.size() - 1];
                QString subPath = linename.mid(linename.indexOf("[]") + 3);
                handlelinetype(targetItem, lineindex, subPath, text);
            }

        } else {
            // 数组还不存在，创建数组并添加第一个对象
            json newObject = json::object();
            QString subPath = linename.mid(linename.indexOf("[]") + 3);
            handlelinetype(newObject, lineindex, subPath, text);

            json newArray = json::array();
            newArray.push_back(newObject);
            jsonObj[realname_str] = newArray;
        }
    }
}

QString MainWindow::getnextlinename(QString input)
{
    int pairIdx = input.indexOf("[]");
    if (pairIdx == -1) return QString();

    QString target = input.mid(pairIdx + 2);

    int dotPos = target.indexOf('.');
    int bracketPos = target.indexOf('[');

    int cutPos = -1;
    if (dotPos != -1) cutPos = dotPos;
    if (bracketPos != -1 && (cutPos == -1 || bracketPos < cutPos)) {
        cutPos = bracketPos;
    }

    return (cutPos == -1) ? target : target.left(cutPos);
}

bool MainWindow::isIntegerString(const QString &str)
{
    QString s = str.trimmed();
    if (s.isEmpty()) return false;
    int start = (s[0] == '-' || s[0] == '+') ? 1 : 0;
    for (int i = start; i < s.size(); ++i) {
        if (!s[i].isDigit()) return false;
    }
    return true;
}



void MainWindow::loadSheetNames(const QString &fileName)
{
    if (!ui->comboBox) return;

    Document xlsx(fileName);

    if (!xlsx.dimension().isValid()) {
        return;
    }

    QStringList sheetNames = xlsx.sheetNames();
    ui->comboBox->clear();
    ui->comboBox->addItems(sheetNames);
}


void MainWindow::on_pushButton_clicked()
{
    // 【修改点】使用 exe 所在目录作为初始目录
    QString defaultDir = QCoreApplication::applicationDirPath();

    QString fileName = QFileDialog::getOpenFileName(this,
                                                    tr("选择Excel文件"),
                                                    defaultDir,  // ← 这里改了
                                                    tr("Excel Files (*.xlsx)"));

    if (!fileName.isEmpty()) {
        if (ui->lineEdit_filePath) {
            ui->lineEdit_filePath->setText(fileName);
        }

        // 自动加载工作表名称
        loadSheetNames(fileName);
    }
}

void MainWindow::on_pushButton_2_clicked()
{
    QString fileName;
    if (ui->lineEdit_filePath) {
        fileName = ui->lineEdit_filePath->text();
    }

    if (fileName.isEmpty()) {
        QMessageBox::warning(this, tr("警告"), tr("请先选择Excel文件！"));
        return;
    }

    Document xlsx(fileName);

    if (!xlsx.dimension().isValid()) {
        QMessageBox::critical(this, tr("错误"), tr("无法加载Excel文件！"));
        return;
    }

    // 根据用户选择确定要解析的工作表
    QString sheetName;
    if (ui->comboBox && ui->comboBox->currentIndex() >= 0) {
        sheetName = ui->comboBox->currentText();
        if (!sheetName.isEmpty()) {
            if (!xlsx.selectSheet(sheetName)) {
                QMessageBox::critical(this, tr("错误"), tr("无法切换到指定工作表: %1").arg(sheetName));
                return;
            }
        }
    } else {
        QStringList sheetNames = xlsx.sheetNames();
        if (!sheetNames.isEmpty()) {
            sheetName = sheetNames[0];
            xlsx.selectSheet(sheetName);
        } else {
            QMessageBox::critical(this, tr("错误"), tr("Excel文件中没有可用的工作表！"));
            return;
        }
    }

    // 解析 Excel 数据
    parseExcel(xlsx);

    if (jsonArray.empty()) {
        QMessageBox::information(this, tr("提示"), tr("未解析到有效数据"));
        return;
    }

    std::string jsonString = jsonArray.dump(4);

    // 【修改点】保存对话框也使用 exe 所在目录
    QString defaultDir = QCoreApplication::applicationDirPath();

    // 生成默认文件名（只保留文件名部分，不带路径）
    QString defaultFileName = QFileInfo(fileName).completeBaseName() + ".json";

    // 拼接完整默认路径：exe目录 + 文件名
    QString defaultPath = defaultDir + "/" + defaultFileName;

    QString savePath = QFileDialog::getSaveFileName(this,
                                                    tr("保存JSON文件"),
                                                    defaultPath,  // ← 这里改了
                                                    tr("JSON Files (*.json)"));

    if (savePath.isEmpty()) {
        return;
    }

    QFile file(savePath);
    if (file.open(QIODevice::WriteOnly | QIODevice::Text)) {
        file.write(QByteArray::fromStdString(jsonString));
        file.close();
        QMessageBox::information(this, tr("成功"), tr("转换成功！\n%1").arg(savePath));
    } else {
        QMessageBox::critical(this, tr("错误"), tr("无法保存文件！"));
    }
}
