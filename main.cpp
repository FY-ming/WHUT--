// main.cpp文件
// 是系统的入口，用于启动系统。
// 初始化并启动SystemWindow
#include "systemwindow.h"
#include <QApplication>

using namespace std;

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    SystemWindow w;
    w.show();
    return a.exec();
}




