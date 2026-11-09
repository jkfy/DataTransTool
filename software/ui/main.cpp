#include "MainWindow.h"

#include <QApplication>
#include <QFont>

int main(int argc, char *argv[])
{
    QApplication application(argc, argv);
    application.setApplicationName(QStringLiteral("DataTransTool"));
    application.setApplicationVersion(QStringLiteral("0.1.0"));
    application.setOrganizationName(QStringLiteral("DataTransTool"));

    QFont font(QStringLiteral("Microsoft YaHei UI"));
    font.setPointSize(9);
    application.setFont(font);

    MainWindow window;
    window.show();
    return application.exec();
}
