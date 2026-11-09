#include "BackendClient.h"

#include <QCoreApplication>
#include <QDebug>
#include <QTimer>

int main(int argc, char *argv[])
{
    QCoreApplication application(argc, argv);
    if (argc < 3) {
        qCritical() << "usage: backend_client_smoke <input.xlsx> <output-dir>";
        return 64;
    }

    BackendClient client;
    QObject::connect(&client, &BackendClient::progress,
                     [](int current, int total, const QString &stage, const QString &message) {
                         qInfo().noquote() << QStringLiteral("progress %1/%2 [%3] %4")
                                                  .arg(current)
                                                  .arg(total)
                                                  .arg(stage, message);
                     });
    QObject::connect(&client, &BackendClient::logMessage,
                     [](const QString &level, const QString &message) {
                         qInfo().noquote() << QStringLiteral("[%1] %2").arg(level, message);
                     });
    QObject::connect(&client, &BackendClient::failed, [&application](const QString &message) {
        qCritical().noquote() << "FAILED:" << message;
        application.exit(2);
    });
    QObject::connect(&client, &BackendClient::finished,
                     [&application](bool ok, const QString &status, const QString &message,
                                    const QString &resultPath, const QStringList &outputs) {
                         qInfo().noquote() << "FINISHED:" << ok << status << message;
                         qInfo().noquote() << "RESULT:" << resultPath;
                         for (const QString &output : outputs) {
                             qInfo().noquote() << "OUTPUT:" << output;
                         }
                         application.exit(ok ? 0 : 1);
                     });

    QTimer::singleShot(0, [&client, argv]() {
        const bool started = client.startJob({QString::fromLocal8Bit(argv[1])},
                                             QString::fromLocal8Bit(argv[2]), true, true, false);
        if (!started) {
            qCritical() << "startJob returned false";
        }
    });
    QTimer::singleShot(120000, [&application]() {
        qCritical() << "TIMEOUT";
        application.exit(3);
    });

    return application.exec();
}
