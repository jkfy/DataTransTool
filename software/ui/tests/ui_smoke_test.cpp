#include "MainWindow.h"

#include <QApplication>
#include <QCheckBox>
#include <QDir>
#include <QFileInfo>
#include <QLineEdit>
#include <QListWidget>
#include <QPlainTextEdit>
#include <QPushButton>
#include <QSettings>
#include <QTemporaryDir>
#include <QtTest>

class UiSmokeTest : public QObject
{
    Q_OBJECT
private slots:
    void initTestCase();
    void processSample();
};

void UiSmokeTest::initTestCase()
{
    qputenv("DATATRANS_NO_DIALOGS", "1");
}

void UiSmokeTest::processSample()
{
    const QString sample = qEnvironmentVariable("DATATRANS_SAMPLE");
    QVERIFY2(!sample.isEmpty(), "DATATRANS_SAMPLE is not set");
    QVERIFY2(QFileInfo::exists(sample), qPrintable(sample));

    QTemporaryDir output;
    QVERIFY(output.isValid());

    QSettings settings(QStringLiteral("DataTransTool"), QStringLiteral("DataTransTool"));
    settings.clear();
    settings.sync();

    MainWindow window;
    window.setGeometry(100, 100, 960, 780);
    window.show();
    QTest::qWait(300);

    auto *fileList = window.findChild<QListWidget *>();
    QVERIFY(fileList);
    auto *outputEdit = window.findChild<QLineEdit *>();
    QVERIFY(outputEdit);
    auto *log = window.findChild<QPlainTextEdit *>();
    QVERIFY(log);

    outputEdit->setText(output.path());

    auto *item = new QListWidgetItem(QFileInfo(sample).fileName(), fileList);
    item->setData(Qt::UserRole, sample);

    QPushButton *startButton = nullptr;
    for (QPushButton *button : window.findChildren<QPushButton *>()) {
        if (button->text().contains(QStringLiteral("开始处理"))) {
            startButton = button;
            break;
        }
    }
    QVERIFY(startButton);

    startButton->click();

    QTRY_VERIFY_WITH_TIMEOUT(
        log->toPlainText().contains(QStringLiteral("结果文件："))
            || log->toPlainText().contains(QStringLiteral("处理失败")),
        30000);
    QVERIFY2(log->toPlainText().contains(QStringLiteral("结果文件：")),
             qPrintable(log->toPlainText()));

    const QStringList files = QDir(output.path()).entryList(
        {QStringLiteral("TC_*.xlsx"), QStringLiteral("EPlan_*.xlsx")}, QDir::Files);
    QCOMPARE(files.size(), 2);

    window.close();
}

QTEST_MAIN(UiSmokeTest)
#include "ui_smoke_test.moc"
