#include "MainWindow.h"

#include "BackendClient.h"
#include "ui_MainWindow.h"

#include <QCloseEvent>
#include <QDesktopServices>
#include <QDir>
#include <QDragEnterEvent>
#include <QDropEvent>
#include <QFileDialog>
#include <QFileInfo>
#include <QGuiApplication>
#include <QListWidget>
#include <QMessageBox>
#include <QMimeData>
#include <QPushButton>
#include <QScreen>
#include <QStandardPaths>
#include <QStatusBar>
#include <QTime>
#include <QUrl>
#include <QWidget>

namespace {
bool quietMode()
{
    return qEnvironmentVariableIsSet("DATATRANS_NO_DIALOGS");
}
} // namespace

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
    , m_backend(new BackendClient(this))
    , m_settings(QStringLiteral("DataTransTool"), QStringLiteral("DataTransTool"))
{
    ui->setupUi(this);
    ui->logEdit->setMaximumBlockCount(5000);
    setupConnections();

    ui->tcCheck->setChecked(m_settings.value(QStringLiteral("exportTc"), true).toBool());
    ui->eplanCheck->setChecked(m_settings.value(QStringLiteral("exportEplan"), true).toBool());
    ui->skipInvalidCheck->setChecked(
        m_settings.value(QStringLiteral("skipInvalid"), false).toBool());

    const QString lastOutput =
        m_settings.value(QStringLiteral("lastOutputDir"), defaultOutputDirectory()).toString();
    ui->outputEdit->setText(lastOutput);

    const QByteArray savedGeometry = m_settings.value(QStringLiteral("geometry")).toByteArray();
    if (savedGeometry.isEmpty()) {
        resize(960, 780);
    } else {
        restoreGeometry(savedGeometry);
    }

    // 多显示器拔插后，旧 geometry 可能落在屏幕外；兜底移到主屏中央。
    if (QScreen *screen = QGuiApplication::primaryScreen()) {
        const QRect available = screen->availableGeometry();
        if (savedGeometry.isEmpty() || !available.intersects(geometry())) {
            const int width = qMax(800, qMin(960, available.width() - 40));
            const int height = qMax(620, qMin(780, available.height() - 40));
            resize(width, height);
            move(available.center() - rect().center());
        }
    }

    statusBar()->showMessage(QStringLiteral("就绪"));
    appendLog(QStringLiteral("INFO"),
              QStringLiteral("就绪。添加一个或多个 .xlsx 物料总表，选择输出目录后点击“开始处理”。"));
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::setupConnections()
{
    connect(ui->addButton, &QPushButton::clicked, this, &MainWindow::addFiles);
    connect(ui->removeButton, &QPushButton::clicked, this, &MainWindow::removeSelected);
    connect(ui->clearButton, &QPushButton::clicked, this, &MainWindow::clearFiles);
    connect(ui->browseButton, &QPushButton::clicked, this, &MainWindow::browseOutput);
    connect(ui->startButton, &QPushButton::clicked, this, &MainWindow::startProcessing);
    connect(ui->cancelButton, &QPushButton::clicked, this, &MainWindow::cancelProcessing);
    connect(ui->openButton, &QPushButton::clicked, this, &MainWindow::openOutputDir);

    connect(m_backend, &BackendClient::progress, this, &MainWindow::onProgress);
    connect(m_backend, &BackendClient::logMessage, this, &MainWindow::onLog);
    connect(m_backend, &BackendClient::finished, this, &MainWindow::onFinished);
    connect(m_backend, &BackendClient::failed, this, &MainWindow::onFailed);
}

void MainWindow::addFiles()
{
    const QString lastDir =
        m_settings.value(QStringLiteral("lastInputDir"),
                         QStandardPaths::writableLocation(QStandardPaths::DocumentsLocation))
            .toString();
    const QStringList files = QFileDialog::getOpenFileNames(
        this, QStringLiteral("选择物料总表（.xlsx）"), lastDir,
        QStringLiteral("Excel 工作簿 (*.xlsx)"));
    if (files.isEmpty()) {
        return;
    }
    m_settings.setValue(QStringLiteral("lastInputDir"), QFileInfo(files.first()).absolutePath());
    for (const QString &file : files) {
        addFile(file);
    }
}

void MainWindow::addFile(const QString &path)
{
    const QFileInfo info(path);
    if (!info.exists() || !info.isFile()) {
        QMessageBox::warning(this, QStringLiteral("文件不存在"),
                             QStringLiteral("找不到文件：\n%1").arg(path));
        return;
    }
    if (info.suffix().compare(QStringLiteral("xlsx"), Qt::CaseInsensitive) != 0) {
        QMessageBox::warning(
            this, QStringLiteral("格式不支持"),
            QStringLiteral("只支持 .xlsx 格式：\n%1").arg(info.fileName()));
        return;
    }

    const QString absolutePath = info.absoluteFilePath();
    for (int row = 0; row < ui->fileList->count(); ++row) {
        if (ui->fileList->item(row)->data(Qt::UserRole).toString() == absolutePath) {
            return;
        }
    }

    auto *item = new QListWidgetItem(info.fileName(), ui->fileList);
    item->setToolTip(absolutePath);
    item->setData(Qt::UserRole, absolutePath);
}

void MainWindow::removeSelected()
{
    const QList<QListWidgetItem *> selected = ui->fileList->selectedItems();
    for (QListWidgetItem *item : selected) {
        delete item;
    }
}

void MainWindow::clearFiles()
{
    ui->fileList->clear();
}

void MainWindow::browseOutput()
{
    const QString directory = QFileDialog::getExistingDirectory(
        this, QStringLiteral("选择输出目录"), outputDirectory());
    if (!directory.isEmpty()) {
        ui->outputEdit->setText(QDir::toNativeSeparators(directory));
    }
}

void MainWindow::startProcessing()
{
    if (m_busy) {
        return;
    }

    const QStringList files = inputFiles();
    if (files.isEmpty()) {
        if (!quietMode()) {
            QMessageBox::warning(this, QStringLiteral("没有输入文件"),
                                 QStringLiteral("请先添加至少一个 .xlsx 物料总表。"));
        }
        return;
    }

    if (!ui->tcCheck->isChecked() && !ui->eplanCheck->isChecked()) {
        if (!quietMode()) {
            QMessageBox::warning(this, QStringLiteral("没有输出"),
                                 QStringLiteral("请至少勾选一种输出表（TC 或 Eplan）。"));
        }
        return;
    }

    const QString directory = outputDirectory();
    if (directory.isEmpty()) {
        if (!quietMode()) {
            QMessageBox::warning(this, QStringLiteral("没有输出目录"),
                                 QStringLiteral("请选择输出目录。"));
        }
        return;
    }
    if (!QDir().mkpath(directory)) {
        if (!quietMode()) {
            QMessageBox::critical(this, QStringLiteral("输出目录不可用"),
                                  QStringLiteral("无法创建或写入输出目录：\n%1").arg(directory));
        }
        return;
    }

    m_settings.setValue(QStringLiteral("lastOutputDir"), directory);
    m_settings.setValue(QStringLiteral("exportTc"), ui->tcCheck->isChecked());
    m_settings.setValue(QStringLiteral("exportEplan"), ui->eplanCheck->isChecked());
    m_settings.setValue(QStringLiteral("skipInvalid"), ui->skipInvalidCheck->isChecked());

    ui->logEdit->clear();
    appendLog(QStringLiteral("INFO"),
              QStringLiteral("开始处理 %1 个文件，输出目录：%2")
                  .arg(files.size())
                  .arg(directory));
    setBusy(true);
    ui->progressBar->setRange(0, 0);
    ui->statusLabel->setText(QStringLiteral("正在启动后端…"));
    statusBar()->showMessage(QStringLiteral("正在启动后端…"));

    if (!m_backend->startJob(files, directory, ui->tcCheck->isChecked(),
                             ui->eplanCheck->isChecked(), ui->skipInvalidCheck->isChecked())) {
        // startJob 失败时已经通过 failed 信号给出提示，这里只恢复界面状态。
        if (!m_backend->isRunning()) {
            setBusy(false);
            ui->progressBar->setRange(0, 100);
            ui->progressBar->setValue(0);
        }
    }
}

void MainWindow::cancelProcessing()
{
    if (!m_backend->isRunning()) {
        return;
    }
    const auto answer = QMessageBox::question(
        this, QStringLiteral("取消任务"),
        QStringLiteral("确定取消当前处理吗？已写出的临时/输出文件可能会保留。"),
        QMessageBox::Yes | QMessageBox::No, QMessageBox::No);
    if (answer == QMessageBox::Yes) {
        ui->cancelButton->setEnabled(false);
        ui->statusLabel->setText(QStringLiteral("正在取消…"));
        m_backend->cancel();
    }
}

void MainWindow::openOutputDir()
{
    const QString directory = outputDirectory();
    if (directory.isEmpty() || !QDir(directory).exists()) {
        QMessageBox::information(this, QStringLiteral("输出目录"),
                                 QStringLiteral("输出目录还不存在。"));
        return;
    }
    QDesktopServices::openUrl(QUrl::fromLocalFile(directory));
}

void MainWindow::onProgress(int current, int total, const QString &stage,
                            const QString &message)
{
    Q_UNUSED(stage)
    if (total > 0) {
        ui->progressBar->setRange(0, total);
        ui->progressBar->setValue(current);
    } else {
        ui->progressBar->setRange(0, 0);
    }
    if (!message.isEmpty()) {
        ui->statusLabel->setText(message);
        statusBar()->showMessage(message);
        appendLog(QStringLiteral("INFO"), message);
    }
}

void MainWindow::onLog(const QString &level, const QString &message)
{
    appendLog(level.isEmpty() ? QStringLiteral("INFO") : level.toUpper(), message);
}

void MainWindow::onFinished(bool ok, const QString &status, const QString &message,
                            const QString &resultPath, const QStringList &outputs)
{
    setBusy(false);

    if (status == QStringLiteral("canceled")) {
        ui->progressBar->setRange(0, 100);
        ui->progressBar->setValue(0);
        ui->statusLabel->setText(QStringLiteral("已取消"));
        statusBar()->showMessage(QStringLiteral("已取消"));
        appendLog(QStringLiteral("WARNING"), QStringLiteral("任务已取消。"));
        return;
    }

    ui->progressBar->setRange(0, 100);
    ui->progressBar->setValue(ok ? 100 : 0);
    ui->statusLabel->setText(ok ? QStringLiteral("完成") : QStringLiteral("失败"));
    statusBar()->showMessage(message);

    appendLog(ok ? QStringLiteral("INFO") : QStringLiteral("ERROR"), message);
    for (const QString &path : outputs) {
        appendLog(QStringLiteral("INFO"), QStringLiteral("输出文件：%1").arg(path));
    }
    if (!resultPath.isEmpty()) {
        appendLog(QStringLiteral("INFO"), QStringLiteral("结果文件：%1").arg(resultPath));
    }

    ui->openButton->setEnabled(!outputDirectory().isEmpty());
    showResult(ok, status, message, outputs);
}

void MainWindow::onFailed(const QString &message)
{
    setBusy(false);
    ui->progressBar->setRange(0, 100);
    ui->progressBar->setValue(0);
    ui->statusLabel->setText(QStringLiteral("失败"));
    statusBar()->showMessage(QStringLiteral("处理失败"));
    appendLog(QStringLiteral("ERROR"), message);
    if (!quietMode()) {
        QMessageBox::critical(this, QStringLiteral("处理失败"), message);
    }
}

void MainWindow::showResult(bool ok, const QString &status, const QString &message,
                            const QStringList &outputs)
{
    QString text = message;
    if (!outputs.isEmpty()) {
        text += QStringLiteral("\n\n输出文件：\n") + outputs.join(QStringLiteral("\n"));
    } else {
        text += QStringLiteral("\n\n未生成输出文件。");
    }
    if (!outputDirectory().isEmpty()) {
        text += QStringLiteral("\n\n输出目录：\n") + outputDirectory();
    }

    if (quietMode()) {
        return;
    }
    if (!ok) {
        QMessageBox::critical(this, QStringLiteral("处理失败"), text);
    } else if (status == QStringLiteral("partial")) {
        QMessageBox::warning(this, QStringLiteral("处理完成（有警告）"), text);
    } else {
        QMessageBox::information(this, QStringLiteral("处理完成"), text);
    }
}

void MainWindow::setBusy(bool busy)
{
    m_busy = busy;
    for (QWidget *widget : {static_cast<QWidget *>(ui->addButton),
                            static_cast<QWidget *>(ui->removeButton),
                            static_cast<QWidget *>(ui->clearButton),
                            static_cast<QWidget *>(ui->browseButton),
                            static_cast<QWidget *>(ui->startButton),
                            static_cast<QWidget *>(ui->outputEdit),
                            static_cast<QWidget *>(ui->fileList),
                            static_cast<QWidget *>(ui->tcCheck),
                            static_cast<QWidget *>(ui->eplanCheck),
                            static_cast<QWidget *>(ui->skipInvalidCheck)}) {
        widget->setEnabled(!busy);
    }
    ui->cancelButton->setEnabled(busy);
    ui->openButton->setEnabled(!busy && !outputDirectory().isEmpty());
}

void MainWindow::appendLog(const QString &level, const QString &message)
{
    if (!ui || !ui->logEdit) {
        return;
    }
    const QString line = QStringLiteral("%1  [%2]  %3")
                             .arg(QTime::currentTime().toString(QStringLiteral("HH:mm:ss")),
                                  level.toUpper(), message);
    ui->logEdit->appendPlainText(line);
}

QStringList MainWindow::inputFiles() const
{
    QStringList files;
    for (int row = 0; row < ui->fileList->count(); ++row) {
        const QString path = ui->fileList->item(row)->data(Qt::UserRole).toString();
        if (!path.isEmpty()) {
            files << path;
        }
    }
    return files;
}

QString MainWindow::outputDirectory() const
{
    return ui && ui->outputEdit ? ui->outputEdit->text().trimmed() : QString();
}

QString MainWindow::defaultOutputDirectory() const
{
    QString documents = QStandardPaths::writableLocation(QStandardPaths::DocumentsLocation);
    if (documents.isEmpty()) {
        documents = QDir::homePath();
    }
    return QDir(documents).filePath(QStringLiteral("DataTransTool输出"));
}

void MainWindow::dragEnterEvent(QDragEnterEvent *event)
{
    if (!event->mimeData()->hasUrls()) {
        return;
    }
    for (const QUrl &url : event->mimeData()->urls()) {
        if (url.isLocalFile()
            && QFileInfo(url.toLocalFile()).suffix().compare(QStringLiteral("xlsx"),
                                                             Qt::CaseInsensitive)
                   == 0) {
            event->acceptProposedAction();
            return;
        }
    }
}

void MainWindow::dropEvent(QDropEvent *event)
{
    bool accepted = false;
    for (const QUrl &url : event->mimeData()->urls()) {
        if (!url.isLocalFile()) {
            continue;
        }
        const QString path = url.toLocalFile();
        if (QFileInfo(path).suffix().compare(QStringLiteral("xlsx"), Qt::CaseInsensitive) == 0) {
            addFile(path);
            accepted = true;
        }
    }
    if (accepted) {
        event->acceptProposedAction();
        statusBar()->showMessage(QStringLiteral("已添加文件"), 3000);
    }
}

void MainWindow::closeEvent(QCloseEvent *event)
{
    if (m_backend->isRunning()) {
        if (!quietMode()) {
            const auto answer = QMessageBox::question(
                this, QStringLiteral("任务正在运行"),
                QStringLiteral("当前任务仍在运行，退出会取消它。确定退出吗？"),
                QMessageBox::Yes | QMessageBox::No, QMessageBox::No);
            if (answer != QMessageBox::Yes) {
                event->ignore();
                return;
            }
        }
        m_backend->cancel();
    }
    m_settings.setValue(QStringLiteral("geometry"), saveGeometry());
    event->accept();
}
