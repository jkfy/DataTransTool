#pragma once

#include <QMainWindow>
#include <QSettings>
#include <QStringList>

class BackendClient;
class QCloseEvent;
class QDragEnterEvent;
class QDropEvent;

QT_BEGIN_NAMESPACE
namespace Ui {
class MainWindow;
}
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT
public:
    explicit MainWindow(QWidget *parent = nullptr);
    ~MainWindow() override;

protected:
    void dragEnterEvent(QDragEnterEvent *event) override;
    void dropEvent(QDropEvent *event) override;
    void closeEvent(QCloseEvent *event) override;

private slots:
    void addFiles();
    void removeSelected();
    void clearFiles();
    void browseOutput();
    void startProcessing();
    void cancelProcessing();
    void openOutputDir();
    void onProgress(int current, int total, const QString &stage, const QString &message);
    void onLog(const QString &level, const QString &message);
    void onFinished(bool ok,
                    const QString &status,
                    const QString &message,
                    const QString &resultPath,
                    const QStringList &outputs);
    void onFailed(const QString &message);

private:
    void setupConnections();
    void addFile(const QString &path);
    QStringList inputFiles() const;
    QString outputDirectory() const;
    QString defaultOutputDirectory() const;
    void appendLog(const QString &level, const QString &message);
    void setBusy(bool busy);
    void showResult(bool ok,
                    const QString &status,
                    const QString &message,
                    const QStringList &outputs);

    Ui::MainWindow *ui = nullptr;
    BackendClient *m_backend = nullptr;
    QSettings m_settings;
    bool m_busy = false;
};
