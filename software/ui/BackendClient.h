#pragma once

#include <QByteArray>
#include <QJsonObject>
#include <QObject>
#include <QProcess>
#include <QString>
#include <QStringList>

#include <memory>

class QTemporaryFile;

/// 负责启动 Python 后端 CLI，并通过 JSONL 协议接收进度/结果。
class BackendClient : public QObject
{
    Q_OBJECT
public:
    explicit BackendClient(QObject *parent = nullptr);
    ~BackendClient() override;

    bool isRunning() const;
    bool startJob(const QStringList &inputs,
                  const QString &outputDir,
                  bool exportTc,
                  bool exportEplan,
                  bool skipInvalid);
    void cancel();

    QString lastJobPath() const { return m_lastJobPath; }

signals:
    void started();
    void progress(int current, int total, const QString &stage, const QString &message);
    void logMessage(const QString &level, const QString &message);
    void finished(bool ok,
                  const QString &status,
                  const QString &message,
                  const QString &resultPath,
                  const QStringList &outputs);
    void failed(const QString &message);

private slots:
    void onReadyReadStandardOutput();
    void onReadyReadStandardError();
    void onProcessFinished(int exitCode, QProcess::ExitStatus exitStatus);
    void onProcessError(QProcess::ProcessError error);

private:
    bool prepareJobFile(const QStringList &inputs,
                        const QString &outputDir,
                        bool exportTc,
                        bool exportEplan,
                        bool skipInvalid);
    QString resolveProgram(QStringList &arguments, QString &workingDirectory, QString &error) const;
    void processStdout();
    void handleLine(const QString &line);
    void handleResult(const QJsonObject &object);
    void finishWithFailure(const QString &message);
    void resetJobFile();

    QProcess *m_process = nullptr;
    std::unique_ptr<QTemporaryFile> m_jobFile;
    QByteArray m_stdoutBuffer;
    QByteArray m_stderrBuffer;
    bool m_running = false;
    bool m_cancelled = false;
    bool m_resultSeen = false;
    QJsonObject m_resultObject;
    QString m_lastJobPath;
    QString m_outputDir;
};
