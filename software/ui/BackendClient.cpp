#include "BackendClient.h"

#include <QCoreApplication>
#include <QDir>
#include <QFileInfo>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonParseError>
#include <QProcessEnvironment>
#include <QTemporaryFile>
#include <QUuid>

#ifdef Q_OS_WIN
#include <windows.h>
#endif

BackendClient::BackendClient(QObject *parent)
    : QObject(parent)
    , m_process(new QProcess(this))
{
    m_process->setProcessChannelMode(QProcess::SeparateChannels);
    connect(m_process, &QProcess::readyReadStandardOutput,
            this, &BackendClient::onReadyReadStandardOutput);
    connect(m_process, &QProcess::readyReadStandardError,
            this, &BackendClient::onReadyReadStandardError);
    connect(m_process, QOverload<int, QProcess::ExitStatus>::of(&QProcess::finished),
            this, &BackendClient::onProcessFinished);
    connect(m_process, &QProcess::errorOccurred,
            this, &BackendClient::onProcessError);
#ifdef Q_OS_WIN
    m_process->setCreateProcessArgumentsModifier(
        [](QProcess::CreateProcessArguments *args) {
            args->flags |= CREATE_NO_WINDOW;
        });
#endif
}

BackendClient::~BackendClient() = default;

bool BackendClient::isRunning() const
{
    return m_running;
}

bool BackendClient::prepareJobFile(const QStringList &inputs,
                                   const QString &outputDir,
                                   bool exportTc,
                                   bool exportEplan,
                                   bool skipInvalid,
                                   const QString &engine)
{
    if (!QDir().mkpath(outputDir)) {
        emit failed(QStringLiteral("无法创建输出目录：%1").arg(outputDir));
        return false;
    }

    m_jobFile = std::make_unique<QTemporaryFile>(
        QDir(outputDir).filePath(QStringLiteral("datatrans_job_XXXXXX.json")));
    m_jobFile->setAutoRemove(true);
    if (!m_jobFile->open()) {
        emit failed(QStringLiteral("无法创建临时任务文件：%1").arg(m_jobFile->errorString()));
        m_jobFile.reset();
        return false;
    }

    QJsonArray inputArray;
    for (const QString &input : inputs) {
        inputArray.append(input);
    }

    QJsonObject options;
    options.insert(QStringLiteral("export_tc"), exportTc);
    options.insert(QStringLiteral("export_eplan"), exportEplan);
    options.insert(QStringLiteral("on_invalid"), skipInvalid ? QStringLiteral("skip")
                                                             : QStringLiteral("abort"));
    options.insert(QStringLiteral("engine"),
                   engine.isEmpty() ? QStringLiteral("auto") : engine);

    QJsonObject payload;
    payload.insert(QStringLiteral("job_id"),
                   QUuid::createUuid().toString(QUuid::WithoutBraces).left(8));
    payload.insert(QStringLiteral("inputs"), inputArray);
    payload.insert(QStringLiteral("output_dir"), outputDir);
    payload.insert(QStringLiteral("options"), options);

    const QByteArray json = QJsonDocument(payload).toJson(QJsonDocument::Indented);
    if (m_jobFile->write(json) != json.size()) {
        emit failed(QStringLiteral("写入临时任务文件失败：%1").arg(m_jobFile->errorString()));
        resetJobFile();
        return false;
    }
    m_jobFile->flush();
    m_lastJobPath = m_jobFile->fileName();
    m_jobFile->close();
    m_outputDir = outputDir;
    return true;
}

QString BackendClient::resolveProgram(QStringList &arguments,
                                      QString &workingDirectory,
                                      QString &error) const
{
    const QString appDir = QCoreApplication::applicationDirPath();
    const QString core = qEnvironmentVariable("DATATRANS_CORE");
    const QString python = qEnvironmentVariable("DATATRANS_PYTHON");
    const QString backendDir = qEnvironmentVariable("DATATRANS_BACKEND_DIR");

    QString program;
    if (!core.isEmpty()) {
        const QFileInfo info(core);
        if (info.suffix().compare(QStringLiteral("py"), Qt::CaseInsensitive) == 0) {
            program = python.isEmpty() ? QStringLiteral("python") : python;
            arguments << core;
        } else {
            program = core;
        }
    } else if (!python.isEmpty()) {
        // 开发模式：使用 Python 模块运行，不要求先打 PyInstaller。
        program = python;
        arguments << QStringLiteral("-m") << QStringLiteral("datatrans");
        if (!backendDir.isEmpty()) {
            workingDirectory = backendDir;
        }
    } else {
        program = QDir(appDir).filePath(QStringLiteral("backend/datatrans-core.exe"));
    }

    if (program.isEmpty()) {
        error = QStringLiteral("未配置后端程序");
        return {};
    }

    const QFileInfo programInfo(program);
    if (!programInfo.exists() && programInfo.isAbsolute()) {
        error = QStringLiteral(
                    "找不到后端程序：%1\n\n"
                    "请先运行 packaging/build.ps1 打包；开发时可设置环境变量 "
                    "DATATRANS_PYTHON 和 DATATRANS_BACKEND_DIR。")
                    .arg(program);
        return {};
    }

    return program;
}

bool BackendClient::startJob(const QStringList &inputs,
                             const QString &outputDir,
                             bool exportTc,
                             bool exportEplan,
                             bool skipInvalid,
                             const QString &engine)
{
    if (m_running) {
        emit failed(QStringLiteral("任务正在运行，请等待完成或先取消。"));
        return false;
    }
    if (inputs.isEmpty()) {
        emit failed(QStringLiteral("未选择输入文件。"));
        return false;
    }

    resetJobFile();
    if (!prepareJobFile(inputs, outputDir, exportTc, exportEplan, skipInvalid, engine)) {
        return false;
    }

    QStringList arguments;
    QString workingDirectory;
    QString resolveError;
    const QString program = resolveProgram(arguments, workingDirectory, resolveError);
    if (program.isEmpty()) {
        emit failed(resolveError);
        resetJobFile();
        return false;
    }

    arguments << QStringLiteral("--job") << m_lastJobPath << QStringLiteral("--progress");

    QProcessEnvironment environment = QProcessEnvironment::systemEnvironment();
    environment.insert(QStringLiteral("PYTHONUTF8"), QStringLiteral("1"));
    environment.insert(QStringLiteral("PYTHONIOENCODING"), QStringLiteral("utf-8"));
    const QString backendDir = qEnvironmentVariable("DATATRANS_BACKEND_DIR");
    if (!backendDir.isEmpty()) {
        environment.insert(QStringLiteral("PYTHONPATH"), backendDir);
    }
    m_process->setProcessEnvironment(environment);

    if (!workingDirectory.isEmpty()) {
        m_process->setWorkingDirectory(workingDirectory);
    } else {
        m_process->setWorkingDirectory(QCoreApplication::applicationDirPath());
    }

    m_stdoutBuffer.clear();
    m_stderrBuffer.clear();
    m_resultSeen = false;
    m_cancelled = false;
    m_resultObject = QJsonObject();
    m_running = true;

    m_process->setProgram(program);
    m_process->setArguments(arguments);
    emit started();
    m_process->start();
    return true;
}

void BackendClient::cancel()
{
    if (!m_running) {
        return;
    }
    m_cancelled = true;
    emit logMessage(QStringLiteral("warning"), QStringLiteral("正在取消任务…"));
    m_process->kill();
}

void BackendClient::onReadyReadStandardOutput()
{
    processStdout();
}

void BackendClient::processStdout()
{
    m_stdoutBuffer += m_process->readAllStandardOutput();
    while (true) {
        const int newline = m_stdoutBuffer.indexOf('\n');
        if (newline < 0) {
            break;
        }
        QByteArray line = m_stdoutBuffer.left(newline);
        m_stdoutBuffer.remove(0, newline + 1);
        if (line.endsWith('\r')) {
            line.chop(1);
        }
        handleLine(QString::fromUtf8(line));
    }
}

void BackendClient::handleLine(const QString &line)
{
    const QString trimmed = line.trimmed();
    if (trimmed.isEmpty()) {
        return;
    }

    QJsonParseError parseError;
    const QJsonDocument document = QJsonDocument::fromJson(trimmed.toUtf8(), &parseError);
    if (parseError.error != QJsonParseError::NoError || !document.isObject()) {
        emit logMessage(QStringLiteral("debug"), trimmed);
        return;
    }

    const QJsonObject object = document.object();
    const QString type = object.value(QStringLiteral("type")).toString();
    if (type == QStringLiteral("progress")) {
        emit progress(object.value(QStringLiteral("current")).toInt(),
                      object.value(QStringLiteral("total")).toInt(),
                      object.value(QStringLiteral("stage")).toString(),
                      object.value(QStringLiteral("message")).toString());
    } else if (type == QStringLiteral("log")) {
        emit logMessage(object.value(QStringLiteral("level")).toString(),
                        object.value(QStringLiteral("message")).toString());
    } else if (type == QStringLiteral("result")) {
        handleResult(object);
    } else if (type == QStringLiteral("error")) {
        emit logMessage(QStringLiteral("error"),
                        object.value(QStringLiteral("message")).toString());
    } else {
        emit logMessage(QStringLiteral("debug"), trimmed);
    }
}

void BackendClient::handleResult(const QJsonObject &object)
{
    m_resultSeen = true;
    m_resultObject = object;

    const QJsonArray issues = object.value(QStringLiteral("issues")).toArray();
    for (const QJsonValue &value : issues) {
        const QJsonObject issue = value.toObject();
        const QString level = issue.value(QStringLiteral("level")).toString();
        QString message = issue.value(QStringLiteral("message")).toString();
        const QString file = issue.value(QStringLiteral("file")).toString();
        const int row = issue.value(QStringLiteral("row")).toInt();
        if (!file.isEmpty()) {
            QString location = QFileInfo(file).fileName();
            if (row > 0) {
                location += QStringLiteral(":%1").arg(row);
            }
            message += QStringLiteral("（%1）").arg(location);
        }
        emit logMessage(level.isEmpty() ? QStringLiteral("info") : level, message);
    }
}

void BackendClient::onReadyReadStandardError()
{
    m_stderrBuffer += m_process->readAllStandardError();
}

void BackendClient::onProcessFinished(int exitCode, QProcess::ExitStatus exitStatus)
{
    processStdout();

    if (!m_stdoutBuffer.trimmed().isEmpty()) {
        handleLine(QString::fromUtf8(m_stdoutBuffer));
        m_stdoutBuffer.clear();
    }

    m_running = false;

    if (m_cancelled) {
        emit finished(false, QStringLiteral("canceled"), QStringLiteral("已取消"), {}, {});
        resetJobFile();
        return;
    }

    if (m_resultSeen) {
        const bool ok = m_resultObject.value(QStringLiteral("ok")).toBool();
        const QString status = m_resultObject.value(QStringLiteral("status")).toString();
        const QString message = m_resultObject.value(QStringLiteral("message")).toString();
        QString resultPath = m_resultObject.value(QStringLiteral("result_path")).toString();
        if (resultPath.isEmpty() && !m_outputDir.isEmpty()) {
            resultPath = QDir(m_outputDir).filePath(QStringLiteral("result.json"));
        }

        QStringList outputs;
        const QJsonObject outputObject = m_resultObject.value(QStringLiteral("outputs")).toObject();
        for (const QString &key : {QStringLiteral("tc"), QStringLiteral("eplan")}) {
            const QString path = outputObject.value(key).toString();
            if (!path.isEmpty()) {
                outputs << path;
            }
        }
        emit finished(ok, status, message, resultPath, outputs);
        resetJobFile();
        return;
    }

    QString message;
    if (exitStatus == QProcess::CrashExit) {
        message = QStringLiteral("后端进程异常退出。");
    } else {
        message = QStringLiteral("后端进程退出，退出码 %1。").arg(exitCode);
    }
    const QString stderrText = QString::fromUtf8(m_stderrBuffer).trimmed();
    if (!stderrText.isEmpty()) {
        message += QStringLiteral("\n\n%1").arg(stderrText);
    }
    emit failed(message);
    resetJobFile();
}

void BackendClient::onProcessError(QProcess::ProcessError error)
{
    if (error == QProcess::FailedToStart) {
        const QString message = QStringLiteral("无法启动后端程序：%1\n%2")
                                    .arg(m_process->program(), m_process->errorString());
        m_running = false;
        emit failed(message);
        resetJobFile();
    }
}

void BackendClient::finishWithFailure(const QString &message)
{
    m_running = false;
    emit failed(message);
    resetJobFile();
}

void BackendClient::resetJobFile()
{
    m_jobFile.reset();
    m_lastJobPath.clear();
}
