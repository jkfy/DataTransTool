#requires -Version 5.1
<#
.SYNOPSIS
  构建 DataTransTool：Python 后端（PyInstaller）+ Qt/C++ 前端（CMake/MinGW），
  然后用 windeployqt 收集 Qt 运行库，最后打成一个可离线分发的 zip。

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File packaging\build.ps1
  powershell -ExecutionPolicy Bypass -File packaging\build.ps1 -QtDir "D:\Qt\6.10.0\mingw_64"
#>
param(
    [string]$QtDir = "D:\QtCreate\6.10.0\mingw_64",
    [string]$MinGWDir = "D:\QtCreate\Tools\mingw1310_64",
    [string]$Python = "python",
    [string]$Configuration = "Release",
    [switch]$SkipBackend,
    [switch]$SkipUi,
    [switch]$SkipZip
)

$ErrorActionPreference = "Stop"
$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$backendDir = Join-Path $root "backend"
$uiDir = Join-Path $root "ui"
$templatesDir = Join-Path $root "templates"
$buildDir = Join-Path $root "build"
$distDir = Join-Path $root "dist"
$venvDir = Join-Path $buildDir "venv"
$backendDist = Join-Path $distDir "pyinstaller"
$appDir = Join-Path $distDir "DataTransTool"
$version = "0.1.0"

function Write-Step([string]$Message) {
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Resolve-RequiredPath([string]$Path, [string]$Name) {
    if (-not (Test-Path $Path)) {
        throw "$Name 不存在：$Path"
    }
    return (Resolve-Path $Path).Path
}

$QtDir = Resolve-RequiredPath $QtDir "Qt 目录"
$MinGWDir = Resolve-RequiredPath $MinGWDir "MinGW 目录"
$templatesDir = Resolve-RequiredPath $templatesDir "模板目录"

$env:Path = "$QtDir\bin;$MinGWDir\bin;" + $env:Path
$windeployqt = Join-Path $QtDir "bin\windeployqt.exe"
$cmake = (Get-Command cmake -ErrorAction Stop).Source
$gcc = Join-Path $MinGWDir "bin\gcc.exe"
$gxx = Join-Path $MinGWDir "bin\g++.exe"

New-Item -ItemType Directory -Path $buildDir -Force | Out-Null
New-Item -ItemType Directory -Path $distDir -Force | Out-Null

if (-not $SkipBackend) {
    Write-Step "准备 Python 构建环境"
    if (-not (Test-Path (Join-Path $venvDir "Scripts\python.exe"))) {
        & $Python -m venv $venvDir
    }
    $venvPython = Join-Path $venvDir "Scripts\python.exe"
    & $venvPython -m pip install --disable-pip-version-check --upgrade pip
    & $venvPython -m pip install --disable-pip-version-check -r (Join-Path $backendDir "requirements-build.txt")

    Write-Step "使用 PyInstaller 打包 Python 后端"
    $pyinstallerDist = Join-Path $backendDist "datatrans-core"
    if (Test-Path $pyinstallerDist) {
        Remove-Item -Recurse -Force $pyinstallerDist
    }
    & $venvPython -m PyInstaller `
        --noconfirm `
        --clean `
        --onedir `
        --name "datatrans-core" `
        --paths $backendDir `
        --add-data "$templatesDir;templates" `
        --distpath $backendDist `
        --workpath (Join-Path $buildDir "pyinstaller") `
        --specpath $buildDir `
        (Join-Path $backendDir "datatrans_launcher.py")
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller 失败，退出码 $LASTEXITCODE"
    }
}

if (-not $SkipUi) {
    Write-Step "配置并构建 Qt/C++ 前端"
    $uiBuild = Join-Path $buildDir "ui-release"
    & $cmake -S $uiDir -B $uiBuild -G "MinGW Makefiles" `
        "-DCMAKE_BUILD_TYPE=$Configuration" `
        "-DCMAKE_PREFIX_PATH=$QtDir" `
        "-DCMAKE_C_COMPILER=$gcc" `
        "-DCMAKE_CXX_COMPILER=$gxx" `
        "-DDATATRANS_BUILD_SMOKE_TEST=OFF" `
        "-DDATATRANS_BUILD_UI_TEST=OFF"
    if ($LASTEXITCODE -ne 0) {
        throw "CMake 配置失败，退出码 $LASTEXITCODE"
    }
    & $cmake --build $uiBuild -j 8
    if ($LASTEXITCODE -ne 0) {
        throw "Qt 构建失败，退出码 $LASTEXITCODE"
    }
}

Write-Step "组装离线目录"
if (Test-Path $appDir) {
    Remove-Item -Recurse -Force $appDir
}
New-Item -ItemType Directory -Path $appDir -Force | Out-Null

$uiExe = Join-Path $buildDir "ui-release\DataTransTool.exe"
if (-not (Test-Path $uiExe)) {
    throw "找不到前端产物：$uiExe"
}
Copy-Item $uiExe (Join-Path $appDir "DataTransTool.exe")

& $windeployqt --release --no-translations --compiler-runtime (Join-Path $appDir "DataTransTool.exe")
if ($LASTEXITCODE -ne 0) {
    throw "windeployqt 失败，退出码 $LASTEXITCODE"
}

$appBackend = Join-Path $appDir "backend"
New-Item -ItemType Directory -Path $appBackend -Force | Out-Null
$backendSource = Join-Path $backendDist "datatrans-core"
if (-not (Test-Path $backendSource)) {
    throw "找不到后端产物：$backendSource"
}
Copy-Item (Join-Path $backendSource "*") $appBackend -Recurse -Force

$appTemplates = Join-Path $appDir "templates"
Copy-Item $templatesDir $appTemplates -Recurse -Force
Copy-Item (Join-Path $PSScriptRoot "README_offline.txt") (Join-Path $appDir "README.txt") -Force
Copy-Item (Join-Path $PSScriptRoot "THIRD_PARTY_NOTICES.txt") (Join-Path $appDir "THIRD_PARTY_NOTICES.txt") -Force

# 部分 MinGW 运行库 windeployqt 可能漏拷，兜底再复制一次。
foreach ($dll in @("libgcc_s_seh-1.dll", "libstdc++-6.dll", "libwinpthread-1.dll")) {
    $source = Join-Path $MinGWDir "bin\$dll"
    if ((Test-Path $source) -and -not (Test-Path (Join-Path $appDir $dll))) {
        Copy-Item $source $appDir
    }
}

if (-not $SkipZip) {
    Write-Step "生成离线压缩包"
    $zipPath = Join-Path $distDir "DataTransTool-$version-win64.zip"
    if (Test-Path $zipPath) {
        Remove-Item -Force $zipPath
    }
    Compress-Archive -Path (Join-Path $appDir "*") -DestinationPath $zipPath -CompressionLevel Optimal
    Write-Host ""
    Write-Host "完成：" -ForegroundColor Green
    Write-Host "  目录：$appDir"
    Write-Host "  压缩包：$zipPath"
    $size = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
    Write-Host "  大小：$size MB"
} else {
    Write-Host ""
    Write-Host "完成（未压缩）：$appDir" -ForegroundColor Green
}
