<#
Сборка Windows-исполняемого файла через PyInstaller.
#>
param(
    [string]$ProjectDir = (Get-Location).Path,
    [string]$VenvName = "venv",
    [string]$ExeName = "KometaNakladnye"
)

Set-Location $ProjectDir

$venvPath = Join-Path $ProjectDir $VenvName
$pythonExe = Join-Path $venvPath "Scripts\python.exe"

if (-not (Test-Path $pythonExe)) {
    if (Test-Path $venvPath) {
        Write-Host "Removing broken venv..."
        Remove-Item -Recurse -Force $venvPath
    }
    Write-Host "Creating virtual environment..."
    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -3 -m venv $venvPath
    } elseif (Get-Command python -ErrorAction SilentlyContinue) {
        & python -m venv $venvPath
    } else {
        Write-Error "Python не найден. Установите Python 3 и повторите попытку."
        exit 1
    }
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $pythonExe)) {
        if (Test-Path $venvPath) { Remove-Item -Recurse -Force $venvPath }
        Write-Error "Не удалось создать venv. Используйте команду: py -3 -m venv venv"
        exit 1
    }
}

$python = $pythonExe

& $python -m pip --version 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Restoring pip in venv..."
    $sitePackages = Join-Path $venvPath "Lib\site-packages"
    Get-ChildItem $sitePackages -Filter "~ip*" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
    & $python -m ensurepip --upgrade
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Recreating venv..."
        Remove-Item -Recurse -Force $venvPath
        & py -3 -m venv $venvPath
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $pythonExe)) {
            Write-Error "Не удалось пересоздать venv."
            exit 1
        }
    }
}

Write-Host "Installing dependencies..."
& $python -m pip install PyQt5 pandas openpyxl xlrd xlwt --default-timeout=100
if ($LASTEXITCODE -ne 0) {
    Write-Error "Не удалось установить зависимости. Проверьте интернет и повторите."
    exit 1
}

Write-Host "Verifying dependencies..."
& $python -c "import PyQt5, pandas, openpyxl, xlrd, xlwt; print('OK:', pandas.__version__)"
if ($LASTEXITCODE -ne 0) {
    Write-Error "Зависимости не импортируются. Удалите venv и запустите скрипт снова."
    exit 1
}

Write-Host "Installing PyInstaller..."
& $python -m pip install pyinstaller
if ($LASTEXITCODE -ne 0) { Write-Error "Не удалось установить PyInstaller."; exit 1 }

Write-Host "Building executable (single portable .exe)..."
& $python -m PyInstaller --noconfirm --onefile --windowed --name $ExeName `
    --hidden-import PyQt5 --hidden-import pandas --hidden-import openpyxl --hidden-import xlrd --hidden-import xlwt `
    --collect-all pandas --collect-all openpyxl --collect-all numpy --collect-all xlrd --collect-all xlwt `
    main.py
if ($LASTEXITCODE -ne 0) { Write-Error "PyInstaller завершился с ошибкой."; exit 1 }

$distExe = Join-Path $ProjectDir "dist\$ExeName.exe"
if (-not (Test-Path $distExe)) {
    Write-Error "Build failed, $distExe not found."
    exit 2
}

$readyDir = Join-Path $ProjectDir "ready"
New-Item -ItemType Directory -Force -Path $readyDir | Out-Null
$portableExe = Join-Path $readyDir "$ExeName.exe"
Copy-Item -Force $distExe $portableExe

$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop "$ExeName.lnk"
$wsh = New-Object -ComObject WScript.Shell
$sc = $wsh.CreateShortcut($shortcutPath)
$sc.TargetPath = $portableExe
$sc.WorkingDirectory = $readyDir
$sc.IconLocation = $portableExe
$sc.WindowStyle = 1
$sc.Save()

Write-Host "Build complete!"
Write-Host "Portable exe: $portableExe"
Write-Host 'Copy this single file anywhere - no _internal folder needed.'
Write-Host "Shortcut: $shortcutPath"

exit 0
