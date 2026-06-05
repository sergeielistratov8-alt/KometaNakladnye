<#
PowerShell script to build a Windows executable using PyInstaller and create a Desktop shortcut.
Run as Administrator or with execution policy allowing script execution.
#>
param(
    [string]$ProjectDir = (Get-Location).Path,
    [string]$VenvName = "venv",
    [string]$ExeName = "KometaNakladnye"
)

Set-Location $ProjectDir

# Create venv if missing
$venvPath = Join-Path $ProjectDir $VenvName
if (-not (Test-Path $venvPath)) {
    Write-Host "Creating virtual environment..."
    if (Get-Command python -ErrorAction SilentlyContinue) {
        python -m venv $venvPath
    } elseif (Get-Command py -ErrorAction SilentlyContinue) {
        py -3 -m venv $venvPath
    } else {
        Write-Error "Python не найден. Установите Python 3 и повторите попытку."
        exit 1
    }
}

$python = Join-Path $venvPath "Scripts\python.exe"
$pip = Join-Path $venvPath "Scripts\pip.exe"

if (-not (Test-Path $python)) {
    Write-Error "Python executable not found in venv. Попробуйте удалить venv и запустить скрипт снова."
    exit 1
}

Write-Host "Upgrading pip..."
& $python -m pip install --upgrade pip setuptools wheel

Write-Host "Installing requirements from requirements.txt..."
& $pip install -r requirements.txt --no-warn-script-location

Write-Host "Installing pyinstaller..."
& $pip install pyinstaller

# Build exe
Write-Host "Building executable with PyInstaller..."
# Use --windowed for GUI apps; onedir mode works better with Qt on Windows than onefile
& $python -m PyInstaller --noconfirm --onedir --windowed --name $ExeName main.py

$distExe = Join-Path $ProjectDir "dist\$ExeName\$ExeName.exe"
if (-not (Test-Path $distExe)) {
    Write-Error "Build failed, $distExe not found. Убедитесь, что PyInstaller успешно завершил работу."
    exit 2
}

# Create Desktop shortcut
$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop "$ExeName.lnk"
$wsh = New-Object -ComObject WScript.Shell
$sc = $wsh.CreateShortcut($shortcutPath)
$sc.TargetPath = $distExe
$sc.WorkingDirectory = (Join-Path $ProjectDir "dist\$ExeName")
$sc.IconLocation = $distExe
$sc.WindowStyle = 1
$sc.Save()

Write-Host "Build complete!"
Write-Host "Executable: $distExe"
Write-Host "Shortcut created on Desktop: $shortcutPath"
Write-Host ""
Write-Host "Примечание: папка dist\$ExeName содержит все зависимости. Не удаляйте её!"

exit 0
