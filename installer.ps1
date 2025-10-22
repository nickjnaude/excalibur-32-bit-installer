<#
.SYNOPSIS
    Installer and Uninstaller for Excalibur 32-bit

.DESCRIPTION
    Handles installation and Uninstallation of Excalibur 32-bit
    Default install directory: C:\tools\Excalibur 32-bit

.EXAMPLES
    .\installer.ps1 -Action install
    .\installer.ps1 -Action install -TargetDir "D:\Apps\Excalibur 32-bit"
    .\installer.ps1 -Action uninstall
#>

param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("install", "uninstall")]
    [string]$Action,

    [string]$TargetDir,

    [switch]$Quiet
)

$AppName = "Excalibur 32-bit"
$InstallBaseDir = "C:\tools"
$ZipUrl = "https://www.hpcalc.org/other/pc/excal32_200.zip"
$AppExecutable = "Excal32.exe"
$AppHelpFile = "Excal32.chm"

$AppDir = "$InstallBaseDir\$AppName"
$ZipFile = "$Env:TEMP\$AppName.zip"

$DesktopShortcutPath = "$Env:Public\Desktop\$AppName.lnk"
$StartMenuDir = "$Env:ProgramData\Microsoft\Windows\Start Menu\Programs\$AppName"
$StartMenuShortcutPath = "$StartMenuDir\$AppName.lnk"
$StartMenuHelpShortcutPath = "$StartMenuDir\$AppName Help.lnk"

function Install-App {
    if (-not $TargetDir) { $TargetDir = "$InstallBaseDIr\$AppName" }
    
    if (-not $Quiet) { Write-Host "Downloading $AppName..." }
    Invoke-WebRequest -Uri $ZipUrl -OutFile $ZipFile
    
    if (!(Test-Path "$InstallBaseDir")) { New-Item -Path "$InstallBaseDir" -ItemType Directory -Force | Out-Null }
    
    if (-not $Quiet) { Write-Host "Extracting to $AppDir..." }
    Expand-Archive -Path $ZipFile -DestinationPath $AppDir -Force | Out-Null
    Move-Item $AppDir\Excal32_200\* -Destination $AppDir\
    Remove-Item $AppDir\Excal32_200
    
    # Optionally crate shortcuts
    $WScriptShell = New-Object -ComObject WScript.Shell
    $DesktopShortcut = $WScriptShell.createShortcut($DesktopShortcutPath)
    $DesktopShortcut.TargetPath = "$AppDir\$AppExecutable"
    $DesktopShortcut.Save()
    
    if (!(Test-Path "$StartMenuDir")) { New-Item -Path "$StartMenuDir" -ItemType Directory -Force | Out-Null }
    
    $WScriptShell = New-Object -ComObject WScript.Shell
    $StartMenuShortcut = $WScriptShell.createShortcut($StartMenuShortcutPath)
    $StartMenuShortcut.TargetPath = "$AppDir\$AppExecutable"
    $StartMenuShortcut.Save()

    $WScriptShell = New-Object -ComObject WScript.Shell
    $StartMenuShortcut = $WScriptShell.createShortcut($StartMenuHelpShortcutPath)
    $StartMenuShortcut.TargetPath = "$AppDir\$AppHelpFile"
    $StartMenuShortcut.Save()
    
    if (-not $Quiet) { Write-Host "Done. $AppName installed to $AppDir" }
}

function Uninstall-App {
    $AppDir = "$InstallBaseDir\$AppName"
    $DesktopShortcutPath = "$Env:Public\Desktop\$AppName.lnk"
    if (-not $Quiet) { Write-Host "Uninstalling $AppName..." }
    Remove-Item -Recurse -Force $AppDir -ErrorAction SilentlyContinue
    
    Remove-Item -Force $DesktopShortcutPath, $StartMenuShortcutPath, $StartMenuHelpShortcutPath -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force "$StartMenuDir" -ErrorAction SilentlyContinue
    if (-not $Quiet) { Write-Host "$AppName uninstalled successfully" }
}

switch ($Action.ToLower()) {
    "install" { Install-App }
    "uninstall" { uninstall-App }
    default {
        Write-Error "Invalid action '$Action'. Use 'install' or 'uninstall'."
        exit 1
    }
}
