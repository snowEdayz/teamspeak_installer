<#
    .SYNOPSIS
    TeamSpeak 3 中文安装向导

    .DESCRIPTION
    一个安装同目录 TeamSpeak 3 客户端与汉化包的脚本。

    .PARAMETER location
    指定 TeamSpeak 3 客户端的安装路径。

    .PARAMETER currentUser
    启用仅为本用户安装。

    .EXAMPLE
    iex "& {$(irm https://gist.github.com/SummonHIM/83b2d102d184167c733d05226e1bc10a/raw/installTS3.ps1)}"
    简易使用命令。

    .EXAMPLE
    .\installer.ps1 -location `"$env:LocalAPPData\Programs\TeamSpeak 3 Client`" -currentUser
    全参数使用示例。

    .LINK
    https://gist.github.com/SummonHIM/83b2d102d184167c733d05226e1bc10a
#>

param(
    [Parameter(ValueFromPipeline = $true)][string]$location = "$env:ProgramFiles\TeamSpeak 3 Client",
    [Parameter(ValueFromPipeline = $true)][switch]$currentUser = $false
)

if ($host.Version -le 5.1) {
    Write-Output ""
    Write-Output "你的 PowerShell 版本过低。请升级你的 PowerShell。"
    Write-Output "你可下载 Windows Management Framework 5.1 来升级你的 PowerShell 版本："
    Write-Output "https://docs.microsoft.com/powershell/scripting/windows-powershell/wmf/setup/install-configure?view=powershell-5.1"
    Write-Output ""
    Write-Output "按 Ctrl+C 或等待 60 秒后退出…"
    Start-Sleep 60
    Exit 1
}

if ([System.Environment]::OSVersion.Platform -ne "Win32NT") {
    Write-Error "本脚本仅支持 Windows 系统！"
    Exit 2
}

function assertLocalInstallers {
    $Global:scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $Global:clientFileName = "ts.exe"
    $Global:clientPath = Join-Path $Global:scriptDir $Global:clientFileName
    $Global:zhCNTranslationFileName = "Chinese_Translation_zh-CN.ts3_translation"
    $Global:zhCNTranslationPath = Join-Path $Global:scriptDir $Global:zhCNTranslationFileName

    if (!(Test-Path $Global:clientPath)) {
        Write-Error "未找到安装包：$Global:clientPath"
        Exit 3
    }
    if (!(Test-Path $Global:zhCNTranslationPath)) {
        Write-Error "未找到汉化包：$Global:zhCNTranslationPath"
        Exit 4
    }

    $echoClientInfo = '{"安装包": "' + $Global:clientPath + '", "汉化包": "' + $Global:zhCNTranslationPath + '"}' | ConvertFrom-Json
    Write-Host "已找到本地安装文件：$echoClientInfo" -ForegroundColor Green
}

function uninstallOverwolf {
    Write-Host "正在尝试移除 Overwolf（若存在）…" -ForegroundColor Yellow
    $uninstallEntries = @()
    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $registryPaths) {
        $uninstallEntries += Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName -like "*Overwolf*" }
    }

    foreach ($entry in $uninstallEntries) {
        $command = if ($entry.QuietUninstallString) { $entry.QuietUninstallString } else { $entry.UninstallString }
        if ([string]::IsNullOrWhiteSpace($command)) { continue }

        Write-Host "正在执行卸载命令：$($entry.DisplayName)" -ForegroundColor Yellow
        if ($command -match "msiexec") {
            if ($command -notmatch "/qn") { $command = "$command /qn /norestart" }
        }
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command" -Wait -WindowStyle Hidden
    }

    $fallbackUninstaller = @(
        "$env:ProgramFiles\Overwolf\OWUninstaller.exe",
        "$env:ProgramFiles(x86)\Overwolf\OWUninstaller.exe"
    )

    foreach ($uninstaller in $fallbackUninstaller) {
        if (Test-Path $uninstaller) {
            Write-Host "发现 Overwolf 卸载器，正在静默卸载：$uninstaller" -ForegroundColor Yellow
            Start-Process -FilePath $uninstaller -ArgumentList "/S" -Wait -WindowStyle Hidden
        }
    }

    Write-Host "Overwolf 处理步骤已完成。" -ForegroundColor Green
}

Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "* TeamSpeak 3 中文安装向导" -ForegroundColor Blue
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "* [1]安装 TS3" -ForegroundColor Green -NoNewline; Write-Host " > [2]安装汉化包 > [3] 完成"
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
assertLocalInstallers
Write-Host "正在启动自动安装向导。安装位置为：$location、仅为本用户安装：$currentUser…" -ForegroundColor Yellow
if ($currentUser) {
    Invoke-Expression -Command "`"$Global:clientPath`" /CurrentUser /S /D=$location"
    $shortcutLocation = [Environment]::GetFolderPath("Desktop") + "\TeamSpeak 3 Client.lnk"
} else {
    Invoke-Expression -Command "`"$Global:clientPath`" /S /D=$location"
    $shortcutLocation = [Environment]::GetFolderPath("CommonDesktopDirectory") + "\TeamSpeak 3 Client.lnk"
}
Write-Host "正在安装中，请稍后…" -ForegroundColor Yellow
while (!(Test-Path "$shortcutLocation")) {
    Start-Sleep 1
}
uninstallOverwolf
Write-Host "安装结束，正在启动 TeamSpeak 3…" -ForegroundColor Green
Invoke-Expression -Command "& '$shortcutLocation'"

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "* [1]安装 TS3 > [2]安装汉化包 > [3] 完成" -ForegroundColor Green
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "正在启动汉化包安装向导。请手动互动界面来安装汉化包…" -ForegroundColor Green
Invoke-Expression -Command "`"$Global:zhCNTranslationPath`""

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "* [1]安装 TS3 > [2]安装汉化包 > [3] 完成" -ForegroundColor Green
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
