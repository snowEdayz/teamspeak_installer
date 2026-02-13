<#
    .SYNOPSIS
    TeamSpeak 3 中文安装向导

    .DESCRIPTION
    一个从脚本同目录读取安装包并安装 TeamSpeak 3 客户端与汉化包的脚本。

    .PARAMETER ts3FileHost
    为 files.teamspeak-services.com 使用反向代理。

    .PARAMETER location
    指定 TeamSpeak 3 客户端的安装路径。

    .PARAMETER currentUser
    启用仅为本用户安装。

    .EXAMPLE
    iex "& {$(irm https://gist.github.com/SummonHIM/83b2d102d184167c733d05226e1bc10a/raw/installTS3.ps1)}"
    简易使用命令。

    .EXAMPLE
    iex "& {$(irm https://gist.github.com/SummonHIM/83b2d102d184167c733d05226e1bc10a/raw/installTS3.ps1)} -ts3FileHost xxyyzz.com -location `"$env:LocalAPPData\Programs\TeamSpeak 3 Client`" -currentUser"
    全参数使用示例。

    .EXAMPLE
    iex "& {$(irm http://ghproxy.com/https://gist.github.com/SummonHIM/83b2d102d184167c733d05226e1bc10a/raw/installTS3.ps1)}"
    使用 GhProxy.com 反向代理。

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

function getInstallerFiles {
    $Global:scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $Global:clientFileName = "ts.exe"
    $Global:zhCNTranslationFileName = "Chinese_Translation_zh-CN.ts3_translation"
    $Global:clientFilePath = Join-Path $Global:scriptDir $Global:clientFileName
    $Global:zhCNTranslationFilePath = Join-Path $Global:scriptDir $Global:zhCNTranslationFileName

    if (!(Test-Path $Global:clientFilePath)) {
        Write-Error "未找到安装包：$Global:clientFilePath"
        Exit 3
    }

    if (!(Test-Path $Global:zhCNTranslationFilePath)) {
        Write-Error "未找到汉化包：$Global:zhCNTranslationFilePath"
        Exit 4
    }

    Write-Host "找到安装包：$Global:clientFilePath" -ForegroundColor Green
    Write-Host "找到汉化包：$Global:zhCNTranslationFilePath" -ForegroundColor Green
}

function uninstallOverwolf {
    Write-Host "正在检查并卸载 Overwolf（如果已安装）…" -ForegroundColor Yellow
    $keys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $overwolfEntries = foreach ($key in $keys) {
        Get-ItemProperty -Path $key -ErrorAction SilentlyContinue | Where-Object {
            $_.DisplayName -and $_.DisplayName -match "Overwolf"
        }
    }

    if ($null -eq $overwolfEntries -or $overwolfEntries.Count -eq 0) {
        Write-Host "未检测到 Overwolf，跳过卸载。" -ForegroundColor Green
        return
    }

    foreach ($entry in $overwolfEntries) {
        if ($entry.QuietUninstallString) {
            Write-Host "执行静默卸载：$($entry.DisplayName)" -ForegroundColor Yellow
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c $($entry.QuietUninstallString)" -Wait
        }
        elseif ($entry.UninstallString) {
            Write-Host "执行卸载：$($entry.DisplayName)" -ForegroundColor Yellow
            $uninstallCmd = $entry.UninstallString
            if ($uninstallCmd -notmatch "(?i)/quiet|/qn|/s") {
                if ($uninstallCmd -match "msiexec") {
                    $uninstallCmd = "$uninstallCmd /qn"
                }
                else {
                    $uninstallCmd = "$uninstallCmd /S"
                }
            }
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallCmd" -Wait
        }
    }
}

Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "* TeamSpeak 3 中文安装向导" -ForegroundColor Blue
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "* [1]安装 TS3" -ForegroundColor Green -NoNewline; Write-Host " > [2]安装汉化包 > [3] 完成"
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
getInstallerFiles
Write-Host "正在启动自动安装向导。安装位置为：$location、仅为本用户安装：$currentUser…" -ForegroundColor Yellow
if ($currentUser) {
    Invoke-Expression -Command "`"$Global:clientFilePath`" /CurrentUser /S /D=$location"
    $shortcutLocation = [Environment]::GetFolderPath("Desktop") + "\TeamSpeak 3 Client.lnk"
} else {
    Invoke-Expression -Command "`"$Global:clientFilePath`" /S /D=$location"
    $shortcutLocation = [Environment]::GetFolderPath("CommonDesktopDirectory") + "\TeamSpeak 3 Client.lnk"
}
Write-Host "正在安装中，请稍后…" -ForegroundColor Yellow
while (!(Test-Path "$shortcutLocation")) {
    Start-Sleep 1
}
Write-Host "安装结束，正在启动 TeamSpeak 3…" -ForegroundColor Green
Invoke-Expression -Command "& '$shortcutLocation'"
uninstallOverwolf

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "* [1]安装 TS3 > [2]安装汉化包 > [3] 完成" -ForegroundColor Green
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "正在启动汉化包安装向导。请手动互动界面来安装汉化包…" -ForegroundColor Green
Invoke-Expression -Command "`"$Global:zhCNTranslationFilePath`""

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "* [1]安装 TS3 > [2]安装汉化包 > [3] 完成" -ForegroundColor Green
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
