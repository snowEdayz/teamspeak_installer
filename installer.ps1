<#
    .SYNOPSIS
    TeamSpeak 3 中文安装向导

    .DESCRIPTION
    一个能从 files.teamspeak-services.com 上自动下载并安装 TeamSpeak 3 客户端和（语音）汉化包的脚本。

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
    [Parameter(ValueFromPipeline = $true)][string]$ts3FileHost = "files.teamspeak-services.com",
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

function getClientInfo {
    Write-Host "正在获取最新版客户端信息…" -ForegroundColor Yellow
    $clientInfo = Invoke-RestMethod "https://teamspeak.com/versions/client.json"
    if ($null -eq $clientInfo) {
        Write-Error "获取最新版客户端信息失败！"
        Exit 3
    }
    if ([Environment]::Is64BitOperatingSystem -eq $true) { $osBit = "x86_64" } else { $osBit = "x86" }
    $Global:clientUrl = $clientInfo.windows.$osBit.mirrors[0].PSObject.Properties.Value.replace('files.teamspeak-services.com', $ts3FileHost)
    $Global:clientFileName = $clientUrl.Split("/")[-1]
    $Global:clientCheckSum = $clientInfo.windows.$osBit.checksum
    $echoClientInfo = '{"下载链接": "' + $Global:clientUrl + '", "文件名称": "' + $Global:clientFileName + '", "校验码": "' + $Global:clientCheckSum + '"}' | ConvertFrom-Json
    Write-Host "获取到最新版客户端信息：$echoClientInfo" -ForegroundColor Green
}

function downloadClient {
    Write-Host "正在下载 $Global:clientFileName 至 $Global:tempPath$Global:clientFileName…" -ForegroundColor Yellow
    Invoke-WebRequest "$Global:clientUrl" -OutFile "$Global:tempPath$Global:clientFileName"
}

function checkClientSum {
    Write-Host "正在校验下载好的文件…" -ForegroundColor Yellow
    $downloadedSum = Get-FileHash -path "$Global:tempPath$Global:clientFileName"
    if ($Global:clientCheckSum -eq $downloadedSum.Hash) {
        Write-Host "文件已完整下载！" -ForegroundColor Green
    }
    else {
        Write-Error "文件校验码不匹配。文件未完整下载！"
        exit 4
    }
}

function downloadzhCNTranslation {
    $Global:zhCNTranslationFileName = "Chinese_Translation_zh-CN.ts3_translation"
    $zhCNTranslationUrl = "https://dl.tmspk.wiki/https:/github.com/VigorousPro/TS3-Translation_zh-CN/releases/download/snapshot/Chinese_Translation_zh-CN.ts3_translation"
    Write-Host "正在下载 $zhCNTranslationUrl 至 $Global:tempPath$Global:zhCNTranslationFileName…" -ForegroundColor Yellow
    Invoke-WebRequest "$zhCNTranslationUrl" -OutFile "$Global:tempPath$Global:zhCNTranslationFileName"
}

function downloadzhCNSoundpack {
    $Global:zhCNSoundpackFileName = "Chinese_Soundpack_zh-CN.ts3_soundpack"
    $zhCNSoundpackUrl = "https://addons-content.teamspeak.com/9c1ed6f5-15ac-458d-84fb-c0c2b025ef97/files/1/Chinese%20Soundpack%20(zh-CN).ts3_soundpack"
    Write-Host "正在下载 $zhCNSoundpackUrl 至 $Global:tempPath$Global:zhCNSoundpackFileName…" -ForegroundColor Yellow
    Invoke-WebRequest "$zhCNSoundpackUrl" -OutFile "$Global:tempPath$Global:zhCNSoundpackFileName"
}

$Global:tempPath = [System.IO.Path]::GetTempPath()
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "* TeamSpeak 3 中文安装向导" -ForegroundColor Blue
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "* [1]安装 TS3" -ForegroundColor Green -NoNewline; Write-Host " > [2]安装汉化包 > [3]安装语音包 > [4] 完成"
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
getClientInfo
downloadClient
checkClientSum
Write-Host "正在启动自动安装向导。安装位置为：$location、仅为本用户安装：$currentUser…" -ForegroundColor Yellow
if ($currentUser) {
    Invoke-Expression -Command "$Global:tempPath$Global:clientFileName /CurrentUser /S /D=$location"
    $shortcutLocation = [Environment]::GetFolderPath("Desktop") + "\TeamSpeak 3 Client.lnk"
} else {
    Invoke-Expression -Command "$Global:tempPath$Global:clientFileName /S /D=$location"
    $shortcutLocation = [Environment]::GetFolderPath("CommonDesktopDirectory") + "\TeamSpeak 3 Client.lnk"
}
Write-Host "正在安装中，请稍后…" -ForegroundColor Yellow
while (!(Test-Path "$shortcutLocation")) {
    Start-Sleep 1
}
Write-Host "安装结束，正在启动 TeamSpeak 3…" -ForegroundColor Green
Invoke-Expression -Command "& '$shortcutLocation'"

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "* [1]安装 TS3 > [2]安装汉化包" -ForegroundColor Green -NoNewline; Write-Host " > [3]安装语音包 > [4] 完成"
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
downloadzhCNTranslation
Write-Host "正在启动汉化包安装向导。请手动互动界面来安装汉化包…" -ForegroundColor Green
Invoke-Expression -Command "$Global:tempPath$Global:zhCNTranslationFileName"

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "* [1]安装 TS3 > [2]安装汉化包 > [3]安装语音包" -ForegroundColor Green -NoNewline; Write-Host " > [4] 完成"
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
downloadzhCNSoundpack
Write-Host "正在启动语音包安装向导。请手动互动界面来安装语音包…" -ForegroundColor Green
Invoke-Expression -Command "$Global:tempPath$Global:zhCNSoundpackFileName"

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "* [1]安装 TS3 > [2]安装汉化包 > [3]安装语音包 > [4] 完成" -ForegroundColor Green
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
function choiceClean {
    $readClean = Read-Host "是否要清理下载的内容？[Y/n]"
    Switch ($readClean) {
        Y { $readCleanChoice = $true }
        N { $readCleanChoice = $false }
        Default {
            Write-Warning "未知选项，使用默认选项：是（Y）。"
            $readCleanChoice = $true
        }
    }
    return $readCleanChoice
}

if (choiceClean) {
    Write-Warning "正在删除 $Global:tempPath$Global:clientFileName…"
    Remove-Item "$Global:tempPath$Global:clientFileName"
    Write-Warning "正在删除 $Global:tempPath$Global:zhCNTranslationFileName…"
    Remove-Item "$Global:tempPath$Global:zhCNTranslationFileName"
    Write-Warning "正在删除 $Global:tempPath$Global:zhCNSoundpackFileName…"
    Remove-Item "$Global:tempPath$Global:zhCNSoundpackFileName"
}
