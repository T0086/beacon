# beacon 安装脚本
# 用法：双击「安装.bat」，或右键此文件 → 使用 PowerShell 运行
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------- self-elevate
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "需要管理员权限，正在提权..."
    Start-Process powershell -Verb RunAs -ArgumentList @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`""
    )
    exit
}

$src = Join-Path $PSScriptRoot 'app'
$dst = Join-Path $env:LOCALAPPDATA 'beacon'

if (-not (Test-Path $src)) {
    Write-Host "错误：找不到 app 目录。请把整个安装包解压后再运行。"
    Read-Host "按回车退出"
    exit 1
}

Write-Host ""
Write-Host "=== beacon 安装 ===" -ForegroundColor Cyan
Write-Host "目标目录: $dst"
Write-Host ""

# ---------------------------------------------------------------- stop running instance
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -match 'beacon\.ps1' } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
Start-Sleep -Milliseconds 800

# ---------------------------------------------------------------- copy files
if (Test-Path $dst) {
    # keep user data: notes.json / imported.json / state.json
    Get-ChildItem $dst -File | Where-Object { $_.Name -notin 'notes.json', 'imported.json', 'state.json' } | Remove-Item -Force -ErrorAction SilentlyContinue
    Get-ChildItem $dst -Directory | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
} else {
    New-Item -ItemType Directory -Path $dst -Force | Out-Null
}
Copy-Item "$src\*" $dst -Recurse -Force
Write-Host "[1/4] 文件已复制"

# ---------------------------------------------------------------- scheduled task (elevated, needed for LHM temperature)
$vbs = Join-Path $dst '启动 beacon.vbs'
if (-not (Test-Path $vbs)) {
    Write-Host "错误：找不到启动器 $vbs"
    Read-Host "按回车退出"
    exit 1
}

$action = New-ScheduledTaskAction -Execute 'C:\Windows\System32\wscript.exe' -Argument "`"$vbs`"" -WorkingDirectory $dst
$trigger = New-ScheduledTaskTrigger -AtLogOn
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -MultipleInstances IgnoreNew `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -DontStopOnIdleEnd

Register-ScheduledTask -TaskName 'Beacon' -Action $action -Trigger $trigger `
    -Principal $principal -Settings $settings `
    -Description 'beacon: floating desktop widget (elevated for hardware sensors)' -Force | Out-Null
Write-Host "[2/4] 计划任务已创建（登录时自动启动，最高权限）"

# ---------------------------------------------------------------- drop the old startup shortcut
$lnk = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup\beacon.lnk'
if (Test-Path $lnk) {
    Remove-Item $lnk -Force -ErrorAction SilentlyContinue
    Write-Host "[3/4] 已移除旧的启动文件夹快捷方式"
} else {
    Write-Host "[3/4] 无旧快捷方式"
}

# ---------------------------------------------------------------- start it
Start-ScheduledTask -TaskName 'Beacon'
Start-Sleep -Seconds 4
$running = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -match 'beacon\.ps1' }).Count
Write-Host "[4/4] 已启动（检测到 $running 个实例）"

Write-Host ""
Write-Host "=== 安装完成 ===" -ForegroundColor Green
Write-Host ""
Write-Host "接下来："
Write-Host "  - 桌面上会出现一条悬浮长条，位置在屏幕右上角（或多屏的副屏左上）"
Write-Host "  - 点「设置」齿轮可以粘贴 .ics 日历链接导入课程表"
Write-Host "  - 第一次读 CPU 温度可能需要几秒"
Write-Host ""
Write-Host "卸载：删除计划任务「Beacon」和目录 $dst 即可。"
Write-Host ""
Read-Host "按回车关闭"
