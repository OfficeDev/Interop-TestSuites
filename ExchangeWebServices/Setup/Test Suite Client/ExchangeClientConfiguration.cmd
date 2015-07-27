@echo off
pushd %~dp0
set powershellExist=0
if exist %SystemRoot%\syswow64\WindowsPowerShell\v1.0\powershell.exe (set powershellExist=1
%SystemRoot%\syswow64\WindowsPowerShell\v1.0\powershell.exe -command invoke-command "{if(!(New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {write-host "ExchangeClientConfiguration.cmd is not run as administrator";exit 2}else{Set-ExecutionPolicy RemoteSigned -force}}"
)
if %ERRORLEVEL% equ 2 (
echo You need to run ExchangeClientConfiguration.cmd using "Run as administrator"!
Pause
Exit
)
if exist %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe (set powershellExist=1
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe -command invoke-command "{if(!(New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {write-host "ExchangeClientConfiguration.cmd is not run as administrator";exit 2}else{Set-ExecutionPolicy RemoteSigned -force}}"
)
if %ERRORLEVEL% equ 2 (
echo You need to run ExchangeClientConfiguration.cmd using "Run as administrator"!
Pause
Exit
)
if %powershellExist% equ 1 (PowerShell.exe -ExecutionPolicy ByPass .\ExchangeClientConfiguration.ps1) else (echo PowerShell is not installed, you should install it first.)
Pause