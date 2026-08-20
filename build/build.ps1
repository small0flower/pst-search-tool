# 建置腳本（在 Windows 上執行；需要 .NET SDK 8+ 或 Visual Studio Build Tools）
# 用法：powershell -ExecutionPolicy Bypass -File build\build.ps1
param(
    [switch]$SkipRestore
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$proj = Join-Path $root "src\PstSearch\PstSearch.csproj"
$dist = Join-Path $root "dist"

dotnet --version | Out-Null

if (-not $SkipRestore) { dotnet restore $proj }
dotnet build $proj -c Release -p:Platform=AnyCPU

if (Test-Path $dist) { Remove-Item $dist -Recurse -Force }
New-Item -ItemType Directory -Path $dist -Force | Out-Null
Copy-Item (Join-Path $root "src\PstSearch\bin\Release\net48\*") $dist -Recurse -Force

Write-Host ""
Write-Host "完成。應用程式輸出：$dist"
Write-Host "若要製作安裝程式，請以 Inno Setup 6 的 ISCC.exe 編譯 build\installer.iss"
