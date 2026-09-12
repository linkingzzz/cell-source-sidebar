# 构建「表格批注插件」单文件安装包（不需要管理员权限，不需要第三方工具）
# 产物: installer\output\表格批注插件-Setup.exe
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$repo = Split-Path -Parent $here
$build = Join-Path $here "build"
$outDir = Join-Path $here "output"
$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"

# 1) 组装 payload（全部 ASCII 文件名，安装时再改成中文名）
$payload = Join-Path $build "payload"
if (Test-Path $payload) { Remove-Item -LiteralPath $payload -Recurse -Force }
New-Item -ItemType Directory -Force -Path (Join-Path $payload "workbook") | Out-Null
Copy-Item -LiteralPath (Join-Path $here "manifest.xml") -Destination (Join-Path $payload "manifest.xml")
Copy-Item -LiteralPath (Join-Path $here "assets\表格批注插件.xlsx") -Destination (Join-Path $payload "workbook\annotation-workbook.xlsx")
Copy-Item -LiteralPath (Join-Path $repo "wps-plugin") -Destination (Join-Path $payload "wps") -Recurse
$count = (Get-ChildItem $payload -Recurse -File).Count
Write-Host "payload: $count 个文件"

# 2) 打包成 payload.zip
$zip = Join-Path $build "payload.zip"
if (Test-Path $zip) { Remove-Item -LiteralPath $zip -Force }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($payload, $zip, [System.IO.Compression.CompressionLevel]::Optimal, $false)
Write-Host "payload.zip: $((Get-Item $zip).Length) 字节"

# 3) 编译安装程序
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$exe = Join-Path $outDir "表格批注插件-Setup.exe"
$icon = Join-Path $here "assets\app.ico"
$cscArgs = @(
    "/nologo", "/target:winexe", "/codepage:65001", "/optimize+",
    "/out:$exe",
    "/resource:$zip,payload.zip",
    "/r:System.dll", "/r:System.Core.dll", "/r:System.Windows.Forms.dll",
    "/r:System.Drawing.dll", "/r:System.Xml.dll",
    "/r:System.IO.Compression.dll", "/r:System.IO.Compression.FileSystem.dll",
    (Join-Path $here "src\installer.cs")
)
if (Test-Path $icon) { $cscArgs = @("/win32icon:$icon") + $cscArgs }
& $csc @cscArgs
if ($LASTEXITCODE -ne 0) { throw "csc 编译失败，退出码 $LASTEXITCODE" }
Write-Host ""
Write-Host "构建完成: $exe  ($((Get-Item $exe).Length) 字节)"