param(
    [string]$SourcePath = (Join-Path $PSScriptRoot '..\MarkdownToWord\Samples\syntax-tour.md'),
    [string]$OutputPath = (Join-Path $PSScriptRoot '..\_output\syntax-tour.docx'),
    [string[]]$Extra = @()
)

$exe = Join-Path $PSScriptRoot '..\MarkdownToWord\bin\Debug\net10.0-windows10.0.19041.0\win-x64\MarkdownToWord.exe'
$outFile = Join-Path $PSScriptRoot '..\_output\cli.out.txt'
$errFile = Join-Path $PSScriptRoot '..\_output\cli.err.txt'
New-Item -ItemType Directory -Force -Path (Split-Path $OutputPath) | Out-Null

$arguments = @('--convert', $SourcePath, $OutputPath, '--validate') + $Extra
$process = Start-Process -FilePath $exe -ArgumentList $arguments -NoNewWindow -Wait -PassThru `
    -RedirectStandardOutput $outFile -RedirectStandardError $errFile

Write-Host ("EXIT=" + $process.ExitCode)
Write-Host "--- stdout ---"
Get-Content $outFile -ErrorAction SilentlyContinue
Write-Host "--- stderr ---"
Get-Content $errFile -ErrorAction SilentlyContinue
