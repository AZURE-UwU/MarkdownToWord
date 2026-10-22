param([Parameter(Mandatory = $true)][string]$Latex)

$exe = Join-Path $PSScriptRoot '..\MarkdownToWord\bin\Debug\net10.0-windows10.0.19041.0\win-x64\MarkdownToWord.exe'
$outFile = Join-Path $PSScriptRoot '..\_output\math.out.txt'
$errFile = Join-Path $PSScriptRoot '..\_output\math.err.txt'

$quoted = '"' + ($Latex -replace '"', '\"') + '"'
Start-Process -FilePath $exe -ArgumentList @('--math-probe', $quoted) -NoNewWindow -Wait `
    -RedirectStandardOutput $outFile -RedirectStandardError $errFile | Out-Null

$raw = Get-Content $outFile -Raw -Encoding UTF8
($raw -replace '><', ">`n<")
if ((Get-Item $errFile).Length -gt 0) {
    Write-Host "--- stderr ---"
    Get-Content $errFile
}
