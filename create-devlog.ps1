$date = Get-Date -Format "yyyyMMdd"
$displayDate = Get-Date -Format "yyyy-MM-dd"
$vaultPath = "C:\Users\kevin.crump\OneDrive - Candid\Documents\DevDocs\Dev Docs"
$devlogDir = Join-Path $vaultPath "devlog"
$fileName = "devlog $date.md"
$filePath = Join-Path $devlogDir $fileName

if (-not (Test-Path $devlogDir)) {
    New-Item -ItemType Directory -Path $devlogDir | Out-Null
}

if (-not (Test-Path $filePath)) {
    Set-Content -Path $filePath -Value "# Devlog - $displayDate`n`n" -Encoding UTF8
}

$encodedFile = [Uri]::EscapeDataString("devlog/devlog $date")
$uri = 'obsidian://open?vault=Dev%20Docs&file=' + $encodedFile
Start-Process $uri
