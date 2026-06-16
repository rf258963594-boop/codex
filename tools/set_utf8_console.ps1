$ErrorActionPreference = "Stop"

[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$global:OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

try {
  chcp 65001 | Out-Null
} catch {
  # Some restricted shells do not allow changing the active code page.
}

Write-Host "UTF-8 console settings applied for this PowerShell session."
