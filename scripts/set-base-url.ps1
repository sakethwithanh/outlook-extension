<#
    Stamps your GitHub Pages URL into manifest.xml.

    Usage:
      .\scripts\set-base-url.ps1 -BaseUrl "https://ragirollasaketh.github.io/outlook-export-thread"

    Re-runnable: it rewrites from manifest.template.xml each time, so you can
    point the manifest at a new host without hand-editing.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$BaseUrl
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$template = Join-Path $root 'manifest.template.xml'
$target = Join-Path $root 'manifest.xml'

if (-not (Test-Path $template)) {
    throw "manifest.template.xml not found at $template"
}

$clean = $BaseUrl.TrimEnd('/')
if ($clean -notmatch '^https://') {
    throw "BaseUrl must start with https:// (Office rejects plain HTTP)"
}

# BASE_URL appears in the template already prefixed with https://
$hostPart = $clean -replace '^https://', ''

(Get-Content $template -Raw) -replace 'BASE_URL', $hostPart |
    Set-Content $target -Encoding UTF8 -NoNewline

Write-Host "manifest.xml written -> https://$hostPart" -ForegroundColor Green
Write-Host "Sideload this file in Outlook, or upload it in the M365 admin center."
