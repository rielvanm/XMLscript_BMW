param (
    [Parameter(Mandatory = $true)]
    [string]$MetaXml,

    [Parameter(Mandatory = $true)]
    [string]$TransXml
)

# Betrouwbaar scriptpad (werkt in PS 5.1 en 7+)
$ScriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ScriptDir)) {
    $ScriptDir = (Get-Location).Path
}

# Exe staat bij jou in dist
$ExePath = Join-Path $ScriptDir "dist\XML_BMW_EXE.exe"

if (-not (Test-Path $ExePath)) {
    Write-Error "Exe niet gevonden: $ExePath"
    exit 1
}

# Als user alleen bestandsnaam meegeeft: zoek in scriptmap
if (-not (Test-Path $MetaXml))  { $MetaXml  = Join-Path $ScriptDir $MetaXml }
if (-not (Test-Path $TransXml)) { $TransXml = Join-Path $ScriptDir $TransXml }

if (-not (Test-Path $MetaXml)) {
    Write-Error "RG_META bestand niet gevonden: $MetaXml"
    exit 1
}
if (-not (Test-Path $TransXml)) {
    Write-Error "RG_TRANS bestand niet gevonden: $TransXml"
    exit 1
}

Write-Host "Start BMW XML analyse..."
Write-Host "EXE  : $ExePath"
Write-Host "META : $MetaXml"
Write-Host "TRANS: $TransXml"
Write-Host ""

& $ExePath $MetaXml $TransXml
