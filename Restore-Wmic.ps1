#Requires -RunAsAdministrator
<#
    Restore-Wmic.ps1

    Restores WMIC on Windows 11 24H2/25H2 after KB5120998 removed it.

    Order of preference:
      1. Genuine Microsoft-signed files already on this PC (WinSxS / CbsTemp).
      2. The kit shipped next to this script (WMIC.exe + support files).
    Downloads nothing. Every file is checked before it is copied: either a
    valid Microsoft signature, or an exact SHA-256 match with the known
    genuine Microsoft file (needed because WMIC is only catalog-signed, and
    its catalog is removed together with WMIC).

    Usage:   powershell -ExecutionPolicy Bypass -File .\Restore-Wmic.ps1
    Undo:    powershell -ExecutionPolicy Bypass -File .\Restore-Wmic.ps1 -Uninstall
    Exit code: 0 = WMIC works, 1 = it does not.
#>
[CmdletBinding()]
param(
    [switch]$Uninstall,
    # Optional: folder holding a WMIC kit (default: the folder this script is in)
    [string]$FromKit,
    # Optional: write the local files out to a folder so they can be carried to another PC
    [string]$ExportKit
)

$ErrorActionPreference = 'Stop'
$Wbem   = Join-Path $env:SystemRoot 'System32\Wbem'
$Marker = Join-Path $Wbem 'wmic-restore.installed'

# Files that make up a working WMIC install.
$BaseFiles = @('texttable.xsl','textvaluelist.xsl','rawxml.xsl','xsl-mappings.xml','cli.mof','cliegaliases.mof')
$ResFiles  = @('csv.xsl','hform.xsl','htable.xsl','mof.xsl','xml.xsl','cli.mfl','cliegaliases.mfl')

# SHA-256 of the genuine Microsoft files shipped in the kit (WMIC 10.0.26100.1150, amd64, en-US).
$KnownGood = @{
    'WMIC.exe'               = '632F015286F8E8A632B2274B42F68FF8C2F4481537FC47CD70AD5ABFD0AFAA4E'
    'cli.mof'                = '52DA5B9B304DE82395A97DC6920F52093595F62F4D50710F00DF129E7EE136B3'
    'cliegaliases.mof'       = '37779AC6F8CCF7E6830405C24B719A18EE949FD4F6D7FB6A38891186D183209B'
    'rawxml.xsl'             = 'D2ECB9C04A584489548A8F4FDFD656ABBC3F7F163083B007EFBE744368CE68F1'
    'texttable.xsl'          = '402ADC48C7A0BB166108C3BAD4B950F4E177B8E55C089540D2FD919E61C8E840'
    'textvaluelist.xsl'      = 'D7F6A65E64534DD362E6A5D44C67C184A67627F7CFD0D69F6BA9E642A470AD57'
    'xsl-mappings.xml'       = 'C136476A60D3874A281E89AC4DBB814C692B40EC6A0BD34F3815A4CDA943FA77'
    'en-US\cli.mfl'          = 'E68CC9E5CEB49F4D3CEE6FCFFD25879B1843D5DD726D478BB8FC514C10B7D27E'
    'en-US\cliegaliases.mfl' = '27D6E27BB0B900C0E49C41B22E396ED3FD5A7834BE5CC31C5105A7CC4D156C2C'
    'en-US\csv.xsl'          = 'CB88C3BC35334DE33B7548B98B4ECF83127A47BBF8407858FCF834D007FC655B'
    'en-US\hform.xsl'        = '69BA30A8D4D2F44EEAABFFB9B629FBFA2B7E9B5486820DF33AB6115DF53E7BCE'
    'en-US\htable.xsl'       = '50866E024E226986680B1C43C49FF105CCA03193F241A49A9351E634E4F8DF2D'
    'en-US\mof.xsl'          = '9AE0FFF4D0DF14496E69D841963D5E8B3A13DB01E69BE4A2B24524CE8BA3C2A0'
    'en-US\xml.xsl'          = '3F4A2D2ABCB9A652DA6E26A15F6B561C0F59B8A53644E23BE6886A80E015BF16'
}

function Write-Step { param($m) Write-Host "[*] $m" -ForegroundColor Cyan }
function Write-Ok   { param($m) Write-Host "[+] $m" -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "[!] $m" -ForegroundColor Yellow }
function Write-Err  { param($m) Write-Host "[x] $m" -ForegroundColor Red }

function Test-Wmic {
    # Real functional test: WMIC can exit 0 and print NOTHING when its XSL
    # stylesheets are missing, so "the exe exists" is not good enough.
    $exe = Join-Path $Wbem 'WMIC.exe'
    if (-not (Test-Path $exe)) { return $false }
    $o = Join-Path $env:TEMP ("wmic_probe_{0}.txt" -f $PID)
    $e = Join-Path $env:TEMP ("wmic_probe_{0}.err" -f $PID)
    try {
        Start-Process $exe -ArgumentList 'os','get','Caption','/value' `
            -NoNewWindow -Wait -RedirectStandardOutput $o -RedirectStandardError $e | Out-Null
        $out = (Get-Content $o -Raw -ErrorAction SilentlyContinue)
        return [bool]($out -and $out -match 'Caption=\S')
    } catch { return $false }
      finally { Remove-Item $o,$e -Force -ErrorAction SilentlyContinue }
}

function Test-Trusted {
    # True if the file is signed by Microsoft (embedded or catalog) or is
    # byte-for-byte the known genuine file.
    param([string]$Path, [string]$Key)
    if ($KnownGood.ContainsKey($Key) -and (Get-FileHash $Path -Algorithm SHA256).Hash -eq $KnownGood[$Key]) { return $true }
    try {
        $sig = Get-AuthenticodeSignature $Path
        return ($sig.Status -eq 'Valid' -and $sig.SignerCertificate.Subject -match 'O=Microsoft Corporation')
    } catch { return $false }
}

function Get-Newest {
    param([string[]]$Paths)
    $Paths | Where-Object { $_ } |
        Sort-Object {
            if ($_ -match '_10\.0\.(\d+)\.(\d+)_') { [version]("10.0.{0}.{1}" -f $matches[1],$matches[2]) }
            else { [version]'0.0.0.0' }
        } -Descending | Select-Object -First 1
}

function Find-LocalSource {
    # NOTE: WinSxS truncates long component names, so WMIC lives in directories
    # called "...w..ommand-line-utility..." - a "*wmic*" search finds NOTHING.
    $WinSxS = Join-Path $env:SystemRoot 'WinSxS'
    $arch   = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64' -or $env:PROCESSOR_ARCHITEW6432 -eq 'ARM64') { 'arm64_' }
              elseif ([Environment]::Is64BitOperatingSystem) { 'amd64_' } else { 'x86_' }

    # WMIC.exe - WinSxS or the CbsTemp FoD payload. Inside CbsTemp, the 'f'
    # and 'r' subfolders are binary deltas, not usable executables.
    $exeCandidates = @()
    $exeCandidates += Get-ChildItem $WinSxS -Directory -Filter "$arch*ommand-line-utility_*" -ErrorAction SilentlyContinue |
                        ForEach-Object { Join-Path $_.FullName 'WMIC.exe' } | Where-Object { Test-Path $_ }
    $cbs = Join-Path $env:SystemRoot 'CbsTemp'
    if (Test-Path $cbs) {
        $exeCandidates += Get-ChildItem $cbs -Directory -Recurse -Filter "$arch*ommand-line-utility_*" -ErrorAction SilentlyContinue |
                            Where-Object { $_.Parent.Name -notin 'f','r' } |
                            ForEach-Object { Join-Path $_.FullName 'WMIC.exe' } | Where-Object { Test-Path $_ }
    }
    $exe  = Get-Newest $exeCandidates
    $base = Get-Newest (Get-ChildItem $WinSxS -Directory -Filter "$arch*d-line-utility-base_*" -ErrorAction SilentlyContinue |
                            Select-Object -Exp FullName)

    $culture = (Get-UICulture).Name
    $resDirs = Get-ChildItem $WinSxS -Directory -Filter "$arch*lity-base.resources_*" -ErrorAction SilentlyContinue |
                    Select-Object -Exp FullName
    $res = Get-Newest ($resDirs | Where-Object { $_ -match "_$([regex]::Escape($culture))_" })
    if (-not $res) { $res = Get-Newest ($resDirs | Where-Object { $_ -match '_en-us_' }) }

    if (-not $exe -or -not $base) { return $null }
    return @{ Name = 'this PC''s component store'; Exe = $exe; Base = $base; Res = $res }
}

function Get-KitSource {
    param([string]$Dir)
    if (-not $Dir -or -not (Test-Path (Join-Path $Dir 'WMIC.exe'))) { return $null }
    return @{ Name = "the bundled kit"; Exe = (Join-Path $Dir 'WMIC.exe'); Base = $Dir; Res = (Join-Path $Dir 'en-US') }
}

function Test-Source {
    # Every file that will be copied must pass Test-Trusted.
    param($Src)
    if (-not (Test-Trusted $Src.Exe 'WMIC.exe')) { Write-Warn "WMIC.exe from $($Src.Name) failed verification - skipping it."; return $false }
    foreach ($f in $BaseFiles) {
        $p = Join-Path $Src.Base $f
        if (-not (Test-Path $p)) { Write-Warn "$f missing in $($Src.Name) - skipping it."; return $false }
        if ($KnownGood.ContainsKey($f) -and $Src.Name -eq 'the bundled kit' -and -not (Test-Trusted $p $f)) {
            Write-Warn "$f in $($Src.Name) failed verification - skipping it."; return $false
        }
    }
    if ($Src.Res -and $Src.Name -eq 'the bundled kit') {
        foreach ($f in $ResFiles) {
            $p = Join-Path $Src.Res $f
            if ((Test-Path $p) -and -not (Test-Trusted $p "en-US\$f")) {
                Write-Warn "en-US\$f in $($Src.Name) failed verification - skipping it."; return $false
            }
        }
    }
    return $true
}

function Install-From {
    param($Src)
    $enUS = Join-Path $Wbem 'en-US'
    if (-not (Test-Path $enUS)) { New-Item -ItemType Directory -Path $enUS -Force | Out-Null }
    Copy-Item $Src.Exe (Join-Path $Wbem 'WMIC.exe') -Force
    foreach ($f in $BaseFiles) { Copy-Item (Join-Path $Src.Base $f) $Wbem -Force }
    if ($Src.Res -and (Test-Path $Src.Res)) {
        foreach ($f in $ResFiles) { $p = Join-Path $Src.Res $f; if (Test-Path $p) { Copy-Item $p $enUS -Force } }
    }
    Set-Content -Path $Marker -Value ("Installed by Restore-Wmic.ps1 from {0} on {1:u}" -f $Src.Name, (Get-Date)) -Encoding ASCII
}

# ---------------------------------------------------------------- uninstall --
if ($Uninstall) {
    if (-not (Test-Path $Marker)) {
        Write-Warn 'WMIC on this PC was not installed by this tool, so it is left alone.'
        exit 0
    }
    Write-Step 'Removing restored WMIC files...'
    $n = 0
    foreach ($f in @('WMIC.exe') + $BaseFiles) {
        $p = Join-Path $Wbem $f
        if (Test-Path $p) { Remove-Item $p -Force; $n++ }
    }
    foreach ($f in $ResFiles) {
        $p = Join-Path $Wbem "en-US\$f"
        if (Test-Path $p) { Remove-Item $p -Force; $n++ }
    }
    Remove-Item $Marker -Force
    Write-Ok "Removed $n file(s). WMIC is gone again."
    exit 0
}

Write-Host ''
Write-Host '=== WMIC restore (Windows 11 24H2/25H2) ===' -ForegroundColor White
Write-Host ("OS build: {0}" -f [Environment]::OSVersion.Version) -ForegroundColor DarkGray
Write-Host ''

# ------------------------------------------------------------- export a kit --
if ($ExportKit) {
    $src = Find-LocalSource
    if (-not $src -and (Test-Path (Join-Path $Wbem 'WMIC.exe'))) {
        $src = @{ Name = 'System32\Wbem'; Exe = (Join-Path $Wbem 'WMIC.exe'); Base = $Wbem; Res = (Join-Path $Wbem 'en-US') }
    }
    if (-not $src) { Write-Err 'No WMIC files on this PC to export.'; exit 1 }
    Write-Step "Exporting kit from $($src.Name) to $ExportKit"
    New-Item -ItemType Directory -Path (Join-Path $ExportKit 'en-US') -Force | Out-Null
    Copy-Item $src.Exe $ExportKit -Force
    foreach ($f in $BaseFiles) { $p = Join-Path $src.Base $f; if (Test-Path $p) { Copy-Item $p $ExportKit -Force } }
    if ($src.Res) { foreach ($f in $ResFiles) { $p = Join-Path $src.Res $f; if (Test-Path $p) { Copy-Item $p (Join-Path $ExportKit 'en-US') -Force } } }
    Copy-Item $PSCommandPath $ExportKit -Force -ErrorAction SilentlyContinue
    Write-Ok 'Kit exported. On the other PC run:  .\Restore-Wmic.ps1 -FromKit <folder>'
    exit 0
}

if (Test-Wmic) {
    Write-Ok 'WMIC already works on this PC. Nothing to do.'
    exit 0
}

# ------------------------------------------------ try each source in order --
$sources = @()
if ($FromKit) { $sources += Get-KitSource $FromKit }
else {
    Write-Step 'Looking for genuine WMIC files already on this PC...'
    $sources += Find-LocalSource
    $sources += Get-KitSource $PSScriptRoot
}
$sources = @($sources | Where-Object { $_ })
if (-not $sources) {
    Write-Err 'No WMIC files found on this PC and no kit next to this script.'
    exit 1
}

foreach ($src in $sources) {
    Write-Step "Trying $($src.Name)..."
    if (-not (Test-Source $src)) { continue }
    Write-Ok ("Verified WMIC {0}" -f (Get-Item $src.Exe).VersionInfo.FileVersion)
    Write-Step "Installing into $Wbem ..."
    Install-From $src
    Write-Step 'Testing WMIC for real (not just checking the file exists)...'
    if (Test-Wmic) {
        Write-Host ''
        Write-Ok 'SUCCESS - WMIC is working.'
        Write-Host ''
        & (Join-Path $Wbem 'WMIC.exe') os get Caption /value 2>$null | Where-Object { $_.Trim() }
        & (Join-Path $Wbem 'WMIC.exe') cpu get Name /value      2>$null | Where-Object { $_.Trim() }
        Write-Host ''
        Write-Warn 'A future Windows update, sfc /scannow or DISM /RestoreHealth may remove'
        Write-Warn 'it again. If that happens, just run this again.'
        exit 0
    }
    Write-Warn "WMIC from $($src.Name) did not work - trying the next option."
}

Write-Host ''
Write-Err 'Could not get WMIC working on this PC.'
exit 1
