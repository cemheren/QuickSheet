# Build libghostty-vt from the official, checksum-pinned source tarball and
# install ghostty-vt.dll into native\<rid>\ so `dotnet build` copies it next
# to the QuickSheet executable.
#
# Requires: Zig 0.17.x on PATH, tar.exe (Windows 10+), and optionally
# minisign (signature check is skipped with a note if absent).
# Compatible with Windows PowerShell 5.1 and PowerShell 7.
#
# Optional environment:
#   $env:ZIG_TARGET      passed as -Dtarget=... to zig build
#   $env:LIBGHOSTTY_OUT  output directory (default: <repo>\native\<rid>)
#
# Usage: powershell -ExecutionPolicy Bypass -File scripts\build-libghostty.ps1

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # Invoke-WebRequest is very slow in 5.1 with progress

# ---- Pinned release (bump here only) ---------------------------------------
$LibghosttyVersion = '1.0.0'
$SrcUrl     = "https://release.files.ghostty.org/libghostty-vt/$LibghosttyVersion/libghostty-vt-$LibghosttyVersion.tar.gz"
$SrcSha256  = '4a6ef758be721923e1969e7b81b443dcc92d8dd3fb5a39600c03a8817a753676'
$DepsUrl    = "https://release.files.ghostty.org/libghostty-vt/$LibghosttyVersion/libghostty-vt-$LibghosttyVersion-deps.tar.gz"
$DepsSha256 = '4d5ea5e448e4c7ea8fb55ccbaa8668f2a97d44f438cee831e2b7ecd162d85c29'
$MinisignPubkey = 'RWQlAjJC23149WL2sEpT/l0QKy7hMIFhYdQOFy0Z7z7PbneUgvlsnYcV'
$ZigRequiredPrefix = '0.17.'
$ZigDownloadUrl = 'https://ziglang.org/download/'
# -----------------------------------------------------------------------------

function Fail([string]$Message) {
    [Console]::Error.WriteLine("error: $Message")
    exit 1
}

$RepoRoot = Split-Path -Parent $PSScriptRoot

# ---- Platform -> rid --------------------------------------------------------
$Arch = $env:PROCESSOR_ARCHITEW6432
if (-not $Arch) { $Arch = $env:PROCESSOR_ARCHITECTURE }
switch ($Arch) {
    'AMD64' { $Rid = 'win-x64' }
    'ARM64' { $Rid = 'win-arm64' }
    default { Fail "unsupported architecture: $Arch" }
}

# ---- Toolchain checks -------------------------------------------------------
# Prefer the inbox bsdtar: a GNU tar (e.g. from Git for Windows) can misread C:\ paths.
$Tar = Join-Path $env:SystemRoot 'System32\tar.exe'
if (-not (Test-Path $Tar)) {
    $tarCmd = Get-Command tar.exe -ErrorAction SilentlyContinue
    if (-not $tarCmd) { Fail "tar.exe not found (ships with Windows 10 and later)" }
    $Tar = $tarCmd.Source
}

$zigCmd = Get-Command zig -ErrorAction SilentlyContinue
$ZigVersion = ''
if ($zigCmd) {
    $ZigVersion = (& zig version 2>$null | Select-Object -First 1)
}
if (-not $ZigVersion -or -not $ZigVersion.StartsWith($ZigRequiredPrefix)) {
    Fail "Zig ${ZigRequiredPrefix}x is required (found: $(if ($ZigVersion) { $ZigVersion } else { 'none' })). Install it from $ZigDownloadUrl"
}
Write-Host "Using Zig $ZigVersion"

$HaveMinisign = [bool](Get-Command minisign -ErrorAction SilentlyContinue)

# ---- Work area (cleaned up in finally) --------------------------------------
$Work = Join-Path ([System.IO.Path]::GetTempPath()) ("libghostty-build-" + [System.Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $Work | Out-Null

try {
    $Stage = Join-Path $Work 'stage'
    $ZigCache = Join-Path $Work 'zig-cache'
    New-Item -ItemType Directory -Path $Stage, $ZigCache | Out-Null

    function Get-Verified([string]$Url, [string]$Sha) {
        $name = [System.IO.Path]::GetFileName(([System.Uri]$Url).LocalPath)
        $dest = Join-Path $Work $name
        Write-Host "Downloading $name"
        Invoke-WebRequest -Uri $Url -OutFile $dest -UseBasicParsing

        $actual = (Get-FileHash -Algorithm SHA256 -Path $dest).Hash
        if ($actual -ine $Sha) {
            Fail "SHA-256 mismatch for $name : expected $Sha, got $actual"
        }
        Write-Host "  sha256 OK: $name"

        if ($HaveMinisign) {
            Invoke-WebRequest -Uri "$Url.minisig" -OutFile "$dest.minisig" -UseBasicParsing
            & minisign -V -P $MinisignPubkey -m $dest | Out-Null
            if ($LASTEXITCODE -ne 0) { Fail "minisign signature check failed for $name" }
            Write-Host "  minisign OK: $name"
        } else {
            Write-Host "  note: minisign not found; signature skipped (pinned sha256 was checked)"
        }
        return $dest
    }

    $SrcTar  = Get-Verified $SrcUrl $SrcSha256
    $DepsTar = Get-Verified $DepsUrl $DepsSha256

    $SrcRoot  = Join-Path $Work 'src'
    $DepsRoot = Join-Path $Work 'deps'
    New-Item -ItemType Directory -Path $SrcRoot, $DepsRoot | Out-Null
    & $Tar -xzf $SrcTar -C $SrcRoot
    if ($LASTEXITCODE -ne 0) { Fail "failed to extract $SrcTar" }
    & $Tar -xzf $DepsTar -C $DepsRoot
    if ($LASTEXITCODE -ne 0) { Fail "failed to extract $DepsTar" }

    $SrcDir  = Join-Path $SrcRoot "libghostty-vt-$LibghosttyVersion"
    $DepsDir = Join-Path $DepsRoot "libghostty-vt-$LibghosttyVersion-deps"
    if (-not (Test-Path $SrcDir))  { Fail "unexpected archive layout: missing $SrcDir" }
    if (-not (Test-Path $DepsDir)) { Fail "unexpected archive layout: missing $DepsDir" }

    # ---- Build (fully offline: deps supplied via --system) ------------------
    Write-Host "Building libghostty-vt $LibghosttyVersion for $Rid (this can take a few minutes)"
    $env:ZIG_GLOBAL_CACHE_DIR = $ZigCache
    Push-Location $SrcDir
    try {
        $zigArgs = @('build', '--prefix', $Stage, '--system', $DepsDir,
                     '-Demit-lib-vt', '-Doptimize=ReleaseFast', '-Dcpu=baseline', '-Dstrip=true')
        if ($env:ZIG_TARGET) {
            Write-Host "Zig target override: $env:ZIG_TARGET"
            $zigArgs += "-Dtarget=$($env:ZIG_TARGET)"
        }
        & zig @zigArgs
        if ($LASTEXITCODE -ne 0) { Fail "zig build failed (exit $LASTEXITCODE)" }
    } finally {
        Pop-Location
    }

    $LibDll = Join-Path $Stage 'bin\ghostty-vt.dll'
    if (-not (Test-Path $LibDll)) {
        $found = Get-ChildItem -Path $Stage -Recurse -Filter 'ghostty-vt*.dll' -File | Select-Object -First 1
        if (-not $found) { Fail "build did not produce ghostty-vt.dll under $Stage" }
        $LibDll = $found.FullName
    }

    # ---- Install ------------------------------------------------------------
    if ($env:LIBGHOSTTY_OUT) { $DestDir = $env:LIBGHOSTTY_OUT }
    else { $DestDir = Join-Path $RepoRoot "native\$Rid" }
    New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
    $Dest = Join-Path $DestDir 'ghostty-vt.dll'
    Copy-Item -Path $LibDll -Destination $Dest -Force

    # License notices for what is linked into the binary (libghostty MIT + bundled third-party list).
    $LicSrc = Join-Path $SrcDir 'LICENSE'
    if (-not (Test-Path $LicSrc)) { Fail "source tarball has no LICENSE" }
    Copy-Item -Path $LicSrc -Destination (Join-Path $DestDir 'libghostty-vt-LICENSE.txt') -Force
    $ThirdParty = Join-Path $SrcDir 'deps\libghostty-vt-third-party.json'
    if (Test-Path $ThirdParty) {
        Copy-Item -Path $ThirdParty -Destination (Join-Path $DestDir 'libghostty-vt-third-party.json') -Force
    } else {
        Write-Host 'note: deps/libghostty-vt-third-party.json not present in the source tarball'
    }

    Write-Host ''
    Write-Host "Installed: $Dest"
    Write-Host "Next: run 'dotnet build ExcelConsole.csproj' so the library is copied to the output directory."
    Write-Host 'QuickSheet will then use it automatically (or set QUICKSHEET_LIBGHOSTTY to a full path).'
}
finally {
    if (Test-Path $Work) { Remove-Item -Recurse -Force $Work -ErrorAction SilentlyContinue }
}
