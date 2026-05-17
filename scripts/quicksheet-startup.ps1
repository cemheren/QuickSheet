<#
.SYNOPSIS
    Launch QuickSheet in desktop-wallpaper mode at user logon.

.DESCRIPTION
    Designed to be wired into Windows Task Scheduler ("At log on") or dropped as a
    shortcut into the shell:startup folder. By default it just runs
    `dotnet run -c Release --project ExcelConsole.csproj -- --desktop` from the
    repo root. With -Update it first git-pulls and rebuilds, aborting startup if
    either step fails (so a bad push doesn't leave you with no wallpaper at
    next reboot).

.PARAMETER Update
    Pull from origin/main and rebuild before launching. Aborts on non-zero exit.

.PARAMETER CsvPath
    Optional CSV file to open. Default: QuickSheet's autosave.

.EXAMPLE
    .\scripts\quicksheet-startup.ps1
    .\scripts\quicksheet-startup.ps1 -Update
    .\scripts\quicksheet-startup.ps1 -CsvPath C:\Users\me\dashboard.csv

.NOTES
    See docs/install-startup.md for Task Scheduler walkthrough.
#>
[CmdletBinding()]
param(
    [switch]$Update,
    [string]$CsvPath = ""
)

$ErrorActionPreference = 'Stop'

# Repo root is the parent of the scripts/ directory this file lives in.
$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location -Path $RepoRoot

if ($Update) {
    Write-Host "QuickSheet: pulling latest from origin/main..."
    & git pull --rebase --autostash origin main
    if ($LASTEXITCODE -ne 0) {
        Write-Error "git pull failed (exit $LASTEXITCODE); aborting startup."
        exit 1
    }

    Write-Host "QuickSheet: rebuilding..."
    & dotnet build .\ExcelConsole.csproj -c Release --nologo
    if ($LASTEXITCODE -ne 0) {
        Write-Error "dotnet build failed (exit $LASTEXITCODE); aborting startup."
        exit 1
    }
}

# Detach so the logon shell doesn't block on dotnet.
$dotnetArgs = @('run', '-c', 'Release', '--project', '.\ExcelConsole.csproj', '--', '--desktop')
if ($CsvPath) { $dotnetArgs += $CsvPath }

Start-Process -FilePath 'dotnet' `
              -ArgumentList $dotnetArgs `
              -WorkingDirectory $RepoRoot `
              -WindowStyle Hidden
