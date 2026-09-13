# Run all configured TenderScrapers sources, then validate expected outputs.
# Place this file in:
# C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\02_TenderScrapers\daily_run.ps1

param(
    [string]$ProjectRoot = $PSScriptRoot,
    [string]$Python = "",
    [switch]$TranslateWorldBank,
    [switch]$WorldBankAutoInstallModels,
    [switch]$SkipConsolidation,
    [switch]$StopOnFailure
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = "C:\Users\Danie le Roux\OneDrive - Henry Fagan\Desktop\Data Analytics\02_TenderScrapers"
}

$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$RegistryPath = Join-Path $ProjectRoot "00_Config\sources_registry.xlsx"
$LogFolder = Join-Path $ProjectRoot "05_SystemLogs\DailyRun"

if (!(Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

$Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$LogFile = Join-Path $LogFolder "daily_run_$Timestamp.log"

function Resolve-PythonCommand {
    param([string]$RequestedPython)

    if (![string]::IsNullOrWhiteSpace($RequestedPython)) {
        return @{
            Command = $RequestedPython
            PrefixArgs = @()
        }
    }

    $PyLauncher = Get-Command py -ErrorAction SilentlyContinue
    if ($PyLauncher) {
        return @{
            Command = "py"
            PrefixArgs = @("-3")
        }
    }

    $PythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($PythonCmd) {
        return @{
            Command = "python"
            PrefixArgs = @()
        }
    }

    throw "Python was not found. Pass -Python 'C:\Path\To\python.exe' or add Python to PATH."
}

$PythonResolved = Resolve-PythonCommand -RequestedPython $Python
$PythonCommand = $PythonResolved.Command
$PythonPrefixArgs = $PythonResolved.PrefixArgs

Start-Transcript -Path $LogFile

try {
    Write-Host "Project root: $ProjectRoot"
    Write-Host "Registry: $RegistryPath"
    Write-Host "Log file: $LogFile"
    Write-Host "Python command: $PythonCommand $($PythonPrefixArgs -join ' ')"

    Set-Location $ProjectRoot

    $RunArgs = @()
    $RunArgs += $PythonPrefixArgs
    $RunArgs += "run_all_scrapers.py"
    $RunArgs += "--project-root"
    $RunArgs += $ProjectRoot
    $RunArgs += "--registry"
    $RunArgs += $RegistryPath

    if ($TranslateWorldBank) {
        $RunArgs += "--translate-worldbank"
    }

    if ($WorldBankAutoInstallModels) {
        $RunArgs += "--worldbank-auto-install-models"
    }

    if ($SkipConsolidation) {
        $RunArgs += "--skip-consolidation"
    }

    if ($StopOnFailure) {
        $RunArgs += "--stop-on-failure"
    }

    Write-Host ""
    Write-Host "Running source scrapers..."
    & $PythonCommand @RunArgs
    $RunExit = $LASTEXITCODE

    Write-Host ""
    Write-Host "Validating source outputs..."
    $ValidateArgs = @()
    $ValidateArgs += $PythonPrefixArgs
    $ValidateArgs += "validate_outputs.py"
    $ValidateArgs += "--project-root"
    $ValidateArgs += $ProjectRoot
    $ValidateArgs += "--registry"
    $ValidateArgs += $RegistryPath

    & $PythonCommand @ValidateArgs
    $ValidateExit = $LASTEXITCODE

    Write-Host ""
    Write-Host "Run exit code: $RunExit"
    Write-Host "Validation exit code: $ValidateExit"

    if ($RunExit -ne 0 -or $ValidateExit -ne 0) {
        throw "Daily run completed with errors. Check 05_SystemLogs for details."
    }

    Write-Host "Daily run completed successfully."
}
finally {
    Stop-Transcript
}
