$ErrorActionPreference = "Stop"

$projectRoot = $PSScriptRoot
$workspaceRoot = Split-Path -Parent $projectRoot
$pythonExe = Join-Path $workspaceRoot ".venv-1\Scripts\python.exe"
$mainPy = Join-Path $projectRoot "main.py"

if (-not (Test-Path $mainPy)) {
    Write-Error "main.py not found at $mainPy"
}

if (Test-Path $pythonExe) {
    Start-Process -FilePath $pythonExe -ArgumentList "`"$mainPy`"" -WorkingDirectory $projectRoot -WindowStyle Hidden
} else {
    # Fallback if venv path changed.
    Start-Process -FilePath "python" -ArgumentList "`"$mainPy`"" -WorkingDirectory $projectRoot -WindowStyle Hidden
}
