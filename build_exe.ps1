Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# build_exe.ps1
# Baut aus SensorMeasureTrace (SMT) eine einzelne,
# eigenstaendige EXE-Datei
# (kein Python/keine Installation auf dem Zielrechner noetig).
# ---------------------------------------------------------------------------

$root = $PSScriptRoot
Set-Location $root

$logoPng = Join-Path $root "resources\logo.png"
$logoIco = Join-Path $root "resources\logo.ico"
$distDir = Join-Path $root "dist"
$buildDir = Join-Path $root "build"
$newExe = Join-Path $distDir "SensorMeasureTrace.exe"
$oldExe = Join-Path $distDir "SensorCalFinder.exe"
$olderExe = Join-Path $distDir "LeepaOrderCompMapFinder.exe"
$preferredPython = "C:\Users\ghk2si\.conda\envs\bosch-sensor-analysis\python.exe"
$condaEnvRoot = Split-Path $preferredPython -Parent
$condaLibBin = Join-Path $condaEnvRoot "Library\bin"
$tclDll = Join-Path $condaEnvRoot "Library\bin\tcl86t.dll"
$tkDll = Join-Path $condaEnvRoot "Library\bin\tk86t.dll"
$tclLibDir = Join-Path $condaEnvRoot "Library\lib\tcl8.6"
$tkLibDir = Join-Path $condaEnvRoot "Library\lib\tk8.6"

if (-not (Test-Path $logoPng)) {
    Write-Warning "resources\logo.png nicht gefunden - die EXE wird ohne Logo/Icon gebaut."
} elseif (-not (Test-Path $logoIco)) {
    Write-Host "Erzeuge logo.ico aus logo.png ..."
    python -c "from PIL import Image; img = Image.open(r'$logoPng'); img.save(r'$logoIco', sizes=[(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)])"
}

Write-Host "Baue EXE mit PyInstaller ..."

Get-Process SensorMeasureTrace -ErrorAction SilentlyContinue |
    Stop-Process -Force -ErrorAction SilentlyContinue

if (Test-Path $newExe) {
    Remove-Item $newExe -Force
}
if (Test-Path $oldExe) {
    Remove-Item $oldExe -Force
}
if (Test-Path $olderExe) {
    Remove-Item $olderExe -Force
}
if (Test-Path $buildDir) {
    Remove-Item $buildDir -Recurse -Force
}

# Hilft PyInstaller beim Aufloesen nativer Abhaengigkeiten aus der conda-Umgebung.
if (Test-Path $condaLibBin) {
    $env:PATH = "$condaLibBin;$env:PATH"
}

$pyInstallerCommand = $null
$venvBuildPython = Join-Path $root ".venv-build\Scripts\python.exe"
$slimBuildPython = Join-Path $root ".venv-smt-build\Scripts\python.exe"
if (Test-Path $venvBuildPython) {
    $pyInstallerCommand = @($venvBuildPython, "-m", "PyInstaller")
} elseif (Test-Path $slimBuildPython) {
    $pyInstallerCommand = @($slimBuildPython, "-m", "PyInstaller")
} else {
    $pyInstaller = Get-Command pyinstaller -ErrorAction SilentlyContinue
    if ($pyInstaller) {
        $pyInstallerCommand = @($pyInstaller.Source)
    } elseif (Test-Path $preferredPython) {
        $pyInstallerCommand = @($preferredPython, "-m", "PyInstaller")
    } else {
        $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
        if ($pythonCmd) {
            $pyInstallerCommand = @($pythonCmd.Source, "-m", "PyInstaller")
        } else {
            throw "PyInstaller nicht gefunden. Bitte conda-Umgebung aktivieren oder PyInstaller installieren."
        }
    }
}

$iconArg = @()
if (Test-Path $logoIco) {
    $iconArg = @("--icon", $logoIco)
}

$tkArgs = @()
if (Test-Path $tclDll) {
    $tkArgs += @("--add-binary", "$tclDll;.")
}
if (Test-Path $tkDll) {
    $tkArgs += @("--add-binary", "$tkDll;.")
}
if (Test-Path $tclLibDir) {
    $tkArgs += @("--add-data", "$tclLibDir;tcl8.6")
}
if (Test-Path $tkLibDir) {
    $tkArgs += @("--add-data", "$tkLibDir;tk8.6")
}

$dllCandidates = @(
    (Join-Path $condaEnvRoot "DLLs\sqlite3.dll"),
    (Join-Path $condaEnvRoot "DLLs\ffi.dll"),
    (Join-Path $condaLibBin "sqlite3.dll"),
    (Join-Path $condaLibBin "ffi.dll"),
    (Join-Path $condaLibBin "libcrypto-3-x64.dll"),
    (Join-Path $condaLibBin "libssl-3-x64.dll"),
    (Join-Path $condaLibBin "liblzma.dll"),
    (Join-Path $condaLibBin "LIBBZ2.dll"),
    (Join-Path $condaLibBin "libbz2.dll")
)

$extraDllArgs = @()
$seen = @{}
foreach ($dllPath in $dllCandidates) {
    if ((Test-Path $dllPath) -and -not $seen.ContainsKey($dllPath.ToLowerInvariant())) {
        $extraDllArgs += @("--add-binary", "$dllPath;.")
        $seen[$dllPath.ToLowerInvariant()] = $true
    }
}

# Die conda-Umgebung enthaelt viele Data-Science-Pakete (matplotlib, IPython,
# sphinx, babel, jedi, ...), die das Tool nicht braucht. Ohne
# --exclude-module bricht PyInstaller beim Modul-Scan mit einem
# RecursionError ab, weil der Abhaengigkeitsgraph dieser Pakete zu tief
# verschachtelt ist.
$excludeArgs = @(
    "matplotlib", "IPython", "jedi", "parso", "astroid", "docutils",
    "sphinx", "babel", "notebook", "jupyter", "jupyter_client", "jupyter_core",
    "tornado", "zmq", "nbconvert", "nbformat", "pytest", "scipy", "numexpr",
    "tables", "lib2to3", "pygments", "pandas.tests",
    "numpy.tests", "jinja2", "markupsafe", "matplotlib_inline", "pyarrow",
    "bottleneck", "numba", "llvmlite", "PyQt5", "PyQt6", "PySide2",
    "PySide6", "PIL.ImageQt"
) | ForEach-Object { "--exclude-module", $_ }

$pyInstallerSubArgs = @()
if ($pyInstallerCommand.Count -gt 1) {
    $pyInstallerSubArgs = $pyInstallerCommand[1..($pyInstallerCommand.Count - 1)]
}

& $pyInstallerCommand[0] @pyInstallerSubArgs `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name SensorMeasureTrace `
    --add-data "resources;resources" `
    @iconArg `
    @tkArgs `
    @extraDllArgs `
    @excludeArgs `
    main.py

Write-Host ""
if (Test-Path $newExe) {
    $versionMatch = Select-String -Path (Join-Path $root "gui.py") -Pattern '^APP_VERSION = "([^"]+)"'
    if ($versionMatch) {
        $version = $versionMatch.Matches[0].Groups[1].Value
        $versionedExe = Join-Path $distDir "SensorMeasureTrace_v$version.exe"
        Copy-Item $newExe $versionedExe -Force
        Write-Host "Versionierte Kopie: $versionedExe"
    }
    Write-Host "Fertig! Die EXE liegt unter: $newExe"
} else {
    Write-Warning "Build fehlgeschlagen - dist\SensorMeasureTrace.exe wurde nicht erzeugt. Bitte PyInstaller-Ausgabe oben pruefen."
}
