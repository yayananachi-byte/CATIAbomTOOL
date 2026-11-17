param()

$ErrorActionPreference = 'Stop'
$installerRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $installerRoot
$buildPython = Join-Path $projectRoot 'CATIA_BOM_GUI\.venv\Scripts\python.exe'
$appSpec = Join-Path $installerRoot 'YR_CATIA_BOM_Manager.spec'
$setupSpec = Join-Path $installerRoot 'YR_CATIA_BOM_Setup.spec'
$requirements = Join-Path $installerRoot 'requirements-build.txt'
$setupExe = Join-Path $installerRoot 'dist\YR_CATIA_BOM_Setup.exe'

if (-not (Test-Path -LiteralPath $buildPython)) {
    throw "Build Python was not found: $buildPython. Run CATIA_BOM_GUI\SETUP_GUI_RUNTIME.cmd first."
}

Push-Location $installerRoot
try {
    Write-Host 'Installing/updating the local build tool...'
    & $buildPython -m pip install -r $requirements
    if ($LASTEXITCODE -ne 0) { throw 'Unable to install PyInstaller.' }

    Write-Host 'Building the standalone application...'
    & $buildPython -m PyInstaller --noconfirm --clean $appSpec
    if ($LASTEXITCODE -ne 0) { throw 'Application build failed.' }

    Write-Host 'Checking the packaged application runtime...'
    $appBundle = Join-Path $installerRoot 'dist\YR_CATIA_BOM_Manager'
    $appExe = Join-Path $appBundle 'YR_CATIA_BOM_Manager.exe'
    $checkLog = Join-Path $installerRoot 'build\packaged-self-check.log'
    $checkMarker = Join-Path $appBundle 'ZYRROOY'
    $createdMarker = -not (Test-Path -LiteralPath $checkMarker)
    try {
        if ($createdMarker) { New-Item -ItemType Directory -Path $checkMarker | Out-Null }
        $checkProcess = Start-Process -FilePath $appExe -ArgumentList @('--self-check', '--self-check-log', ('"' + $checkLog + '"')) -WorkingDirectory $appBundle -WindowStyle Hidden -PassThru
        if (-not $checkProcess.WaitForExit(60000)) {
            $checkProcess.Kill()
            throw 'Packaged application self-check timed out.'
        }
        $checkProcess.Refresh()
        if ($checkProcess.ExitCode -ne 0) {
            $details = if (Test-Path -LiteralPath $checkLog) { Get-Content -LiteralPath $checkLog -Raw } else { 'No diagnostic log was produced.' }
            throw "Packaged application self-check failed: $details"
        }
    }
    finally {
        # 仅移除本次创建的空授权标记；不将它打包进安装载荷。
        if ($createdMarker -and (Test-Path -LiteralPath $checkMarker)) {
            [IO.Directory]::Delete($checkMarker, $false)
        }
    }

    Write-Host 'Building the single-file offline installer...'
    & $buildPython -m PyInstaller --noconfirm --clean $setupSpec
    if ($LASTEXITCODE -ne 0) { throw 'Installer build failed.' }

    if (-not (Test-Path -LiteralPath $setupExe)) {
        throw "Installer output was not created: $setupExe"
    }

    $hash = Get-FileHash -LiteralPath $setupExe -Algorithm SHA256
    Write-Host ''
    Write-Host 'Build completed.'
    Write-Host "Installer: $setupExe"
    Write-Host "SHA256: $($hash.Hash)"
}
finally {
    Pop-Location
}
