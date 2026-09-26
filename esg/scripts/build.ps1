param([switch]$SkipInstall,[switch]$SkipTests)
$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (-not (Get-Command node -ErrorAction SilentlyContinue)) { throw 'Install Node.js 22.13+ and npm first.' }
if (-not $env:JAVA_HOME) { throw 'Set JAVA_HOME to your JDK 17 installation.' }
Push-Location (Join-Path $projectRoot 'web')
try {
    if (-not $SkipInstall) { & npm.cmd ci; if ($LASTEXITCODE) { throw 'Dependency installation failed.' } }
    if (-not $SkipTests) {
        & npm.cmd run typecheck; if ($LASTEXITCODE) { throw 'Type checking failed.' }
        & npm.cmd test; if ($LASTEXITCODE) { throw 'Application tests failed.' }
    }
    & npm.cmd run build; if ($LASTEXITCODE) { throw 'Web build failed.' }
} finally { Pop-Location }
& node (Join-Path $PSScriptRoot 'bundle-web.mjs')
if ($LASTEXITCODE) { throw 'Bundling assets failed.' }
Push-Location (Join-Path $projectRoot 'android')
try {
    & .\gradlew.bat assembleRelease lintRelease --console plain
    if ($LASTEXITCODE) { throw 'Android build or release checks failed.' }
} finally { Pop-Location }
Write-Host 'Unsigned APK: android/app/build/outputs/apk/release/app-release-unsigned.apk'
Write-Host 'Use your own release signing key before distribution. See docs/INSTALLATION_BACKUP_AND_RELEASE.md.'
