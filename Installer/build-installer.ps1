# Builds bin\Installer\TestRunViewerSqlite-<version>-x64.msi
# Requires the WiX tool:  dotnet tool install --global wix --version 5.0.2
$ErrorActionPreference = 'Stop'

$root = Split-Path $PSScriptRoot
$csproj = Join-Path $root 'TestRunViewerSqlite.csproj'
$version = ([xml](Get-Content $csproj)).Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
$publishDir = Join-Path $root 'bin\Publish\win-x64'
$msi = Join-Path $root "bin\Installer\TestRunViewerSqlite-$version-x64.msi"

# Self-contained so target PCs don't need the .NET runtime installed
if (Test-Path $publishDir) { Remove-Item $publishDir -Recurse -Force }
dotnet publish $csproj -c Release -r win-x64 --self-contained true -o $publishDir
if ($LASTEXITCODE) { throw 'dotnet publish failed' }

wix extension add -g WixToolset.UI.wixext/5.0.2
if ($LASTEXITCODE) { throw 'wix extension add failed' }

wix build (Join-Path $PSScriptRoot 'Package.wxs') -arch x64 -ext WixToolset.UI.wixext `
    -d "PublishDir=$publishDir" -d "Version=$version" -o $msi
if ($LASTEXITCODE) { throw 'wix build failed' }

Write-Host "Built $msi"
