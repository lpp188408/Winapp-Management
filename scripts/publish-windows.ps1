param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$project = Join-Path $root "src\WinappManagement\WinappManagement.csproj"
$output = Join-Path $root "dist\$Runtime"
$exe = Join-Path $output "WinappManagement.exe"

if (-not (Test-Path $project)) {
    throw "Project file not found: $project"
}

Write-Host "Project: $project"
Write-Host "Output: $output"

dotnet publish $project `
    -c $Configuration `
    -r $Runtime `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:PublishReadyToRun=false `
    -p:DebugType=none `
    -p:DebugSymbols=false `
    -o $output

if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed with exit code $LASTEXITCODE"
}

if (-not (Test-Path $exe)) {
    throw "Publish completed but exe was not found: $exe"
}

Write-Host "Published to $output"
Write-Host "Exe: $exe"
