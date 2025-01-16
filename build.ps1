# Set error action preference to stop on any error
$ErrorActionPreference = "Stop"

# Project path - adjust if script is not in the root directory
$projectPath = ".\OstToPst\OstToPst.csproj"

Write-Host "🔄 Restoring NuGet packages..." -ForegroundColor Cyan
dotnet restore $projectPath

Write-Host "🏗️ Building project..." -ForegroundColor Cyan
dotnet build $projectPath --configuration Release

Write-Host "🚀 Running application..." -ForegroundColor Green
dotnet run --project $projectPath --configuration Release
