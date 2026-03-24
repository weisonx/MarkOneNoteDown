param(
    [ValidateSet("clean", "build", "run", "all")]
    [string]$Action = "all"
)

$ErrorActionPreference = "Stop"

function Invoke-Clean {
    Write-Host "Cleaning bin/obj..."
    Get-ChildItem -Path "src" -Directory -Recurse -Force |
        Where-Object { $_.Name -in @("bin", "obj") } |
        ForEach-Object { Remove-Item -Recurse -Force $_.FullName }
}

function Invoke-Build {
    Write-Host "Restoring packages..."
    dotnet restore
    Write-Host "Building solution..."
    dotnet build MarkOneNoteDown.sln -c Debug
}

function Invoke-Run {
    Write-Host "Running app..."
    dotnet run --project src\MarkOneNoteDown.App\MarkOneNoteDown.App.csproj -c Debug
}

switch ($Action) {
    "clean" { Invoke-Clean }
    "build" { Invoke-Build }
    "run" { Invoke-Run }
    "all" {
        Invoke-Clean
        Invoke-Build
        Invoke-Run
    }
}
