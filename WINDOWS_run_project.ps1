# 1. Check/Install Python (Global)
Write-Host "--- [1/4] Checking Python ---" -ForegroundColor Cyan
if (!(Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "Python not found. Installing via winget..." -ForegroundColor Yellow
    winget install -e --id Python.Python.3 --silent --accept-package-agreements --accept-source-agreements
    # Update current session Path
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
} else {
    Write-Host "Python is ready." -ForegroundColor Green
}

# 2. Setup Virtual Environment
$venvPath = ".\.venv"
$venvPython = "$venvPath\Scripts\python.exe"

Write-Host "`n--- [2/4] Virtual Environment ---" -ForegroundColor Cyan
if (!(Test-Path $venvPath)) {
    Write-Host "Creating .venv..." -ForegroundColor Yellow
    python -m venv .venv
} else {
    Write-Host "Environment already exists." -ForegroundColor Green
}

# 3. Verify Libraries inside Venv
$libraries = @("pandas", "openpyxl")
Write-Host "`n--- [3/4] Checking Libraries ---" -ForegroundColor Cyan

# Upgrade pip inside venv first
& $venvPython -m pip install --upgrade pip --quiet

foreach ($lib in $libraries) {
    & $venvPython -c "import $lib" 2> $null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Installing $lib..." -ForegroundColor Yellow
        & $venvPython -m pip install $lib --quiet
    } else {
        Write-Host "Library '$lib' is verified." -ForegroundColor Green
    }
}

# 4. Run main.py
Write-Host "`n--- [4/4] Executing Project ---" -ForegroundColor Cyan
if (Test-Path ".\main.py") {
    Write-Host "Running main.py..." -ForegroundColor Magenta
    & $venvPython .\main.py
} else {
    Write-Host "Warning: 'main.py' not found in the current directory." -ForegroundColor Red
    Write-Host "Creating a starter 'main.py' for you..." -ForegroundColor Gray
    @"
import pandas as pd
print("Environment Check Successful!")
print(f"Pandas version: {pd.__version__}")
# Create a dummy dataframe to test openpyxl
df = pd.DataFrame({'Data': [10, 20, 30]})
# df.to_excel('test.xlsx') # Uncomment to test openpyxl
"@ | Out-File -FilePath ".\main.py"
    Write-Host "Starter 'main.py' created. Run the script again to execute it." -ForegroundColor Yellow
}

Write-Host "`nWorkflow Complete." -ForegroundColor Green