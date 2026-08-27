$ErrorActionPreference = "Stop"

Write-Host "=== GIA v0.4.9 - Día 1 / preparación Windows ===" -ForegroundColor Cyan

if (-not (Get-Command py -ErrorAction SilentlyContinue)) {
    throw "No se encontró el lanzador 'py' de Python. Instala Python 3.11+ y vuelve a ejecutar este script."
}

py -m venv .venv
& .\.venv\Scripts\python.exe -m pip install --upgrade pip
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt

if (-not (Test-Path .env)) {
    Copy-Item .env.example .env
    Write-Host "Se creó .env desde .env.example. Edítalo antes del smoke test." -ForegroundColor Yellow
}

& .\.venv\Scripts\python.exe scripts\day1_preflight.py
Write-Host "=== Preflight PASS ===" -ForegroundColor Green
Write-Host "Ahora configura SECRET_KEY en .env y ejecuta:"
Write-Host "  .\\.venv\\Scripts\\python.exe scripts\\day1_smoke.py"
Write-Host "Luego:"
Write-Host "  .\\.venv\\Scripts\\python.exe -m flask --app app:create_app db upgrade"
Write-Host "Y finalmente:"
Write-Host "  .\\.venv\\Scripts\\python.exe app.py"
