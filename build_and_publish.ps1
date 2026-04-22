[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "`n==> $Message" -ForegroundColor Cyan
}

function Run-Command {
    param(
        [string]$Command,
        [string[]]$Arguments
    )

    & $Command @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Falha ao executar: $Command $($Arguments -join ' ')"
    }
}

$repoRoot = Split-Path -Parent $PSCommandPath
Set-Location $repoRoot

Write-Step "Validando pre-requisitos"

$requiredFiles = @(
    "main.py",
    "update_config.json",
    "requirements-dev.txt"
)

foreach ($file in $requiredFiles) {
    if (-not (Test-Path $file)) {
        throw "Arquivo obrigatorio nao encontrado: $file"
    }
}

$pythonExe = Join-Path $repoRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $pythonExe)) {
    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($null -eq $pythonCmd) {
        throw "Python nao encontrado. Crie o ambiente virtual em .venv ou instale o Python no PATH."
    }
    $pythonExe = $pythonCmd.Source
}

Write-Host "Python usado: $pythonExe" -ForegroundColor DarkGray

Write-Step "Instalando dependencias de build"
Run-Command $pythonExe @("-m", "pip", "install", "--disable-pip-version-check", "-r", "requirements-dev.txt")

Write-Step "Gerando executavel GbValidacoes.exe"
Run-Command $pythonExe @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--windowed",
    "--onefile",
    "--name", "GbValidacoes",
    "--collect-data", "customtkinter",
    "main.py"
)

Write-Step "Preparando artefatos de distribuicao"
if (-not (Test-Path "dist")) {
    New-Item -ItemType Directory -Path "dist" | Out-Null
}

if (Test-Path "dist\AtualizadorValidacao.exe") {
    Remove-Item "dist\AtualizadorValidacao.exe" -Force
}
if (Test-Path "dist\AtualizadorValidacao.zip") {
    Remove-Item "dist\AtualizadorValidacao.zip" -Force
}

Copy-Item "update_config.json" "dist\update_config.json" -Force
Compress-Archive -Path "dist\GbValidacoes.exe", "dist\update_config.json" -DestinationPath "dist\GbValidacoes.zip" -Force

$distArtifacts = @(
    "dist\GbValidacoes.exe",
    "dist\GbValidacoes.zip",
    "dist\update_config.json"
)

foreach ($artifact in $distArtifacts) {
    if (-not (Test-Path $artifact)) {
        throw "Artefato esperado nao encontrado: $artifact"
    }
}

Write-Step "Resumo dos artefatos"
Get-Item $distArtifacts | Select-Object Name, Length, LastWriteTime | Format-Table -AutoSize

Write-Step "Build concluido"
Write-Host "Artefatos atualizados com sucesso. Commit/push ficam por sua conta." -ForegroundColor Green
Write-Host ""
Write-Host "Quando quiser publicar, rode manualmente:" -ForegroundColor DarkGray
Write-Host "  git add -A" -ForegroundColor DarkGray
Write-Host "  git commit -m \"sua mensagem\"" -ForegroundColor DarkGray
Write-Host "  git push" -ForegroundColor DarkGray
