param(
    [ValidateSet("Debug", "Final")]
    [string]$Modo = "Debug"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$Raiz = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Raiz

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " RPA CBN - GENERADOR DE EJECUTABLE PORTABLE" -ForegroundColor Cyan
Write-Host " Modo: $Modo" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$ArchivosRequeridos = @(
    ".\app_gui.py",
    ".\assets\ransa_olive.json",
    ".\assets\rpa_cbn.ico"
)

foreach ($Archivo in $ArchivosRequeridos) {
    if (-not (Test-Path $Archivo)) {
        throw "Falta el archivo requerido: $Archivo"
    }
}

Write-Host "[1/7] Verificando Python..." -ForegroundColor Yellow
python --version
if ($LASTEXITCODE -ne 0) {
    throw "No se pudo ejecutar Python desde el entorno actual."
}

Write-Host "[2/7] Verificando librerías..." -ForegroundColor Yellow
python -c "import PyInstaller, customtkinter, playwright, PIL; print('Dependencias de compilación: OK')"
if ($LASTEXITCODE -ne 0) {
    throw "Faltan dependencias. Ejecuta: python -m pip install -r .\requirements_gui.txt"
}

if ($Modo -eq "Debug") {
    $NombreApp = "RPA_CBN_DEBUG"
    $ModoVentana = "--console"
}
else {
    $NombreApp = "RPA_CBN"
    $ModoVentana = "--windowed"
}

Write-Host "[3/7] Limpiando compilaciones anteriores..." -ForegroundColor Yellow
Remove-Item ".\build\$NombreApp" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item ".\dist\$NombreApp" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item ".\$NombreApp.spec" -Force -ErrorAction SilentlyContinue
Remove-Item ".\dist\${NombreApp}_Portable.zip" -Force -ErrorAction SilentlyContinue

Write-Host "[4/7] Compilando $NombreApp..." -ForegroundColor Yellow

$Argumentos = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--onedir",
    $ModoVentana,
    "--name", $NombreApp,
    "--icon", ".\assets\rpa_cbn.ico",
    "--paths", ".",
    "--collect-all", "customtkinter",
    "--collect-all", "playwright",
    ".\app_gui.py"
)

& python @Argumentos

if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller terminó con código de error $LASTEXITCODE."
}

$CarpetaSalida = Join-Path $Raiz "dist\$NombreApp"
$Ejecutable = Join-Path $CarpetaSalida "$NombreApp.exe"

if (-not (Test-Path $Ejecutable)) {
    throw "No se encontró el ejecutable generado: $Ejecutable"
}

Write-Host "[5/7] Copiando recursos portables..." -ForegroundColor Yellow
Copy-Item ".\assets" (Join-Path $CarpetaSalida "assets") -Recurse -Force

foreach ($Carpeta in @("data", "configuracion", "perfiles", "ejecuciones")) {
    New-Item -ItemType Directory -Path (Join-Path $CarpetaSalida $Carpeta) -Force | Out-Null
}

$TextoInicio = @"
RPA CBN - PAQUETE PORTABLE

1. No instalar esta carpeta en Program Files.
2. Descomprimirla en Documentos, Escritorio u otra carpeta del usuario.
3. Ejecutar $NombreApp.exe.
4. La primera vez, iniciar sesión manualmente en CBN.
5. Edge o Chrome debe estar instalado.
6. No mover solamente el EXE: se debe conservar toda esta carpeta.
"@

Set-Content -Path (Join-Path $CarpetaSalida "LEEME_PRIMERO.txt") -Value $TextoInicio -Encoding UTF8

Write-Host "[6/7] Comprimiendo paquete portable..." -ForegroundColor Yellow
$ZipSalida = Join-Path $Raiz "dist\${NombreApp}_Portable.zip"

Compress-Archive `
    -Path (Join-Path $CarpetaSalida "*") `
    -DestinationPath $ZipSalida `
    -CompressionLevel Optimal `
    -Force

Write-Host "[7/7] Compilación terminada." -ForegroundColor Green
Write-Host ""
Write-Host "Ejecutable:" -ForegroundColor Green
Write-Host "  $Ejecutable"
Write-Host ""
Write-Host "Paquete para otra laptop:" -ForegroundColor Green
Write-Host "  $ZipSalida"
Write-Host ""
Write-Host "Prueba local:" -ForegroundColor Green
Write-Host "  & `"$Ejecutable`""
Write-Host ""
