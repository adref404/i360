# Version para Windows (PowerShell) de subir.sh: sube este mismo codigo a la app de PRUEBA o a la REAL.
#   .\subir.ps1 configurar <ID del script de prueba>   (una sola vez)
#   .\subir.ps1 prueba                                  sube a PRUEBA
#   .\subir.ps1 real                                    sube a la REAL (pide confirmacion)
# .clasp.json siempre queda apuntando a la REAL; para PRUEBA se cambia solo mientras dura la subida.
param([string]$accion = '', [string]$id = '')
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

function IdDe($archivo) {
  $m = [regex]::Match((Get-Content $archivo -Raw), '"scriptId"\s*:\s*"([^"]*)"')
  if ($m.Success) { $m.Groups[1].Value } else { '' }
}

# Si una subida a prueba se corto a la mitad, dejar .clasp.json otra vez en la REAL
if (Test-Path .clasp.real.bak) { Move-Item -Force .clasp.real.bak .clasp.json }

switch ($accion) {
  'configurar' {
    if (-not $id) { Write-Host 'Uso: .\subir.ps1 configurar <ID del script de prueba>'; exit 1 }
    Set-Content -Encoding ascii .clasp.prueba.json "{ `"scriptId`": `"$id`" }"
    Write-Host "Listo: la app de PRUEBA es el script $id"
  }
  'prueba' {
    if (-not (Test-Path .clasp.prueba.json)) { Write-Host 'Primero ejecuta: .\subir.ps1 configurar <ID del script de prueba>'; exit 1 }
    $prueba = IdDe .clasp.prueba.json; $real = IdDe .clasp.json
    if (-not $prueba -or $prueba -eq $real) { Write-Host 'El ID de prueba esta vacio o es el mismo que el real. Revisa .clasp.prueba.json'; exit 1 }
    Copy-Item .clasp.json .clasp.real.bak
    try {
      $txt = (Get-Content .clasp.json -Raw) -replace '"scriptId"\s*:\s*"[^"]*"', "`"scriptId`": `"$prueba`""
      [IO.File]::WriteAllText((Join-Path $PSScriptRoot '.clasp.json'), $txt)
      Write-Host 'Subiendo a PRUEBA...'
      clasp push
      if ($LASTEXITCODE -ne 0) { throw 'clasp push fallo' }
      Write-Host 'Listo. Abre la app de prueba con su enlace /dev (Implementar > Probar implementaciones).'
    } finally { Move-Item -Force .clasp.real.bak .clasp.json }
  }
  'real' {
    $r = Read-Host 'Vas a subir a la app REAL que usan los instructores. Seguro? (s/n)'
    if ($r -ne 's') { Write-Host 'Cancelado.'; exit 1 }
    clasp push
    if ($LASTEXITCODE -ne 0) { exit 1 }
    Write-Host 'Listo. Ahora en Apps Script: Implementar > Gestionar implementaciones > Nueva version.'
  }
  default {
    Get-Content $PSCommandPath -TotalCount 5 | Select-Object -Skip 1 | ForEach-Object { $_ -replace '^# ?', '' }
    exit 1
  }
}
