#!/usr/bin/env bash
# Sube este mismo código a la app de PRUEBA o a la REAL (un solo código, dos destinos).
#   ./subir.sh configurar <ID del script de prueba>   (una sola vez)
#   ./subir.sh prueba                                  sube a PRUEBA
#   ./subir.sh real                                    sube a la REAL (pide confirmación)
# .clasp.json siempre queda apuntando a la REAL; para PRUEBA se cambia solo mientras dura la subida.
set -euo pipefail
cd "$(dirname "$0")"

# Si una subida a prueba se cortó a la mitad, dejar .clasp.json otra vez en la REAL
[ -f .clasp.real.bak ] && mv -f .clasp.real.bak .clasp.json

case "${1:-}" in
  configurar)
    id="${2:-}"
    [ -n "$id" ] || { echo "Uso: ./subir.sh configurar <ID del script de prueba>"; exit 1; }
    printf '{ "scriptId": "%s" }\n' "$id" > .clasp.prueba.json
    echo "Listo: la app de PRUEBA es el script $id"
    ;;
  prueba)
    [ -f .clasp.prueba.json ] || { echo "Primero ejecuta: ./subir.sh configurar <ID del script de prueba>"; exit 1; }
    id="$(sed -n 's/.*"scriptId" *: *"\([^"]*\)".*/\1/p' .clasp.prueba.json)"
    real="$(sed -n 's/.*"scriptId" *: *"\([^"]*\)".*/\1/p' .clasp.json)"
    [ -n "$id" ] && [ "$id" != "$real" ] || { echo "El ID de prueba está vacío o es el mismo que el real. Revisa .clasp.prueba.json"; exit 1; }
    cp .clasp.json .clasp.real.bak
    trap 'mv -f .clasp.real.bak .clasp.json' EXIT
    sed -i "s/\"scriptId\": *\"[^\"]*\"/\"scriptId\": \"$id\"/" .clasp.json
    echo "Subiendo a PRUEBA…"
    clasp push
    echo "Listo. Abre la app de prueba con su enlace /dev (Implementar › Probar implementaciones)."
    ;;
  real)
    read -r -p "Vas a subir a la app REAL que usan los instructores. ¿Seguro? (s/n) " r
    [ "$r" = "s" ] || { echo "Cancelado."; exit 1; }
    clasp push
    echo "Listo. Ahora en Apps Script: Implementar › Gestionar implementaciones › Nueva versión."
    ;;
  *)
    sed -n '2,6p' "$0" | sed 's/^# \{0,1\}//'
    exit 1
    ;;
esac
