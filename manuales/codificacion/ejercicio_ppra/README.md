# Ejercicio Integrado PPRA (para alumnos)

Contexto general del paquete (para ubicar etapa):

- `manuales/codificacion/mapa_etapas_ppra.md`  -> "Estamos aqui: Codificacion"

Este ejercicio incluye todos los insumos dentro de `manuales/codificacion/ejercicio_ppra` para practicar:

- `select_multiple`
- `select_one` en modo **padre**
- `select_one` en modo **hijo**
- `text`
- `integer`

Caso de práctica (inspirado en terreno):

- Hogares que consultan por servicios, documentación y protección.
- Se registra:
  - `servicios` (select_multiple)
  - `motivo` (select_one padre)
  - `satisfaccion` (select_one hijo, con especificacion en `razon_baja_txt` cuando es baja)
  - `necesidad_txt` (text)
  - `edad` (integer)

## Estructura

- `insumos/instrumento_ejercicio.xlsx`
- `insumos/datos_ejercicio.xlsx`
- `insumos/familias_ejercicio.xlsx`
- `insumos/familias_ejercicio_resuelta.xlsx`
- `scripts/00_crear_insumos.R`
- `scripts/01_flujo_ejercicio.R`
- `ejercicio_ppra_resuelto.qmd`
- `ejercicio_ppra_template.qmd`
- `salidas/` (archivos generados al correr el flujo)

Notas importantes sobre FAMILIAS:

- `familias_ejercicio.xlsx` se genera con la funcion del paquete:
  `escribir_plantilla_familias()`. Esta es la base original.
- `familias_ejercicio_resuelta.xlsx` es una copia docente de esa base, con
  `use`, `modo_so`, `other_dummy_col` y `text_col` ya completados.

Carga del paquete:

- Estos materiales NO usan `load_all()` local.
- Se instala/carga `prosecnur` desde GitHub con:
  - `PROSECNUR_GITHUB_REPO` (default: `gonzaloalmendariz/prosecnur`)
  - `PROSECNUR_GITHUB_REF` (default: `main`)

## Flujo sugerido para clase

1. Revisar `familias_ejercicio.xlsx`.
2. Editar solo columnas clave en `familias`: `use`, `modo_so`, `other_dummy_col`, `text_col`.
3. Ejecutar `scripts/01_flujo_ejercicio.R` (por defecto usa la version resuelta).
4. Para probar su propia version de familias:
  - Exportar variable de entorno `PPRA_FAMILIAS_PATH` apuntando a su archivo.
  - Volver a correr `scripts/01_flujo_ejercicio.R`.
5. Revisar en `salidas/`:
  - `plantilla_codificacion_ejercicio.xlsx`
  - `datos_adaptados_ejercicio.xlsx`
  - `instrumento_adaptado_ejercicio.xlsx`

## Objetivo pedagogico

Que el alumno entienda de punta a punta:

- como definir familias,
- como se construye la plantilla de codificacion,
- y como se aplican cambios a datos e instrumento.
