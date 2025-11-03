#' Auditar una regla específica del plan de limpieza
#'
#' Extrae y muestra la información detallada de una o varias reglas del plan
#' de limpieza, a partir del objeto de evaluación generado por
#' [evaluar_consistencia()].
#'
#' @param ev Objeto de evaluación de consistencia (devuelto por [evaluar_consistencia()]).
#' @param ids Vector de ID(s) de regla(s) a auditar (p. ej. `"INSE_025"`).
#' @param verbose Lógico; si `TRUE` (por defecto) imprime en consola el resumen,
#' la expresión y los casos. Si `FALSE`, devuelve solo la lista.
#'
#' @return Una lista con tres elementos:
#' \describe{
#'   \item{resumen}{Tibble con id_regla, nombre_regla, n_inconsistencias y porcentaje.}
#'   \item{expresion}{Vector de texto con las expresiones R de procesamiento.}
#'   \item{casos}{Data frame con los casos que incumplen la regla.}
#' }
#' @export
#' @examples
#' \dontrun{
#' auditar_regla(ev, ids = "INSE_025")
#' }
auditar_regla <- function(ev, ids, verbose = TRUE) {

  # ---- 1. Obtener reporte ----
  aud <- prosecnur::reporte_bloques(ev, ids = ids)

  # ---- 2. Extraer partes ----
  resumen <- aud %>%
    dplyr::select(id_regla, nombre_regla, n_inconsistencias, porcentaje)

  expresion <- aud %>% dplyr::pull(procesamiento)
  casos <- aud$casos

  # ---- 3. Mostrar si se desea ----
  if (verbose) {
    cat("\nResumen de auditoría:\n")
    print(resumen)

    cat("\nExpresión R:\n")
    print(expresion)

    cat("\nCasos con inconsistencias:\n")
    print(casos)
  }

  # ---- 4. Retornar ----
  invisible(list(
    resumen   = resumen,
    expresion = expresion,
    casos     = casos
  ))
}
