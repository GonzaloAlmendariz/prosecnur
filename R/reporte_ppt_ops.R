# =============================================================================
# reporte_ppt_ops()
# -----------------------------------------------------------------------------
# Wrapper de reporte_ppt() que usa por defecto la plantilla OPS
# =============================================================================

#' Generar un PPT de frecuencias simples usando la plantilla OPS
#'
#' `reporte_ppt_ops()` es un envoltorio ligero sobre [`reporte_ppt()`] que
#' utiliza por defecto una plantilla de PowerPoint específica para OPS
#' (`plantilla_OPS.pptx`), manteniendo la misma lógica de graficación y
#' recorrido por secciones que el reporte estándar de frecuencias simples.
#'
#' En la práctica:
#' \itemize{
#'   \item Llama internamente a `reporte_ppt()`.
#'   \item Si \code{template_pptx} es \code{NULL}, busca el archivo
#'         \code{"plantillas/plantilla_OPS.pptx"} dentro del paquete
#'         (carpeta \code{inst/plantillas/}).
#'   \item Si se proporciona \code{template_pptx}, usa esa ruta en su lugar.
#' }
#'
#' Esto permite mantener intacta la implementación de `reporte_ppt()` y, a la vez,
#' trabajar con una plantilla específica para un estudio (OPS) o para un cliente.
#'
#' @inheritParams reporte_ppt
#' @param path_pptx Ruta del archivo PPTX a generar. Por defecto
#'   \code{"reporte_ppt_ops.pptx"}.
#' @param template_pptx Ruta a una plantilla PPTX. Si es \code{NULL}, se busca
#'   una plantilla interna llamada \code{"plantillas/plantilla_OPS.pptx"} dentro
#'   del paquete.
#'
#' @return Lo mismo que devuelve `reporte_ppt()`: típicamente una lista con
#'   gráficos y el log de decisiones, de forma invisible si solo se genera
#'   el archivo PPTX.
#'
#' @export
reporte_ppt_ops <- function(
    data,
    instrumento      = NULL,
    secciones        = NULL,
    path_pptx        = "reporte_ppt_ops.pptx",
    template_pptx    = NULL,
    ...
) {

  # ---------------------------------------------------------------------------
  # 1. Resolver plantilla OPS
  # ---------------------------------------------------------------------------
  if (is.null(template_pptx)) {
    template_ops <- system.file(
      "plantillas/plantilla_OPS.pptx",
      package = "prosecnur"
    )

    if (!nzchar(template_ops) || !file.exists(template_ops)) {
      stop(
        "No se encontró `plantillas/plantilla_OPS.pptx` dentro del paquete. ",
        "Crea el archivo y colócalo en `inst/plantillas/` antes de usar ",
        "`reporte_ppt_ops()`, o pasa una ruta explícita en `template_pptx`.",
        call. = FALSE
      )
    }
  } else {
    if (!file.exists(template_pptx)) {
      stop(
        "No se encontró el archivo de plantilla especificado en `template_pptx`: ",
        template_pptx,
        call. = FALSE
      )
    }
    template_ops <- template_pptx
  }

  # ---------------------------------------------------------------------------
  # 2. Delegar en reporte_ppt() usando la plantilla OPS
  # ---------------------------------------------------------------------------
  reporte_ppt(
    data          = data,
    instrumento   = instrumento,
    secciones     = secciones,
    path_pptx     = path_pptx,
    template_pptx = template_ops,
    ...
  )
}
