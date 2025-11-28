# =============================================================================
# graficar_dico()
# -----------------------------------------------------------------------------
# Pies dicotómicos (Sí / No) usando tabla agregada (una fila por indicador)
# =============================================================================

#' Graficar pies dicotómicos (Sí / No) desde una tabla agregada
#'
#' Genera gráficos de torta (pie charts) para uno o varios indicadores
#' dicotómicos a partir de una tabla agregada (una fila por indicador), donde
#' se dispone del porcentaje de respuesta afirmativa (Sí) y del N total.
#' Cada indicador se muestra en una faceta distinta.
#'
#' @param data Data frame en formato agregado: una fila por indicador.
#' @param var_indicador Nombre (string) de la columna que identifica cada
#'   indicador (por ejemplo, `"indicador"`).
#' @param var_porcentaje_si Nombre (string) de la columna con el porcentaje
#'   de respuesta afirmativa (Sí) para cada indicador.
#' @param var_n Nombre (string) de la columna con el N total por indicador.
#' @param escala_valor Escala de los porcentajes: \code{"proporcion_1"} si los
#'   valores están en 0–1, o \code{"proporcion_100"} si están en 0–100.
#'
#' @param etiqueta_si Etiqueta a mostrar para la categoría afirmativa
#'   (por defecto \code{"Sí"}).
#' @param etiqueta_no Etiqueta a mostrar para la categoría negativa
#'   (por defecto \code{"No"}). Se calcula como complemento de Sí.
#' @param colores_respuesta Vector de colores HEX con nombre para las
#'   categorías, usando como nombres \code{etiqueta_si} y \code{etiqueta_no}.
#'
#' @param invertir_respuestas Lógico; si \code{TRUE}, invierte el orden de las
#'   categorías (por ejemplo, muestra primero "No" y luego "Sí") tanto en el
#'   gráfico como en la leyenda.
#'
#' @param mostrar_valores Lógico; si \code{TRUE}, muestra el porcentaje dentro
#'   de cada sector de la torta.
#' @param decimales Número de decimales para los porcentajes.
#' @param umbral_etiqueta Mínima proporción (0–1) para mostrar la etiqueta
#'   de porcentaje (por ejemplo, 0.03 equivale a 3\%).
#'
#' @param incluir_n_en_titulo Lógico; si \code{TRUE}, concatena
#'   una línea adicional con \code{"N = ..."} al nombre del indicador en la
#'   faceta (usando un salto de línea \code{"\\n"}).
#' @param prefijo_n_titulo Texto añadido antes del valor de N cuando se usa
#'   \code{incluir_n_en_titulo} (por defecto \code{"N = "}).
#'
#' @param ncol_facetas Número de columnas en el facetado. Si es \code{NULL},
#'   se intenta un valor razonable según el número de indicadores.
#'
#' @param titulo Título general del gráfico.
#' @param subtitulo Subtítulo opcional.
#' @param nota_pie Nota o fuente para el pie de página (alineado a la derecha).
#'
#' @param color_titulo Color del título.
#' @param size_titulo Tamaño del título.
#' @param color_subtitulo Color del subtítulo.
#' @param size_subtitulo Tamaño del subtítulo.
#' @param color_nota_pie Color del texto del pie de página.
#' @param size_nota_pie Tamaño del texto del pie de página.
#' @param color_leyenda Color del texto de la leyenda.
#' @param size_leyenda Tamaño del texto de la leyenda.
#' @param color_texto_segmentos Color del texto de los porcentajes dentro de
#'   los sectores.
#' @param size_texto_segmentos Tamaño del texto de los porcentajes.
#' @param color_titulos_pie Color de los títulos de cada pie (strip de facetas).
#' @param size_titulos_pie Tamaño de los títulos de cada pie.
#'
#' @param mostrar_leyenda Lógico; si \code{FALSE}, oculta la leyenda.
#' @param posicion_leyenda Posición de la leyenda (por ejemplo,
#'   \code{"bottom"}, \code{"right"}, \code{"none"}).
#' @param invertir_leyenda Lógico; si \code{TRUE}, invierte el orden de los
#'   ítems en la leyenda (sin alterar el orden de las categorías en el gráfico).
#'
#' @param textos_negrita Vector de caracteres que indica qué elementos deben
#'   mostrarse en negrita. Puede incluir cualquiera de:
#'   \code{"titulo"}, \code{"porcentajes"}, \code{"leyenda"}, \code{"titulos_pie"}.
#'
#' @param exportar Método de salida: \code{"rplot"} (devuelve un objeto
#'   \code{ggplot}), \code{"png"} (exporta un archivo PNG) o \code{"ppt"}
#'   (exporta una diapositiva PPTX con el gráfico incrustado).
#' @param path_salida Ruta del archivo a crear cuando \code{exportar} no es
#'   \code{"rplot"}.
#' @param ancho Ancho del gráfico (cuando se exporta a archivo).
#' @param alto Alto del gráfico (cuando se exporta a archivo).
#' @param dpi Resolución en puntos por pulgada (solo para PNG).
#'
#' @return Un objeto \code{ggplot} si \code{exportar = "rplot"}. De forma
#'   invisible, el gráfico exportado si se utiliza \code{"png"} o \code{"ppt"}.
#' @export
graficar_dico <- function(
    data,
    var_indicador,
    var_porcentaje_si,
    var_n,
    escala_valor           = c("proporcion_1", "proporcion_100"),
    etiqueta_si            = "Sí",
    etiqueta_no            = "No",
    colores_respuesta      = NULL,
    invertir_respuestas    = FALSE,
    mostrar_valores        = TRUE,
    decimales              = 1,
    umbral_etiqueta        = 0.03,
    incluir_n_en_titulo    = FALSE,
    prefijo_n_titulo       = "N = ",
    ncol_facetas           = NULL,
    titulo                 = NULL,
    subtitulo              = NULL,
    nota_pie               = NULL,
    color_titulo           = "#000000",
    size_titulo            = 12,
    color_subtitulo        = "#000000",
    size_subtitulo         = 10,
    color_nota_pie         = "#000000",
    size_nota_pie          = 8,
    color_leyenda          = "#000000",
    size_leyenda           = 8,
    color_texto_segmentos  = "#FFFFFF",
    size_texto_segmentos   = 3,
    color_titulos_pie      = "#000000",
    size_titulos_pie       = 9,
    mostrar_leyenda        = TRUE,
    posicion_leyenda       = "bottom",
    invertir_leyenda       = FALSE,
    textos_negrita         = NULL,
    exportar               = c("rplot", "png", "ppt"),
    path_salida            = NULL,
    ancho                  = 10,
    alto                   = 6,
    dpi                    = 300
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y

  escala_valor   <- match.arg(escala_valor)
  exportar       <- match.arg(exportar)
  textos_negrita <- textos_negrita %||% character(0)

  # ---------------------------------------------------------------------------
  # 0. Validaciones
  # ---------------------------------------------------------------------------
  if (!var_indicador %in% names(data)) {
    stop("`var_indicador` no existe en `data`.", call. = FALSE)
  }
  if (!var_porcentaje_si %in% names(data)) {
    stop("`var_porcentaje_si` no existe en `data`.", call. = FALSE)
  }
  if (!var_n %in% names(data)) {
    stop("`var_n` no existe en `data`.", call. = FALSE)
  }

  df <- data

  # ---------------------------------------------------------------------------
  # 1. Construir tabla larga con Sí y No
  # ---------------------------------------------------------------------------
  df_proc <- df |>
    dplyr::select(
      dplyr::all_of(c(var_indicador, var_porcentaje_si, var_n))
    )

  if (!is.numeric(df_proc[[var_porcentaje_si]])) {
    stop("`var_porcentaje_si` debe ser numérica.", call. = FALSE)
  }

  if (escala_valor == "proporcion_100") {
    df_proc$prop_si <- df_proc[[var_porcentaje_si]] / 100
  } else {
    df_proc$prop_si <- df_proc[[var_porcentaje_si]]
  }

  # Asegurar 0–1 y calcular No como complemento
  df_proc$prop_si <- pmax(pmin(df_proc$prop_si, 1), 0)
  df_proc$prop_no <- 1 - df_proc$prop_si

  df_long <- df_proc |>
    tidyr::pivot_longer(
      cols      = c("prop_si", "prop_no"),
      names_to  = ".cat_interna",
      values_to = "prop"
    ) |>
    dplyr::mutate(
      categoria = dplyr::case_when(
        .cat_interna == "prop_si" ~ etiqueta_si,
        .cat_interna == "prop_no" ~ etiqueta_no,
        TRUE                     ~ NA_character_
      )
    )

  df_long <- df_long[!is.na(df_long$categoria), , drop = FALSE]

  # Etiquetas de facetas (indicadores)
  df_long$indicador_label <- df_long[[var_indicador]]

  # Si se desea incluir N en el título de cada faceta (con salto de línea)
  if (incluir_n_en_titulo) {
    df_long <- df_long |>
      dplyr::group_by(.data[[var_indicador]]) |>
      dplyr::mutate(
        indicador_label = paste0(
          .data[[var_indicador]],
          "\n",
          prefijo_n_titulo,
          .data[[var_n]]
        )
      ) |>
      dplyr::ungroup()
  }

  # Orden de facetas según aparición
  df_long$indicador_label <- factor(
    df_long$indicador_label,
    levels = unique(df_long$indicador_label)
  )

  # ---------------------------------------------------------------------------
  # 2. Orden de categorías (Sí / No) con opción de invertir
  # ---------------------------------------------------------------------------
  niveles_cat <- c(etiqueta_si, etiqueta_no)
  if (invertir_respuestas) {
    niveles_cat <- rev(niveles_cat)
  }
  df_long$categoria <- factor(df_long$categoria, levels = niveles_cat)

  # ---------------------------------------------------------------------------
  # 3. Preparar datos para el pie y etiquetas
  # ---------------------------------------------------------------------------
  df_plot <- df_long |>
    dplyr::group_by(indicador_label) |>
    dplyr::mutate(
      prop = ifelse(is.na(prop), 0, prop)
    ) |>
    dplyr::ungroup()

  # Etiquetas de porcentaje
  df_plot$lab <- scales::percent(df_plot$prop, accuracy = 10^(-decimales))
  df_plot$lab[df_plot$prop < umbral_etiqueta] <- ""

  # ---------------------------------------------------------------------------
  # 4. Gráfico base: pies en facetas
  # ---------------------------------------------------------------------------
  p <- ggplot2::ggplot(
    df_plot,
    ggplot2::aes(
      x    = 1,
      y    = prop,
      fill = categoria
    )
  ) +
    ggplot2::geom_col(width = 1, color = "white") +
    ggplot2::coord_polar(theta = "y")

  # Facetado
  if (is.null(ncol_facetas)) {
    n_ind <- length(unique(df_plot$indicador_label))
    ncol_facetas <- min(n_ind, 3L)
  }

  p <- p +
    ggplot2::facet_wrap(
      ~ indicador_label,
      ncol = ncol_facetas
    )

  # Etiquetas de porcentaje DENTRO del área (usando position_stack)
  if (mostrar_valores) {
    p <- p +
      ggplot2::geom_text(
        data    = df_plot,
        ggplot2::aes(
          x     = 1,
          y     = prop,
          label = lab
        ),
        position    = ggplot2::position_stack(vjust = 0.5),
        inherit.aes = TRUE,
        color       = color_texto_segmentos,
        size        = size_texto_segmentos,
        fontface    = if ("porcentajes" %in% textos_negrita) "bold" else "plain",
        show.legend = FALSE
      )
  }

  # ---------------------------------------------------------------------------
  # 5. Colores y tema
  # ---------------------------------------------------------------------------
  if (is.null(colores_respuesta)) {
    colores_respuesta <- c(
      setNames("#004B8D", etiqueta_si),
      setNames("#D9D9D9", etiqueta_no)
    )
  }

  p <- p +
    ggplot2::scale_fill_manual(values = colores_respuesta) +
    ggplot2::theme_minimal(base_size = 9) +
    ggplot2::theme(
      panel.grid        = ggplot2::element_blank(),
      axis.title        = ggplot2::element_blank(),
      axis.text         = ggplot2::element_blank(),
      axis.ticks        = ggplot2::element_blank(),
      strip.text        = ggplot2::element_text(
        color = color_titulos_pie,
        size  = size_titulos_pie,
        face  = if ("titulos_pie" %in% textos_negrita) "bold" else "plain"
      ),
      legend.position   = if (mostrar_leyenda) posicion_leyenda else "none",
      legend.title      = ggplot2::element_blank(),
      legend.text       = ggplot2::element_text(
        color = color_leyenda,
        size  = size_leyenda,
        face  = if ("leyenda" %in% textos_negrita) "bold" else "plain"
      ),
      plot.title.position = "plot",
      plot.title        = ggplot2::element_text(
        hjust  = 0.5,
        color  = color_titulo,
        size   = size_titulo,
        face   = if ("titulo" %in% textos_negrita) "bold" else "plain",
        margin = ggplot2::margin(b = 12)
      ),
      plot.subtitle     = ggplot2::element_text(
        hjust  = 0.5,
        color  = color_subtitulo,
        size   = size_subtitulo,
        margin = ggplot2::margin(b = 8)
      ),
      # Caption SIEMPRE a la derecha
      plot.caption      = ggplot2::element_text(
        hjust  = 1,
        color  = color_nota_pie,
        size   = size_nota_pie,
        margin = ggplot2::margin(t = 10)
      ),
      plot.margin       = ggplot2::margin(
        t = 30, r = 15, b = 20, l = 15
      )
    )

  p <- p + ggplot2::labs(
    title    = titulo,
    subtitle = subtitulo,
    caption  = nota_pie
  )

  # Invertir leyenda si se solicita (solo leyenda)
  if (invertir_leyenda && mostrar_leyenda) {
    p <- p + ggplot2::guides(
      fill = ggplot2::guide_legend(reverse = TRUE)
    )
  }

  # ---------------------------------------------------------------------------
  # 6. Exportación
  # ---------------------------------------------------------------------------
  if (exportar == "rplot") {
    return(p)
  }

  if (is.null(path_salida) || !nzchar(path_salida)) {
    stop("Debe especificar `path_salida` cuando `exportar` no es 'rplot'.", call. = FALSE)
  }

  if (exportar == "png") {
    ggplot2::ggsave(
      filename = path_salida,
      plot     = p,
      width    = ancho,
      height   = alto,
      dpi      = dpi
    )
    return(invisible(p))
  }

  if (exportar == "ppt") {
    if (!requireNamespace("officer", quietly = TRUE) ||
        !requireNamespace("rvg", quietly = TRUE)) {
      stop(
        "Para exportar a PPT se requieren los paquetes 'officer' y 'rvg'.",
        call. = FALSE
      )
    }

    doc <- officer::read_pptx()
    doc <- officer::add_slide(doc, layout = "Blank", master = "Office Theme")
    doc <- officer::ph_with(
      doc,
      rvg::dml(ggobj = p),
      location = officer::ph_location_fullsize()
    )
    print(doc, target = path_salida)

    return(invisible(p))
  }

  p
}
