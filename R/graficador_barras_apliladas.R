# =============================================================================
# graficar_barras_apiladas()
# -----------------------------------------------------------------------------
# Graficador genérico de barras horizontales apiladas a partir de una tabla
# de indicadores en formato ancho, pensado para reportes tipo ACNUR / OPS.
# =============================================================================

#' Graficar barras horizontales apiladas a partir de una tabla ancha
#'
#' Genera un gráfico de barras horizontales apiladas donde cada categoría
#' (por ejemplo, enfermedades, regiones o grupos etarios) se muestra como una
#' barra compuesta por proporciones asociadas a diferentes estados o respuestas.
#' La función está pensada para usarse directamente con tablas de indicadores
#' ya calculadas (formato ancho), sin necesidad de reestructurar los datos
#' manualmente.
#'
#' @param data Data frame en formato ancho que contiene la categoría, la base N
#'   y las columnas de porcentajes.
#' @param var_categoria Nombre (string) de la columna que define cada barra
#'   (por ejemplo, "condicion").
#' @param var_n Nombre (string) de la columna con el N total por categoría
#'   (se usa como base para la barra extra; típicamente el N).
#' @param cols_porcentaje Vector de nombres de columnas con los porcentajes
#'   que deben formar los segmentos apilados.
#' @param etiquetas_grupos Vector de caracteres con nombre (named) que asigna
#'   etiquetas legibles a las columnas de \code{cols_porcentaje}. Los nombres
#'   deben coincidir exactamente con los nombres de \code{cols_porcentaje}.
#' @param escala_valor Escala de los porcentajes: \code{"proporcion_1"} si los
#'   valores están en 0-1, o \code{"proporcion_100"} si están en 0-100.
#' @param colores_grupos Vector de colores HEX con nombre (opcional) para los
#'   segmentos, usando como nombres las etiquetas de \code{etiquetas_grupos}.
#'   Si se omite, \pkg{ggplot2} asignará colores por defecto.
#' @param mostrar_valores Lógico; si \code{TRUE}, muestra porcentajes dentro
#'   de cada segmento.
#' @param decimales Número de decimales para los porcentajes internos.
#' @param umbral_etiqueta Mínima proporción para mostrar texto dentro de un
#'   segmento (por ejemplo, 0.03 equivale a 3%).
#'
#' @param mostrar_barra_extra Lógico; si \code{TRUE}, agrega una “barra extra”
#'   de texto al extremo derecho de cada barra (típicamente para mostrar el N
#'   o algún indicador agregado como “Top 2 Box”).
#' @param prefijo_barra_extra Texto añadido antes del valor mostrado en la
#'   barra extra (por defecto \code{"N="}).
#' @param titulo_barra_extra Texto que se coloca como encabezado encima de la
#'   columna de barra extra (por ejemplo, "Total", "Top 2 Box"). Si es
#'   \code{NULL}, no se agrega encabezado.
#'
#' @param titulo Título del gráfico.
#' @param subtitulo Subtítulo opcional del gráfico.
#' @param nota_pie Nota o fuente para colocar en el pie de página del gráfico.
#'
#' @param color_titulo Color del título.
#' @param size_titulo Tamaño del título.
#' @param color_subtitulo Color del subtítulo.
#' @param size_subtitulo Tamaño del subtítulo.
#' @param color_nota_pie Color del texto del pie de página.
#' @param size_nota_pie Tamaño del texto del pie de página.
#' @param color_leyenda Color del texto de la leyenda.
#' @param size_leyenda Tamaño del texto de la leyenda.
#' @param color_texto_barras Color del texto dentro de los segmentos.
#' @param size_texto_barras Tamaño del texto dentro de los segmentos.
#' @param color_barra_extra Color del texto de la barra extra.
#' @param size_barra_extra Tamaño del texto de la barra extra.
#' @param color_ejes Color del texto de las categorías en el eje Y.
#' @param size_ejes Tamaño del texto de las categorías en el eje Y.
#'
#' @param extra_derecha_rel Porcentaje adicional a la derecha para ubicar la
#'   barra extra, relativo a la longitud de la barra más larga.
#' @param invertir_leyenda Lógico; si \code{TRUE}, invierte el orden de los
#'   ítems de la leyenda sin alterar el orden de las barras ni de los segmentos.
#' @param invertir_barras Lógico; si \code{TRUE}, invierte el orden en que las
#'   categorías aparecen en el eje Y (orden vertical de las barras).
#' @param textos_negrita Vector de caracteres que indica qué elementos deben
#'   mostrarse en negrita. Puede incluir cualquiera de:
#'   \code{"titulo"}, \code{"porcentajes"}, \code{"leyenda"}, \code{"barra_extra"}.
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
#'
#' @examples
#' \dontrun{
#' colores_frec <- c(
#'   "Nunca faltan"          = "#F6D55C",
#'   "Rara vez faltan"       = "#88CC88",
#'   "A veces faltan"        = "#1B9E77",
#'   "Faltan frecuentemente" = "#F4A261",
#'   "Siempre faltan"        = "#D95F02"
#' )
#'
#' p <- graficar_barras_apiladas(
#'   data             = tab_stock,
#'   var_categoria    = "condicion",
#'   var_n            = "n_base",
#'   cols_porcentaje  = c("pct_nunca","pct_rara","pct_aveces","pct_freq","pct_siempre"),
#'   etiquetas_grupos = c(
#'     pct_nunca   = "Nunca faltan",
#'     pct_rara    = "Rara vez faltan",
#'     pct_aveces  = "A veces faltan",
#'     pct_freq    = "Faltan frecuentemente",
#'     pct_siempre = "Siempre faltan"
#'   ),
#'   escala_valor        = "proporcion_100",
#'   colores_grupos      = colores_frec,
#'   titulo              = "Disponibilidad de medicamentos básicos",
#'   nota_pie            = "Fuente: Pulso PUCP 2025",
#'   mostrar_barra_extra = TRUE,
#'   prefijo_barra_extra = "N=",
#'   titulo_barra_extra  = "Total",
#'   textos_negrita      = c("titulo", "porcentajes", "barra_extra"),
#'   invertir_barras     = FALSE,
#'   invertir_leyenda    = TRUE
#' )
#' }
#'
#' @export
graficar_barras_apiladas <- function(
    data,
    var_categoria,
    var_n,
    cols_porcentaje,
    etiquetas_grupos,
    escala_valor        = c("proporcion_1", "proporcion_100"),
    colores_grupos      = NULL,
    mostrar_valores     = TRUE,
    decimales           = 1,
    umbral_etiqueta     = 0.03,
    mostrar_barra_extra = TRUE,
    prefijo_barra_extra = "N=",
    titulo_barra_extra  = NULL,
    titulo              = NULL,
    subtitulo           = NULL,
    nota_pie            = NULL,
    # Estilo de texto y layout
    color_titulo        = "#000000",
    size_titulo         = 11,
    color_subtitulo     = "#000000",
    size_subtitulo      = 9,
    color_nota_pie      = "#000000",
    size_nota_pie       = 8,
    color_leyenda       = "#000000",
    size_leyenda        = 8,
    color_texto_barras  = "white",
    size_texto_barras   = 3,
    color_barra_extra   = "#000000",
    size_barra_extra    = 3,
    color_ejes          = "#000000",
    size_ejes           = 9,
    extra_derecha_rel   = 0.10,
    invertir_leyenda    = FALSE,
    invertir_barras     = FALSE,
    textos_negrita      = NULL,
    exportar            = c("rplot", "png", "ppt"),
    path_salida         = NULL,
    ancho               = 10,
    alto                = 6,
    dpi                 = 300
) {

  # Operador auxiliar interno
  `%||%` <- function(x, y) if (!is.null(x)) x else y

  escala_valor <- match.arg(escala_valor)
  exportar     <- match.arg(exportar)

  # ---------------------------------------------------------------------------
  # 0. Validaciones
  # ---------------------------------------------------------------------------
  if (!var_categoria %in% names(data)) {
    stop("`var_categoria` no existe en `data`.", call. = FALSE)
  }
  if (!var_n %in% names(data)) {
    stop("`var_n` no existe en `data`.", call. = FALSE)
  }
  if (!all(cols_porcentaje %in% names(data))) {
    faltan <- cols_porcentaje[!cols_porcentaje %in% names(data)]
    stop(
      "Las siguientes columnas de `cols_porcentaje` no existen en `data`: ",
      paste(faltan, collapse = ", "),
      call. = FALSE
    )
  }
  if (!all(names(etiquetas_grupos) %in% cols_porcentaje)) {
    stop(
      "Los nombres de `etiquetas_grupos` deben coincidir con columnas de `cols_porcentaje`.",
      call. = FALSE
    )
  }

  # Normalizar textos_negrita
  textos_negrita <- textos_negrita %||% character(0)

  df <- data

  # ---------------------------------------------------------------------------
  # 1. Ancho → largo
  # ---------------------------------------------------------------------------
  df_long <- df |>
    dplyr::select(
      dplyr::all_of(c(var_categoria, var_n, cols_porcentaje))
    ) |>
    tidyr::pivot_longer(
      cols      = dplyr::all_of(cols_porcentaje),
      names_to  = ".col_pct",
      values_to = ".valor"
    ) |>
    dplyr::mutate(
      .grupo = dplyr::recode(.data$.col_pct, !!!etiquetas_grupos)
    )

  if (!is.numeric(df_long$.valor)) {
    stop("Las columnas de porcentaje deben ser numéricas.", call. = FALSE)
  }

  # Escala interna: proporción (si viene en 0–100 la llevamos a 0–1)
  if (escala_valor == "proporcion_100") {
    df_long$.valor_plot <- df_long$.valor / 100
  } else {
    df_long$.valor_plot <- df_long$.valor
  }

  # Orden de grupos = orden en etiquetas_grupos
  df_long$.grupo <- factor(df_long$.grupo, levels = unname(etiquetas_grupos))

  # Orden de categorías = orden de aparición (y opción de invertir)
  cat_vec  <- df_long[[var_categoria]]
  cat_lvls <- unique(cat_vec)
  if (invertir_barras) {
    cat_lvls <- rev(cat_lvls)
  }
  df_long[[var_categoria]] <- factor(cat_vec, levels = cat_lvls)

  # Suma por categoría (longitud total de cada barra)
  df_sum <- df_long |>
    dplyr::group_by(.data[[var_categoria]]) |>
    dplyr::summarise(suma = sum(.valor_plot, na.rm = TRUE), .groups = "drop")

  max_suma <- max(df_sum$suma, na.rm = TRUE)
  # margen derecho ajustable: p.ej. 10% más que la barra más larga
  x_max <- max_suma * (1 + extra_derecha_rel)

  # ---------------------------------------------------------------------------
  # 2. Gráfico base (una barra apilada por categoría)
  # ---------------------------------------------------------------------------
  p <- ggplot2::ggplot(
    df_long,
    ggplot2::aes_string(
      x    = ".valor_plot",
      y    = var_categoria,
      fill = ".grupo"
    )
  ) +
    ggplot2::geom_col(width = 0.7)

  # ---------------------------------------------------------------------------
  # 3. Etiquetas dentro de las barras (centradas en cada bloque)
  # ---------------------------------------------------------------------------
  if (mostrar_valores) {
    df_lab <- df_long

    df_lab$lab <- scales::percent(df_lab$.valor_plot, accuracy = 10^(-decimales))
    # Ocultar segmentos muy pequeños o 0
    df_lab$lab[df_lab$.valor_plot < umbral_etiqueta] <- ""

    p <- p +
      ggplot2::geom_text(
        data        = df_lab,
        mapping     = ggplot2::aes_string(
          x     = ".valor_plot",
          y     = var_categoria,
          label = "lab",
          fill  = ".grupo"
        ),
        inherit.aes = FALSE,
        position    = ggplot2::position_stack(vjust = 0.5),
        color       = color_texto_barras,
        size        = size_texto_barras,
        fontface    = if ("porcentajes" %in% textos_negrita) "bold" else "plain",
        show.legend = FALSE
      )
  }

  # ---------------------------------------------------------------------------
  # 4. Escala X sin eje visible (con margen a la izquierda)
  # ---------------------------------------------------------------------------
  p <- p +
    ggplot2::scale_x_continuous(
      limits = c(0, x_max),
      # margen pequeño a la izquierda para que la barra no "pegue" al eje
      expand = ggplot2::expansion(mult = c(0.05, 0))
    ) +
    ggplot2::coord_cartesian(clip = "off")

  # ---------------------------------------------------------------------------
  # 5. Barra extra al costado derecho (N, Top 2 Box, etc.)
  # ---------------------------------------------------------------------------
  if (mostrar_barra_extra) {
    df_extra <- df_sum |>
      dplyr::left_join(
        df |>
          dplyr::select(dplyr::all_of(c(var_categoria, var_n))) |>
          dplyr::distinct(),
        by = var_categoria
      ) |>
      dplyr::mutate(
        xpos      = suma * (1 + extra_derecha_rel * 0.5),
        lab_extra = paste0(prefijo_barra_extra, .data[[var_n]])
      )

    # Texto principal de barra_extra
    p <- p +
      ggplot2::geom_text(
        data        = df_extra,
        mapping     = ggplot2::aes_string(
          x     = "xpos",
          y     = var_categoria,
          label = "lab_extra"
        ),
        inherit.aes = FALSE,
        hjust       = 0,
        size        = size_barra_extra,
        color       = color_barra_extra,
        fontface    = if ("barra_extra" %in% textos_negrita) "bold" else "plain"
      )

    # Encabezado encima de la columna de barra_extra (opcional)
    if (!is.null(titulo_barra_extra) && nzchar(titulo_barra_extra)) {

      # Determinar categoría superior según inversión
      lvls <- levels(df_long[[var_categoria]])
      cat_superior <- if (invertir_barras) tail(lvls, 1) else head(lvls, 1)

      df_header <- df_extra[df_extra[[var_categoria]] == cat_superior, , drop = FALSE]

      if (nrow(df_header) == 1L) {
        p <- p +
          ggplot2::geom_text(
            data        = df_header,
            mapping     = ggplot2::aes_string(
              x = "xpos",
              y = var_categoria
            ),
            label       = titulo_barra_extra,
            inherit.aes = FALSE,
            hjust       = 0,      # alineado con la izquierda del "N=..."
            vjust       = -1.2,   # un poco por encima de la primera barra
            size        = size_barra_extra,
            color       = color_barra_extra,
            fontface    = if ("barra_extra" %in% textos_negrita) "bold" else "plain"
          )
      }
    }
  }

  # ---------------------------------------------------------------------------
  # 6. Colores y tema
  # ---------------------------------------------------------------------------
  if (!is.null(colores_grupos)) {
    p <- p +
      ggplot2::scale_fill_manual(values = colores_grupos)
  }

  p <- p +
    ggplot2::theme_minimal(base_size = 9) +
    ggplot2::theme(
      panel.grid.major.y = ggplot2::element_blank(),
      panel.grid.minor   = ggplot2::element_blank(),
      panel.grid.major.x = ggplot2::element_blank(),
      axis.title.x       = ggplot2::element_blank(),
      axis.text.x        = ggplot2::element_blank(),
      axis.ticks.x       = ggplot2::element_blank(),
      axis.title.y       = ggplot2::element_blank(),
      axis.text.y        = ggplot2::element_text(
        color = color_ejes,
        size  = size_ejes,
        hjust = 0.5,
        vjust = 0.5
      ),
      axis.line.y        = ggplot2::element_line(color = color_ejes, linewidth = 0.3),
      legend.title       = ggplot2::element_blank(),
      legend.position    = "bottom",
      legend.text        = ggplot2::element_text(
        color = color_leyenda,
        size  = size_leyenda,
        face  = if ("leyenda" %in% textos_negrita) "bold" else "plain"
      ),
      plot.margin        = ggplot2::margin(t = 15, r = 80, b = 5, l = 5),
      plot.title         = ggplot2::element_text(
        hjust = 0.5,
        color = color_titulo,
        size  = size_titulo,
        face  = if ("titulo" %in% textos_negrita) "bold" else "plain"
      ),
      plot.subtitle      = ggplot2::element_text(
        hjust = 0.5,
        color = color_subtitulo,
        size  = size_subtitulo
      ),
      plot.caption       = ggplot2::element_text(
        hjust = 0,
        color = color_nota_pie,
        size  = size_nota_pie
      )
    )

  # Invertir solo el orden de la leyenda si se solicita
  if (invertir_leyenda) {
    p <- p + ggplot2::guides(fill = ggplot2::guide_legend(reverse = TRUE))
  }

  p <- p + ggplot2::labs(
    title    = titulo,
    subtitle = subtitulo,
    caption  = nota_pie
  )

  # ---------------------------------------------------------------------------
  # 7. Exportación
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
