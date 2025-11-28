# =============================================================================
# graficar_barras_agrupadas()
# -----------------------------------------------------------------------------
# Barras horizontales agrupadas para comparar 2+ indicadores por categoría
# (por ejemplo, % capacitado en el área vs % capacitado en M&R).
# =============================================================================

#' Graficar barras horizontales agrupadas a partir de una tabla ancha
#'
#' Genera un gráfico de barras horizontales agrupadas donde cada categoría
#' (por ejemplo, servicios, enfermedades, regiones) se muestra en el eje
#' vertical, y para cada una se comparan dos o más porcentajes en el eje
#' horizontal (por ejemplo, "\% capacitado en el área" y "\% capacitado en M\&R").
#'
#' La función está pensada para usarse directamente con tablas de indicadores
#' ya calculadas (formato ancho), sin necesidad de reestructurar los datos
#' manualmente.
#'
#' @param data Data frame en formato ancho que contiene la categoría, la base N
#'   y las columnas de porcentajes.
#' @param var_categoria Nombre (string) de la columna que define cada categoría
#'   en el eje vertical (por ejemplo, `"servicio"` o `"condicion"`).
#' @param var_n Nombre (string) de la columna con el N total por categoría
#'   (se usa para mostrar una etiqueta adicional, típicamente `"N="`, al
#'   extremo derecho de las barras).
#' @param cols_porcentaje Vector de nombres de columnas con los porcentajes
#'   que se desean comparar como series agrupadas (por ejemplo,
#'   `c("pct_area", "pct_mr")`).
#' @param etiquetas_series Vector de caracteres con nombre (named) que asigna
#'   etiquetas legibles a las columnas de \code{cols_porcentaje}. Los nombres
#'   deben coincidir exactamente con los nombres de \code{cols_porcentaje}.
#'   Los valores (etiquetas) se usan en la leyenda.
#' @param escala_valor Escala de los porcentajes: \code{"proporcion_1"} si los
#'   valores están en 0–1, o \code{"proporcion_100"} si están en 0–100.
#' @param colores_series Vector de colores HEX con nombre (opcional) para las
#'   series, usando como nombres las etiquetas de \code{etiquetas_series}. Si
#'   se omite, \pkg{ggplot2} asignará colores por defecto.
#' @param mostrar_valores Lógico; si \code{TRUE}, muestra los porcentajes
#'   asociados a cada barra.
#' @param decimales Número de decimales para los porcentajes.
#' @param umbral_etiqueta Mínima proporción (en escala 0–1) para mostrar texto
#'   (por ejemplo, 0.03 equivale a 3\%). Por debajo de este valor, no se
#'   muestra la etiqueta.
#' @param umbral_posicion Proporción (0–1) que define desde qué tamaño de barra
#'   el porcentaje se dibuja centrado dentro de la barra. Si la barra es
#'   menor a este umbral, el porcentaje se dibuja ligeramente a la derecha del
#'   final de la barra.
#'
#' @param mostrar_barra_extra Lógico; si \code{TRUE}, agrega una "barra extra"
#'   de texto al extremo derecho de cada categoría (típicamente el N, por
#'   ejemplo `"N=305"`).
#' @param prefijo_barra_extra Texto añadido antes del valor mostrado en la
#'   barra extra (por defecto \code{"N="}).
#' @param titulo_barra_extra Texto que se coloca como encabezado encima de la
#'   columna de barra extra (por ejemplo, `"Total"`). Si es \code{NULL}, no se
#'   agrega encabezado.
#'
#' @param titulo Título del gráfico.
#' @param subtitulo Subtítulo opcional del gráfico.
#' @param nota_pie Nota o fuente para el pie de página (por ejemplo,
#'   `"Fuente: Pulso PUCP 2025"`).
#' @param nota_pie_derecha Texto adicional para el pie de página que se
#'   concatenará a la derecha de \code{nota_pie} (por ejemplo,
#'   `"N total = 52"`). Ambos se muestran en una única línea de caption,
#'   alineada a la derecha.
#'
#' @param color_titulo Color del título.
#' @param size_titulo Tamaño del título.
#' @param color_subtitulo Color del subtítulo.
#' @param size_subtitulo Tamaño del subtítulo.
#' @param color_nota_pie Color del texto del pie de página (caption).
#' @param size_nota_pie Tamaño del texto del pie de página (caption).
#' @param color_leyenda Color del texto de la leyenda.
#' @param size_leyenda Tamaño del texto de la leyenda.
#' @param color_texto_barras Color del texto de los porcentajes.
#' @param size_texto_barras Tamaño del texto de los porcentajes.
#' @param color_barra_extra Color del texto de la barra extra (N).
#' @param size_barra_extra Tamaño del texto de la barra extra (N).
#' @param color_ejes Color del texto de las categorías en el eje.
#' @param size_ejes Tamaño del texto de las categorías en el eje.
#'
#' @param extra_derecha_rel Porcentaje adicional de espacio a la derecha,
#'   relativo a la barra más larga, para ubicar la barra extra.
#' @param mostrar_leyenda Lógico; si \code{FALSE}, oculta la leyenda (útil
#'   cuando solo hay una serie).
#' @param invertir_leyenda Lógico; si \code{TRUE}, invierte el orden de los
#'   ítems de la leyenda.
#' @param invertir_barras Lógico; si \code{TRUE}, invierte el orden en que las
#'   categorías aparecen en el eje vertical (orden de las barras).
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
#' @export
graficar_barras_agrupadas <- function(
    data,
    var_categoria,
    var_n,
    cols_porcentaje,
    etiquetas_series,
    escala_valor              = c("proporcion_1", "proporcion_100"),
    colores_series            = NULL,
    mostrar_valores           = TRUE,
    decimales                 = 1,
    umbral_etiqueta           = 0.03,
    umbral_posicion           = 0.15,
    mostrar_barra_extra       = TRUE,
    prefijo_barra_extra       = "N=",
    titulo_barra_extra        = NULL,
    titulo                    = NULL,
    subtitulo                 = NULL,
    nota_pie                  = NULL,
    nota_pie_derecha          = NULL,
    # Estilo de texto y layout
    color_titulo              = "#000000",
    size_titulo               = 11,
    color_subtitulo           = "#000000",
    size_subtitulo            = 9,
    color_nota_pie            = "#000000",
    size_nota_pie             = 8,
    color_leyenda             = "#000000",
    size_leyenda              = 8,
    color_texto_barras        = "white",
    size_texto_barras         = 3,
    color_barra_extra         = "#000000",
    size_barra_extra          = 3,
    color_ejes                = "#000000",
    size_ejes                 = 9,
    extra_derecha_rel         = 0.10,
    mostrar_leyenda           = TRUE,
    invertir_leyenda          = FALSE,
    invertir_barras           = FALSE,
    textos_negrita            = NULL,
    exportar                  = c("rplot", "png", "ppt"),
    path_salida               = NULL,
    ancho                     = 10,
    alto                      = 6,
    dpi                       = 300
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y

  escala_valor <- match.arg(escala_valor)
  exportar     <- match.arg(exportar)

  # ---------------------------------------------------------------------------
  # 0. Validaciones básicas
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
  if (!all(names(etiquetas_series) %in% cols_porcentaje)) {
    stop(
      "Los nombres de `etiquetas_series` deben coincidir con columnas de `cols_porcentaje`.",
      call. = FALSE
    )
  }

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
    )

  if (!is.numeric(df_long$.valor)) {
    stop("Las columnas de porcentaje deben ser numéricas.", call. = FALSE)
  }

  # Etiquetas legibles de series
  df_long$.serie <- dplyr::recode(df_long$.col_pct, !!!etiquetas_series)

  # Escala interna 0–1
  if (escala_valor == "proporcion_100") {
    df_long$.valor_plot <- df_long$.valor / 100
  } else {
    df_long$.valor_plot <- df_long$.valor
  }

  # Orden de series según etiquetas_series
  df_long$.serie <- factor(df_long$.serie, levels = unname(etiquetas_series))

  # Orden de categorías según aparición (con opción de invertir)
  cat_vec  <- df_long[[var_categoria]]
  cat_lvls <- unique(cat_vec)
  if (invertir_barras) {
    cat_lvls <- rev(cat_lvls)
  }
  df_long[[var_categoria]] <- factor(cat_vec, levels = cat_lvls)

  # Máximo valor para definir escala y espacio extra
  max_valor <- max(df_long$.valor_plot, na.rm = TRUE)
  y_max     <- max_valor * (1 + extra_derecha_rel)

  # ---------------------------------------------------------------------------
  # 2. Gráfico base (antes de coord_flip)
  # ---------------------------------------------------------------------------
  p <- ggplot2::ggplot(
    df_long,
    ggplot2::aes_string(
      x    = var_categoria,
      y    = ".valor_plot",
      fill = ".serie"
    )
  ) +
    ggplot2::geom_col(
      position = ggplot2::position_dodge(width = 0.7),
      width    = 0.6
    )

  # ---------------------------------------------------------------------------
  # 3. Etiquetas de porcentaje dentro o fuera de la barra
  # ---------------------------------------------------------------------------
  if (mostrar_valores) {
    df_lab <- df_long
    df_lab$lab <- scales::percent(df_lab$.valor_plot, accuracy = 10^(-decimales))

    # No mostrar etiquetas por debajo del umbral mínimo
    df_lab$lab[df_lab$.valor_plot < umbral_etiqueta] <- ""

    # Offset pequeño para las que van fuera
    offset_lab <- max_valor * 0.015

    df_lab$valor_label <- ifelse(
      df_lab$.valor_plot >= umbral_posicion,
      # Barras grandes: etiqueta al centro de la barra
      df_lab$.valor_plot / 2,
      # Barras pequeñas: etiqueta ligeramente a la derecha del final
      df_lab$.valor_plot + offset_lab
    )

    p <- p +
      ggplot2::geom_text(
        data        = df_lab,
        mapping     = ggplot2::aes_string(
          x     = var_categoria,
          y     = "valor_label",
          label = "lab",
          group = ".serie"
        ),
        inherit.aes = FALSE,
        position    = ggplot2::position_dodge(width = 0.7),
        hjust       = ifelse(df_lab$.valor_plot >= umbral_posicion, 0.5, 0),
        vjust       = 0.5,
        color       = color_texto_barras,
        size        = size_texto_barras,
        fontface    = if ("porcentajes" %in% textos_negrita) "bold" else "plain",
        show.legend = FALSE
      )
  }

  # ---------------------------------------------------------------------------
  # 4. Escala Y (valores) con espacio extra
  # ---------------------------------------------------------------------------
  p <- p +
    ggplot2::scale_y_continuous(
      limits = c(0, y_max),
      expand = ggplot2::expansion(mult = c(0, 0.05))
    )

  # ---------------------------------------------------------------------------
  # 5. Barra extra con N a la derecha
  # ---------------------------------------------------------------------------
  if (mostrar_barra_extra) {
    df_extra <- df |>
      dplyr::select(dplyr::all_of(c(var_categoria, var_n))) |>
      dplyr::distinct() |>
      dplyr::mutate(
        # llevamos N casi hasta el borde derecho del eje
        ypos      = max_valor * (1 + extra_derecha_rel * 0.95),
        lab_extra = paste0(prefijo_barra_extra, .data[[var_n]])
      )

    # Texto N=...
    p <- p +
      ggplot2::geom_text(
        data        = df_extra,
        mapping     = ggplot2::aes_string(
          x     = var_categoria,
          y     = "ypos",
          label = "lab_extra"
        ),
        inherit.aes = FALSE,
        hjust       = 0,
        vjust       = 0.5,
        size        = size_barra_extra,
        color       = color_barra_extra,
        fontface    = if ("barra_extra" %in% textos_negrita) "bold" else "plain"
      )

    # Encabezado encima de la columna de barra extra (titulo_barra_extra)
    if (!is.null(titulo_barra_extra) && nzchar(titulo_barra_extra)) {

      lvls <- levels(df_long[[var_categoria]])
      cat_superior <- if (invertir_barras) tail(lvls, 1) else head(lvls, 1)

      df_header <- df_extra[df_extra[[var_categoria]] == cat_superior, , drop = FALSE]

      if (nrow(df_header) == 1L) {
        p <- p +
          ggplot2::geom_text(
            data        = df_header,
            mapping     = ggplot2::aes_string(
              x = var_categoria,
              y = "ypos"
            ),
            label       = titulo_barra_extra,
            inherit.aes = FALSE,
            hjust       = 0,
            vjust       = -1.2,
            size        = size_barra_extra,
            color       = color_barra_extra,
            fontface    = if ("barra_extra" %in% textos_negrita) "bold" else "plain"
          )
      }
    }
  }

  # ---------------------------------------------------------------------------
  # 6. Colores, orientación horizontal y tema
  # ---------------------------------------------------------------------------
  if (!is.null(colores_series)) {
    p <- p +
      ggplot2::scale_fill_manual(values = colores_series)
  }

  # Construir caption combinado (derecha)
  caption_text <- NULL
  if (!is.null(nota_pie) && nzchar(nota_pie) &&
      !is.null(nota_pie_derecha) && nzchar(nota_pie_derecha)) {
    caption_text <- paste0(nota_pie, "   ", nota_pie_derecha)
  } else if (!is.null(nota_pie) && nzchar(nota_pie)) {
    caption_text <- nota_pie
  } else if (!is.null(nota_pie_derecha) && nzchar(nota_pie_derecha)) {
    caption_text <- nota_pie_derecha
  }

  p <- p +
    ggplot2::coord_flip() +
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
      legend.position    = if (mostrar_leyenda) "bottom" else "none",
      legend.text        = ggplot2::element_text(
        color = color_leyenda,
        size  = size_leyenda,
        face  = if ("leyenda" %in% textos_negrita) "bold" else "plain"
      ),
      plot.margin        = ggplot2::margin(t = 15, r = 80, b = 15, l = 5),
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
        hjust = 1,  # SIEMPRE alineado a la derecha
        color = color_nota_pie,
        size  = size_nota_pie
      )
    )

  # Leyenda invertida si se pide
  if (invertir_leyenda && mostrar_leyenda) {
    p <- p + ggplot2::guides(fill = ggplot2::guide_legend(reverse = TRUE))
  }

  # Etiquetas de título y caption
  p <- p + ggplot2::labs(
    title    = titulo,
    subtitle = subtitulo,
    caption  = caption_text
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
