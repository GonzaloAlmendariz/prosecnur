#' Graficar barras horizontales agrupadas a partir de una tabla ancha
#'
#' Genera un gráfico de barras horizontales agrupadas donde cada categoría
#' (por ejemplo, servicios, enfermedades, regiones) se muestra en el eje
#' vertical, y para cada una se comparan dos o más porcentajes en el eje
#' horizontal (por ejemplo, "% capacitado en el área" y "% capacitado en M&R").
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
#' @param mostrar_barra_extra Lógico; si \code{TRUE}, agrega una "barra extra"
#'   de texto al extremo derecho de cada categoría (típicamente el N, por
#'   ejemplo `"N=305"`).
#' @param prefijo_barra_extra Texto añadido antes del valor mostrado en la
#'   barra extra (por defecto \code{"N="}).
#' @param titulo_barra_extra Texto que se coloca como encabezado encima de la
#'   columna de barra extra (por ejemplo, `"Total"`). Si es \code{NULL}, no se
#'   agrega encabezado.
#' @param titulo Título del gráfico.
#' @param subtitulo Subtítulo opcional del gráfico.
#' @param nota_pie Nota o fuente para el pie de página.
#' @param nota_pie_derecha Texto adicional para el pie de página que se
#'   concatenará a la derecha de \code{nota_pie}.
#' @param pos_titulo Alineación horizontal del título: `"centro"`, `"izquierda"`
#'   o `"derecha"`.
#' @param pos_nota_pie Alineación horizontal de la nota al pie (caption):
#'   `"derecha"`, `"izquierda"` o `"centro"`.
#' @param color_titulo,size_titulo,color_subtitulo,size_subtitulo,
#'   color_nota_pie,size_nota_pie,color_leyenda,size_leyenda,
#'   color_texto_barras,color_texto_barras_fuera,size_texto_barras,
#'   color_barra_extra,size_barra_extra,color_ejes,size_ejes,
#'   color_fondo Parámetros estéticos.
#' @param grosor_barras Grosor relativo de cada barra dentro de su grupo.
#' @param extra_derecha_rel Porción adicional a la derecha, en unidades de
#'   proporción, para la barra extra.
#' @param mostrar_leyenda,invertir_leyenda,invertir_barras,invertir_series
#'   Controles de leyenda y orden.
#' @param textos_negrita Vector que puede incluir `"titulo"`, `"porcentajes"`,
#'   `"leyenda"` y/o `"barra_extra"`.
#' @param exportar `"rplot"`, `"png"`, `"ppt"` o `"word"`.
#' @param path_salida Ruta del archivo cuando `exportar != "rplot"`.
#' @param ancho,alto,dpi Parámetros de exportación.
#'
#' @return Un objeto \code{ggplot} si \code{exportar = "rplot"}.
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
    pos_titulo                = c("centro", "izquierda", "derecha"),
    pos_nota_pie              = c("derecha", "izquierda", "centro"),
    # Estilo de texto y layout
    color_titulo              = "#004B8D",
    size_titulo               = 11,
    color_subtitulo           = "#004B8D",
    size_subtitulo            = 9,
    color_nota_pie            = "#004B8D",
    size_nota_pie             = 8,
    color_leyenda             = "#004B8D",
    size_leyenda              = 8,
    color_texto_barras        = "white",
    color_texto_barras_fuera  = "#004B8D",
    size_texto_barras         = 3,
    color_barra_extra         = "#004B8D",
    size_barra_extra          = 3,
    color_ejes                = "#004B8D",
    size_ejes                 = 9,
    color_fondo               = NA,
    grosor_barras             = 0.6,
    extra_derecha_rel         = 0.25,
    mostrar_leyenda           = TRUE,
    invertir_leyenda          = FALSE,
    invertir_barras           = FALSE,
    invertir_series           = FALSE,
    textos_negrita            = NULL,
    exportar                  = c("rplot", "png", "ppt", "word"),
    path_salida               = NULL,
    ancho                     = 10,
    alto                      = 6,
    dpi                       = 300
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y

  escala_valor <- match.arg(escala_valor)
  exportar     <- match.arg(exportar)
  pos_titulo   <- match.arg(pos_titulo)
  pos_nota_pie <- match.arg(pos_nota_pie)

  decimales <- suppressWarnings(as.numeric(decimales))
  if (length(decimales) < 1L || !is.finite(decimales[1])) {
    decimales <- 1
  } else {
    decimales <- decimales[1]
  }

  hjust_from_pos <- function(x) {
    switch(
      x,
      "izquierda" = 0,
      "centro"    = 0.5,
      "derecha"   = 1,
      0.5
    )
  }

  hjust_titulo  <- hjust_from_pos(pos_titulo)
  hjust_caption <- hjust_from_pos(pos_nota_pie)

  # ---------------------------------------------------------------------------
  # Validaciones básicas
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

  df_long$.serie <- dplyr::recode(df_long$.col_pct, !!!etiquetas_series)

  # Escala interna 0–1
  if (escala_valor == "proporcion_100") {
    df_long$.valor_plot <- df_long$.valor / 100
  } else {
    df_long$.valor_plot <- df_long$.valor
  }

  # Mantener ceros; SOLO NA → 0
  df_long$.valor_plot[is.na(df_long$.valor_plot)] <- 0

  # Orden de series
  niveles_series <- unname(etiquetas_series)
  if (invertir_series) {
    niveles_series <- rev(niveles_series)
  }
  df_long$.serie <- factor(df_long$.serie, levels = niveles_series)

  # Orden de categorías
  cat_vec  <- df_long[[var_categoria]]
  cat_lvls <- unique(cat_vec)
  if (invertir_barras) {
    cat_lvls <- rev(cat_lvls)
  }
  df_long[[var_categoria]] <- factor(cat_vec, levels = cat_lvls)

  # Tamaño de texto según nº de series
  n_series <- length(unique(df_long$.serie))
  size_texto_barras_eff <- dplyr::case_when(
    n_series <= 2 ~ size_texto_barras * 1.00,
    n_series == 3 ~ size_texto_barras * 0.85,
    n_series == 4 ~ size_texto_barras * 0.70,
    TRUE          ~ size_texto_barras * 0.55
  )

  max_valor <- max(df_long$.valor_plot, na.rm = TRUE)

  # ---------------------------------------------------------------------------
  # 2. Gráfico base
  # ---------------------------------------------------------------------------
  width_dodge <- 0.7

  p <- ggplot2::ggplot(
    df_long,
    ggplot2::aes_string(
      x    = var_categoria,
      y    = ".valor_plot",
      fill = ".serie"
    )
  ) +
    ggplot2::geom_col(
      position = ggplot2::position_dodge(width = width_dodge),
      width    = grosor_barras
    )

  # ---------------------------------------------------------------------------
  # 3. Etiquetas de porcentaje (un solo geom_text, siempre alineado al dodge)
  # ---------------------------------------------------------------------------
  if (mostrar_valores) {

    df_lab <- df_long

    # Texto de etiqueta
    df_lab$lab <- scales::percent(
      df_lab$.valor_plot,
      accuracy = 10^(-decimales)
    )

    # Nunca etiquetar 0 % ni valores bajo umbral de visibilidad
    df_lab$lab[df_lab$.valor_plot <= 0] <- ""
    df_lab$lab[df_lab$.valor_plot < umbral_etiqueta] <- ""

    # Umbral para decidir dentro / fuera
    umbral_posicion_eff <- umbral_posicion
    if (!is.finite(umbral_posicion_eff) || umbral_posicion_eff <= 0) {
      umbral_posicion_eff <- 0.15
    }

    offset_lab <- max_valor * 0.015

    # ¿Etiqueta "dentro" de la barra?
    df_lab$inside <- df_lab$.valor_plot >= umbral_posicion_eff & df_lab$lab != ""

    # Posición horizontal (antes del coord_flip):
    #  - dentro: en el centro (valor/2)
    #  - fuera: justo después del extremo (valor + offset)
    df_lab$valor_label <- df_lab$.valor_plot
    df_lab$valor_label[df_lab$inside] <-
      df_lab$.valor_plot[df_lab$inside] / 2
    df_lab$valor_label[!df_lab$inside & df_lab$.valor_plot > 0] <-
      df_lab$.valor_plot[!df_lab$inside & df_lab$.valor_plot > 0] + offset_lab

    # hjust por fila
    df_lab$hjust_label <- ifelse(df_lab$inside, 0.5, 0)

    # color por fila
    df_lab$col_label <- ifelse(
      df_lab$inside,
      color_texto_barras,
      color_texto_barras_fuera
    )

    p <- p +
      ggplot2::geom_text(
        data        = df_lab,
        mapping     = ggplot2::aes(
          x       = .data[[var_categoria]],
          y       = valor_label,
          label   = lab,
          group   = .serie,
          colour  = col_label,
          hjust   = hjust_label    # ← AHORA AQUÍ, CORRECTO
        ),
        inherit.aes = FALSE,
        position    = ggplot2::position_dodge(width = width_dodge),
        vjust       = 0.5,
        size        = size_texto_barras_eff,
        fontface    = if ("porcentajes" %in% textos_negrita) "bold" else "plain",
        show.legend = FALSE
      ) +
      ggplot2::scale_colour_identity(guide = "none")
  }

  # ---------------------------------------------------------------------------
  # 4. Escala Y (proporción + espacio para barra extra)
  # ---------------------------------------------------------------------------
  if (escala_valor %in% c("proporcion_1", "proporcion_100")) {

    if (mostrar_barra_extra) {
      y_lim   <- 1 + extra_derecha_rel
      y_extra <- 1 + extra_derecha_rel * 0.5
    } else {
      y_lim   <- 1
      y_extra <- NA_real_
    }

    p <- p +
      ggplot2::scale_y_continuous(
        limits = c(0, y_lim),
        breaks = seq(0.2, 1, by = 0.2),
        labels = scales::percent_format(accuracy = 1),
        expand = ggplot2::expansion(mult = c(0, 0.02))
      )

  } else {
    y_lim   <- max_valor * (1 + extra_derecha_rel)
    y_extra <- max_valor * (1 + extra_derecha_rel * 0.95)

    p <- p +
      ggplot2::scale_y_continuous(
        limits = c(0, y_lim),
        expand = ggplot2::expansion(mult = c(0, 0.05))
      )
  }

  # ---------------------------------------------------------------------------
  # 5. Barra extra con N a la derecha
  # ---------------------------------------------------------------------------
  if (mostrar_barra_extra) {
    df_extra <- df |>
      dplyr::select(dplyr::all_of(c(var_categoria, var_n))) |>
      dplyr::distinct() |>
      dplyr::filter(.data[[var_categoria]] %in% levels(df_long[[var_categoria]])) |>
      dplyr::mutate(
        ypos      = if (!is.na(y_extra)) y_extra else max_valor * (1 + extra_derecha_rel * 0.95),
        lab_extra = paste0(prefijo_barra_extra, .data[[var_n]])
      )

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
  # 6. Colores, orientación y tema
  # ---------------------------------------------------------------------------
  if (!is.null(colores_series)) {
    p <- p +
      ggplot2::scale_fill_manual(values = colores_series)
  }

  caption_text <- NULL
  if (!is.null(nota_pie) && nzchar(nota_pie) &&
      !is.null(nota_pie_derecha) && nzchar(nota_pie_derecha)) {
    caption_text <- paste0(nota_pie, "   ", nota_pie_derecha)
  } else if (!is.null(nota_pie) && nzchar(nota_pie)) {
    caption_text <- nota_pie
  } else if (!is.null(nota_pie_derecha) && nzchar(nota_pie_derecha)) {
    caption_text <- nota_pie_derecha
  }

  base_theme <- ggplot2::theme_minimal(base_size = 9) +
    ggplot2::theme(
      panel.grid.major.y = ggplot2::element_blank(),
      panel.grid.minor   = ggplot2::element_blank(),
      panel.grid.major.x = ggplot2::element_blank(),
      axis.title.x       = ggplot2::element_blank(),
      axis.title.y       = ggplot2::element_blank(),
      axis.text.y        = ggplot2::element_text(
        color = color_ejes,
        size  = size_ejes,
        hjust = 0.5,
        vjust = 0.5
      ),
      axis.line.y        = ggplot2::element_line(color = color_ejes, linewidth = 0.3),
      axis.text.x        = ggplot2::element_blank(),
      axis.ticks.x       = ggplot2::element_blank(),
      axis.line.x        = ggplot2::element_blank(),
      legend.title       = ggplot2::element_blank(),
      legend.position    = if (mostrar_leyenda) "bottom" else "none",
      legend.text        = ggplot2::element_text(
        color = color_leyenda,
        size  = size_leyenda,
        face  = if ("leyenda" %in% textos_negrita) "bold" else "plain"
      ),
      plot.margin        = ggplot2::margin(t = 15, r = 80, b = 15, l = 5),
      plot.title         = ggplot2::element_text(
        hjust = hjust_titulo,
        color = color_titulo,
        size  = size_titulo,
        face  = if ("titulo" %in% textos_negrita) "bold" else "plain"
      ),
      plot.subtitle      = ggplot2::element_text(
        hjust = hjust_titulo,
        color = color_subtitulo,
        size  = size_subtitulo
      ),
      plot.caption       = ggplot2::element_text(
        hjust = hjust_caption,
        color = color_nota_pie,
        size  = size_nota_pie
      ),
      plot.background    = ggplot2::element_rect(fill = color_fondo, color = NA),
      panel.background   = ggplot2::element_rect(fill = color_fondo, color = NA)
    )

  if (escala_valor %in% c("proporcion_1", "proporcion_100")) {
    eje_theme <- ggplot2::theme(
      axis.text.x  = ggplot2::element_text(
        color = "#7F7F7F",
        size  = size_ejes
      ),
      axis.ticks.x = ggplot2::element_line(
        color = "#7F7F7F",
        linewidth = 0.3
      ),
      axis.line.x  = ggplot2::element_line(
        color = "#7F7F7F",
        linewidth = 0.4
      )
    )
  } else {
    eje_theme <- ggplot2::theme()
  }

  p <- p +
    ggplot2::coord_flip() +
    base_theme +
    eje_theme +
    ggplot2::labs(
      title    = titulo,
      subtitle = subtitulo,
      caption  = caption_text
    )

  if (invertir_leyenda && mostrar_leyenda) {
    p <- p + ggplot2::guides(fill = ggplot2::guide_legend(reverse = TRUE))
  }

  # ---------------------------------------------------------------------------
  # 7. Exportación
  # ---------------------------------------------------------------------------
  if (exportar == "rplot") {
    return(p)
  }

  if (is.null(path_salida) || !nzchar(path_salida)) {
    stop("Debe especificar `path_salida` cuando `exportar` no es 'rplot'.", call. = FALSE)
  }

  if (exportar == "word") {
    if (!requireNamespace("officer", quietly = TRUE)) {
      stop(
        "Para exportar a Word se requiere el paquete 'officer'.",
        call. = FALSE
      )
    }

    width_word  <- if (!missing(ancho) && !is.null(ancho)) ancho else 6.5
    height_word <- if (!missing(alto)  && !is.null(alto))  alto  else 4.5

    doc <- officer::read_docx()
    doc <- officer::body_add_gg(
      doc,
      value  = p,
      width  = width_word,
      height = height_word,
      style  = "centered"
    )
    print(doc, target = path_salida)

    return(invisible(p))
  }

  if (exportar == "png") {
    ggplot2::ggsave(
      filename = path_salida,
      plot     = p,
      width    = ancho,
      height   = alto,
      dpi      = dpi,
      bg       = if (is.na(color_fondo)) "transparent" else color_fondo
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
