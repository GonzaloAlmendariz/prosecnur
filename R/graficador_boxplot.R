#' Graficar boxplots profesionales por categoria
#'
#' Construye un grafico de distribucion tipo boxplot para una variable numerica
#' segmentada por categoria.
#'
#' Incluye soporte para:
#' - puntos jitter opcionales,
#' - marcador de media,
#' - etiquetas de base por categoria,
#' - composicion en canvas (titulo/panel/leyenda/pie),
#' - exportacion a PNG/PPT/Word.
#'
#' @param data `data.frame` o `tibble` con las columnas de categoria y valor.
#' @param var_categoria Nombre (string) de la columna categorica.
#' @param var_valor Nombre (string) de la columna numerica.
#' @param orientacion `"vertical"` o `"horizontal"`.
#' @param colores_categorias Vector de colores por categoria (opcional).
#' @param mostrar_puntos Si `TRUE`, superpone puntos jitter.
#' @param alpha_puntos,size_puntos,jitter_width,jitter_height Estilo de puntos.
#' @param mostrar_media Si `TRUE`, dibuja la media como marcador.
#' @param color_media,size_media Estilo del marcador de media.
#' @param mostrar_n_por_categoria Si `TRUE`, imprime base por categoria.
#' @param prefijo_n,size_n,color_n Estilo de texto de base.
#' @param titulo,subtitulo,nota_pie Textos del grafico.
#' @param pos_titulo,pos_nota_pie Alineacion horizontal de titulo y pie.
#' @param color_titulo,size_titulo,color_subtitulo,size_subtitulo
#'   Estilo de encabezado.
#' @param color_nota_pie,size_nota_pie Estilo de pie.
#' @param color_ejes,size_ejes Estilo de ejes.
#' @param color_fondo Color de fondo.
#' @param ancho_caja,tamano_linea_caja Estilo de cajas.
#' @param mostrar_leyenda Si `TRUE`, muestra leyenda por categoria.
#' @param invertir_barras Si `TRUE`, invierte orden de categorias.
#' @param ancho_max_eje_cat Ancho de wrap para etiquetas de categoria.
#'
#' @param usar_canvas Si `TRUE`, compone titulo/panel/leyenda/pie con `cowplot`.
#' @param canvas_h_title,canvas_h_legend,canvas_h_caption,canvas_pad_top
#'   Proporciones relativas del canvas.
#'
#' @param debug_ph_bordes Si `TRUE`, dibuja borde de depuracion.
#' @param debug_ph_col,debug_ph_lwd Color y grosor del borde de depuracion.
#'
#' @param exportar Tipo de salida: `"rplot"`, `"png"`, `"ppt"` o `"word"`.
#' @param path_salida Ruta de salida cuando `exportar != "rplot"`.
#' @param ancho,alto,dpi Dimensiones y resolucion de exportacion.
#' @param ppt_append Si `TRUE` y `path_salida` existe, agrega slide al PPT.
#' @param ppt_layout,ppt_master Layout/master al exportar a PPT.
#'
#' @return Si `exportar = "rplot"`, devuelve el objeto grafico. En otros casos,
#'   exporta y retorna invisiblemente el grafico.
#'
#' @family graficador
#' @export
graficar_boxplot <- function(
    data,
    var_categoria,
    var_valor,
    orientacion       = c("vertical", "horizontal"),
    colores_categorias = NULL,
    mostrar_puntos    = TRUE,
    alpha_puntos      = 0.28,
    size_puntos       = 1.35,
    jitter_width      = 0.15,
    jitter_height     = 0,
    mostrar_media     = TRUE,
    color_media       = "#173B63",
    size_media        = 2.3,
    mostrar_n_por_categoria = FALSE,
    prefijo_n         = "N=",
    size_n            = 2.8,
    color_n           = "#3F5F80",
    titulo            = NULL,
    subtitulo         = NULL,
    nota_pie          = NULL,
    pos_titulo        = c("izquierda", "centro", "derecha"),
    pos_nota_pie      = c("derecha", "izquierda", "centro"),
    color_titulo      = "#173B63",
    size_titulo       = 11,
    color_subtitulo   = "#173B63",
    size_subtitulo    = 9,
    color_nota_pie    = "#173B63",
    size_nota_pie     = 8,
    color_ejes        = "#173B63",
    size_ejes         = 9,
    color_fondo       = NA,
    ancho_caja        = 0.62,
    tamano_linea_caja = 0.45,
    mostrar_leyenda   = FALSE,
    invertir_barras   = FALSE,
    ancho_max_eje_cat = NULL,
    usar_canvas       = TRUE,
    canvas_h_title    = 0.14,
    canvas_h_legend   = 0.08,
    canvas_h_caption  = 0.06,
    canvas_pad_top    = 0.01,
    debug_ph_bordes   = FALSE,
    debug_ph_col      = "#FF00FF",
    debug_ph_lwd      = 0.8,
    exportar          = c("rplot", "png", "ppt", "word"),
    path_salida       = NULL,
    ancho             = 10,
    alto              = 6,
    dpi               = 300,
    ppt_append        = TRUE,
    ppt_layout        = "Blank",
    ppt_master        = "Office Theme"
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y

  .hjust_from_pos <- function(x) switch(x, izquierda = 0, centro = 0.5, derecha = 1, 0.5)

  .mk_text_panel <- function(top = NULL, bottom = NULL, pos = "izquierda",
                             col_top = "#000000", col_bottom = "#000000",
                             size_top = 11, size_bottom = 9) {
    if (!requireNamespace("cowplot", quietly = TRUE)) {
      stop("Para `usar_canvas=TRUE` se requiere 'cowplot'.", call. = FALSE)
    }
    h <- .hjust_from_pos(pos)
    x <- h
    p <- cowplot::ggdraw()
    if (!is.null(top) && nzchar(trimws(as.character(top)[1]))) {
      p <- p + cowplot::draw_label(
        label = as.character(top)[1],
        x = x, y = 0.68, hjust = h, vjust = 0.5,
        fontface = "bold", colour = col_top, size = size_top
      )
    }
    if (!is.null(bottom) && nzchar(trimws(as.character(bottom)[1]))) {
      p <- p + cowplot::draw_label(
        label = as.character(bottom)[1],
        x = x, y = 0.26, hjust = h, vjust = 0.5,
        fontface = "plain", colour = col_bottom, size = size_bottom
      )
    }
    p
  }

  .mk_palette <- function(levels_cat, pal_user = NULL) {
    levels_cat <- as.character(levels_cat)
    levels_cat <- levels_cat[!is.na(levels_cat) & nzchar(trimws(levels_cat))]
    if (!length(levels_cat)) return(character(0))

    base_pal <- c(
      "#0B4F8C", "#2A9D8F", "#E9C46A", "#F4A261",
      "#E76F51", "#7A9E9F", "#6D597A", "#5B8DEF"
    )

    if (is.null(pal_user) || !length(pal_user)) {
      if (length(levels_cat) <= length(base_pal)) {
        vals <- base_pal[seq_along(levels_cat)]
      } else {
        vals <- scales::hue_pal(h = c(200, 360), c = 70, l = 55)(length(levels_cat))
      }
      return(stats::setNames(vals, levels_cat))
    }

    pal_user <- as.character(pal_user)
    if (!is.null(names(pal_user))) {
      names(pal_user) <- trimws(as.character(names(pal_user)))
      vals <- pal_user[levels_cat]
      miss <- is.na(vals) | !nzchar(vals)
      if (any(miss)) {
        fallback <- setdiff(base_pal, vals[!miss])
        if (!length(fallback)) {
          fallback <- scales::hue_pal(h = c(200, 360), c = 70, l = 55)(sum(miss))
        } else if (length(fallback) < sum(miss)) {
          fallback <- c(
            fallback,
            scales::hue_pal(h = c(200, 360), c = 70, l = 55)(sum(miss) - length(fallback))
          )
        }
        vals[miss] <- fallback[seq_len(sum(miss))]
      }
      return(stats::setNames(vals, levels_cat))
    }

    if (length(pal_user) < length(levels_cat)) {
      extra <- scales::hue_pal(h = c(200, 360), c = 70, l = 55)(length(levels_cat) - length(pal_user))
      pal_user <- c(pal_user, extra)
    }
    stats::setNames(pal_user[seq_along(levels_cat)], levels_cat)
  }

  orientacion <- match.arg(orientacion)
  pos_titulo  <- match.arg(pos_titulo)
  pos_nota_pie <- match.arg(pos_nota_pie)
  exportar <- match.arg(exportar)

  if (!requireNamespace("ggplot2", quietly = TRUE) ||
      !requireNamespace("dplyr", quietly = TRUE) ||
      !requireNamespace("scales", quietly = TRUE)) {
    stop("Se requieren 'ggplot2', 'dplyr' y 'scales'.", call. = FALSE)
  }

  if (isTRUE(usar_canvas) && !requireNamespace("cowplot", quietly = TRUE)) {
    stop("Para `usar_canvas=TRUE` se requiere 'cowplot'.", call. = FALSE)
  }

  if (!is.data.frame(data)) stop("`data` debe ser data.frame/tibble.", call. = FALSE)
  if (!var_categoria %in% names(data)) stop("`var_categoria` no existe en `data`.", call. = FALSE)
  if (!var_valor %in% names(data)) stop("`var_valor` no existe en `data`.", call. = FALSE)

  df <- data |>
    dplyr::select(
      categoria = dplyr::all_of(var_categoria),
      valor = dplyr::all_of(var_valor)
    ) |>
    dplyr::mutate(
      categoria = as.character(.data$categoria),
      valor = suppressWarnings(as.numeric(.data$valor))
    ) |>
    dplyr::filter(
      !is.na(.data$categoria),
      nzchar(trimws(.data$categoria)),
      is.finite(.data$valor)
    )

  if (!nrow(df)) stop("No hay datos numericos validos para boxplot.", call. = FALSE)

  lvls <- unique(df$categoria)
  if (isTRUE(invertir_barras)) lvls <- rev(lvls)
  df$categoria <- factor(df$categoria, levels = lvls)

  if (!is.null(ancho_max_eje_cat) &&
      is.finite(ancho_max_eje_cat) &&
      ancho_max_eje_cat > 1 &&
      requireNamespace("stringr", quietly = TRUE)) {
    lvls_lab <- as.character(levels(df$categoria))
    levels(df$categoria) <- stringr::str_wrap(lvls_lab, width = as.integer(ancho_max_eje_cat))
  }

  pal <- .mk_palette(levels(df$categoria), pal_user = colores_categorias)

  p_core <- ggplot2::ggplot(
    df,
    ggplot2::aes(x = .data$categoria, y = .data$valor, fill = .data$categoria)
  ) +
    ggplot2::geom_boxplot(
      outlier.shape = NA,
      width = ancho_caja,
      alpha = 0.88,
      colour = "#1A3552",
      linewidth = tamano_linea_caja,
      show.legend = isTRUE(mostrar_leyenda)
    )

  if (isTRUE(mostrar_puntos)) {
    p_core <- p_core +
      ggplot2::geom_jitter(
        ggplot2::aes(colour = .data$categoria),
        width = jitter_width,
        height = jitter_height,
        alpha = alpha_puntos,
        size = size_puntos,
        stroke = 0,
        show.legend = FALSE
      )
  }

  if (isTRUE(mostrar_media)) {
    p_core <- p_core +
      ggplot2::stat_summary(
        fun = mean,
        geom = "point",
        shape = 23,
        fill = "white",
        colour = color_media,
        size = size_media,
        stroke = 0.8
      )
  }

  if (isTRUE(mostrar_n_por_categoria)) {
    rg <- range(df$valor, na.rm = TRUE)
    span <- rg[2] - rg[1]
    if (!is.finite(span) || span <= 0) span <- max(abs(rg[2]), 1)
    n_df <- df |>
      dplyr::group_by(.data$categoria) |>
      dplyr::summarise(
        n = dplyr::n(),
        y_n = max(.data$valor, na.rm = TRUE) + (span * 0.06),
        .groups = "drop"
      ) |>
      dplyr::mutate(label_n = paste0(prefijo_n, .data$n))

    p_core <- p_core +
      ggplot2::geom_text(
        data = n_df,
        ggplot2::aes(x = .data$categoria, y = .data$y_n, label = .data$label_n),
        inherit.aes = FALSE,
        colour = color_n,
        size = size_n
      )
  }

  p_core <- p_core +
    ggplot2::scale_fill_manual(values = pal, drop = FALSE) +
    ggplot2::scale_colour_manual(values = pal, drop = FALSE) +
    ggplot2::theme_minimal(base_size = 10) +
    ggplot2::theme(
      axis.title = ggplot2::element_blank(),
      axis.text.x = ggplot2::element_text(colour = color_ejes, size = size_ejes),
      axis.text.y = ggplot2::element_text(colour = color_ejes, size = size_ejes),
      panel.grid.minor = ggplot2::element_blank(),
      panel.grid.major.x = ggplot2::element_blank(),
      panel.grid.major.y = ggplot2::element_line(colour = "#D9E2EC", linewidth = 0.35),
      legend.title = ggplot2::element_blank(),
      legend.position = if (isTRUE(mostrar_leyenda)) "bottom" else "none",
      legend.text = ggplot2::element_text(colour = color_ejes, size = max(7, size_ejes - 1)),
      plot.background = ggplot2::element_rect(fill = color_fondo, colour = NA),
      panel.background = ggplot2::element_rect(fill = color_fondo, colour = NA),
      plot.margin = ggplot2::margin(2, 6, 2, 4)
    )

  if (identical(orientacion, "horizontal")) {
    p_core <- p_core + ggplot2::coord_flip()
  }

  p_out <- p_core

  if (isTRUE(usar_canvas)) {
    .has_title <- (!is.null(titulo) && nzchar(trimws(as.character(titulo)[1]))) ||
      (!is.null(subtitulo) && nzchar(trimws(as.character(subtitulo)[1])))
    .has_caption <- !is.null(nota_pie) && nzchar(trimws(as.character(nota_pie)[1]))
    .has_legend <- isTRUE(mostrar_leyenda)

    h_title <- if (.has_title) canvas_h_title else 0
    h_caption <- if (.has_caption) canvas_h_caption else 0
    h_legend <- if (.has_legend) canvas_h_legend else 0
    h_pad <- max(0, canvas_pad_top)
    h_panel <- max(0.20, 1 - (h_title + h_legend + h_caption + h_pad))

    pieces <- list()
    relh <- numeric(0)

    if (h_pad > 0) {
      pieces[[length(pieces) + 1]] <- cowplot::ggdraw()
      relh <- c(relh, h_pad)
    }

    if (.has_title) {
      pieces[[length(pieces) + 1]] <- .mk_text_panel(
        top = titulo, bottom = subtitulo, pos = pos_titulo,
        col_top = color_titulo, col_bottom = color_subtitulo,
        size_top = size_titulo, size_bottom = size_subtitulo
      )
      relh <- c(relh, h_title)
    }

    if (.has_legend) {
      leg <- cowplot::get_legend(
        p_core +
          ggplot2::theme(
            legend.position = "bottom",
            legend.margin = ggplot2::margin(0, 0, 0, 0),
            legend.box.margin = ggplot2::margin(0, 0, 0, 0)
          )
      )
      p_panel <- p_core + ggplot2::theme(legend.position = "none")
      pieces[[length(pieces) + 1]] <- p_panel
      relh <- c(relh, h_panel)
      pieces[[length(pieces) + 1]] <- cowplot::ggdraw(leg)
      relh <- c(relh, h_legend)
    } else {
      pieces[[length(pieces) + 1]] <- p_core
      relh <- c(relh, h_panel)
    }

    if (.has_caption) {
      pieces[[length(pieces) + 1]] <- .mk_text_panel(
        top = nota_pie, bottom = NULL, pos = pos_nota_pie,
        col_top = color_nota_pie, col_bottom = color_nota_pie,
        size_top = size_nota_pie, size_bottom = size_nota_pie
      )
      relh <- c(relh, h_caption)
    }

    p_out <- cowplot::plot_grid(plotlist = pieces, ncol = 1, rel_heights = relh, align = "v")
  }

  if (isTRUE(debug_ph_bordes)) {
    p_out <- p_out +
      ggplot2::theme(
        plot.background = ggplot2::element_rect(
          fill = NA, colour = debug_ph_col, linewidth = debug_ph_lwd
        )
      )
  }

  if (identical(exportar, "rplot")) return(p_out)

  if (is.null(path_salida) || !nzchar(path_salida)) {
    path_salida <- switch(
      exportar,
      png = "boxplot.png",
      ppt = "boxplot.pptx",
      word = "boxplot.docx",
      "boxplot_output"
    )
  }

  if (identical(exportar, "png")) {
    ggplot2::ggsave(
      filename = path_salida,
      plot = p_out,
      width = ancho,
      height = alto,
      dpi = dpi,
      bg = if (is.na(color_fondo)) "white" else color_fondo
    )
    return(invisible(p_out))
  }

  if (!requireNamespace("officer", quietly = TRUE)) {
    stop("Para exportar a PPT/Word se requiere 'officer'.", call. = FALSE)
  }

  if (identical(exportar, "ppt")) {
    if (!requireNamespace("rvg", quietly = TRUE)) {
      stop("Para exportar a PPT se requiere 'rvg'.", call. = FALSE)
    }
    doc <- if (isTRUE(ppt_append) && file.exists(path_salida)) {
      officer::read_pptx(path_salida)
    } else {
      officer::read_pptx()
    }
    doc <- officer::add_slide(doc, layout = ppt_layout, master = ppt_master)
    doc <- officer::ph_with(
      doc,
      value = rvg::dml(ggobj = p_out, bg = "transparent"),
      location = officer::ph_location_fullsize()
    )
    print(doc, target = path_salida)
    return(invisible(p_out))
  }

  doc <- if (file.exists(path_salida)) officer::read_docx(path = path_salida) else officer::read_docx()
  doc <- officer::body_add_gg(doc, value = p_out, width = ancho, height = alto, style = "Normal")
  print(doc, target = path_salida)
  invisible(p_out)
}
