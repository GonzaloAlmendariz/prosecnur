# =============================================================================
# Graficadores de Dimensiones
# - Heatmap semafórico en canvas
# - Wrapper radar/barras con payloads de dimensiones
# - Wrapper radar+tabla, sin tablas nativas de PowerPoint
# =============================================================================

#' @keywords internal
.dim_wrap_debug_canvas <- function(g, debug_ph_bordes = FALSE, debug_ph_col = "#FF00FF", debug_ph_lwd = 0.6) {
  if (!isTRUE(debug_ph_bordes)) return(g)
  cowplot::ggdraw(g) +
    cowplot::draw_grob(
      grid::rectGrob(
        gp = grid::gpar(col = debug_ph_col, fill = NA, lwd = debug_ph_lwd)
      ),
      x = 0, y = 0, width = 1, height = 1
    )
}

#' @keywords internal
.dim_export_canvas <- function(
    canvas,
    exportar = c("rplot", "png", "ppt", "word"),
    path_salida = NULL,
    ancho = 8.5,
    alto = 6.0,
    dpi = 300,
    ppt_append = TRUE,
    ppt_layout = "Blank",
    ppt_master = "Office Theme"
) {
  exportar <- match.arg(exportar)

  if (exportar == "rplot") return(canvas)

  if (is.null(path_salida) || !nzchar(path_salida)) {
    stop("`path_salida` es requerido para exportar.", call. = FALSE)
  }

  if (exportar == "png") {
    ggplot2::ggsave(
      filename = path_salida,
      plot = canvas,
      width = ancho,
      height = alto,
      units = "in",
      dpi = dpi,
      bg = "transparent"
    )
    return(invisible(canvas))
  }

  if (exportar %in% c("ppt", "word")) {
    if (!requireNamespace("officer", quietly = TRUE)) stop("Para exportar a PPT/Word se requiere officer.", call. = FALSE)
    if (!requireNamespace("rvg", quietly = TRUE)) stop("Para exportar a PPT/Word se requiere rvg.", call. = FALSE)

    if (exportar == "ppt") {
      doc <- if (ppt_append && file.exists(path_salida)) officer::read_pptx(path_salida) else officer::read_pptx()
      doc <- officer::add_slide(doc, layout = ppt_layout, master = ppt_master)
      doc <- officer::ph_with(
        doc,
        value = rvg::dml(ggobj = canvas),
        location = officer::ph_location_fullsize()
      )
      print(doc, target = path_salida)
      return(invisible(canvas))
    }

    doc <- if (file.exists(path_salida)) officer::read_docx(path_salida) else officer::read_docx()
    doc <- officer::body_add_par(doc, value = "", style = "Normal")
    doc <- officer::body_add_dml(
      doc,
      value = rvg::dml(ggobj = canvas),
      width = ancho,
      height = alto
    )
    print(doc, target = path_salida)
    return(invisible(canvas))
  }

  stop("Tipo de exportación no soportado.", call. = FALSE)
}

#' @keywords internal
.dim_blank_canvas <- function(
    mensaje = "Sin datos para mostrar",
    debug_ph_bordes = FALSE,
    debug_ph_col = "#FF00FF",
    debug_ph_lwd = 0.6
) {
  .dim_wrap_debug_canvas(
    cowplot::ggdraw() +
      cowplot::draw_label(
        label = mensaje,
        x = 0.5, y = 0.5,
        hjust = 0.5, vjust = 0.5,
        size = 12,
        colour = "#20324d"
      ),
    debug_ph_bordes = debug_ph_bordes,
    debug_ph_col = debug_ph_col,
    debug_ph_lwd = debug_ph_lwd
  )
}

#' @keywords internal
.dim_heat_legend_block <- function(labels, colors, size = 11, colour = "#004B8D") {
  labels <- as.character(labels)
  colors <- as.character(colors)
  n <- min(length(labels), length(colors))
  if (!n) return(cowplot::ggdraw() + cowplot::theme_nothing())

  labels <- labels[seq_len(n)]
  colors <- colors[seq_len(n)]
  width_guess <- pmax(nchar(labels, type = "width"), 8)
  item_units <- 0.045 + (width_guess * 0.012) + 0.04
  total_units <- sum(item_units)
  usable_width <- min(0.92, total_units)
  scale <- if (total_units > 0) usable_width / total_units else 1
  item_widths <- item_units * scale
  start_x <- max(0.03, (1 - sum(item_widths)) / 2)
  x_box <- numeric(n)
  x_text <- numeric(n)
  cur_x <- start_x

  g <- cowplot::ggdraw()
  for (i in seq_len(n)) {
    x_box[i] <- cur_x
    x_text[i] <- min(cur_x + 0.04, 0.96)
    g <- g +
      cowplot::draw_grob(
        grid::rectGrob(gp = grid::gpar(fill = colors[i], col = NA)),
        x = x_box[i], y = 0.5,
        width = 0.028, height = 0.22,
        hjust = 0, vjust = 0.5
      ) +
      cowplot::draw_label(
        label = labels[i],
        x = x_text[i], y = 0.5,
        hjust = 0, vjust = 0.5,
        size = size,
        colour = colour
      )
    cur_x <- cur_x + item_widths[i]
  }
  g
}

#' @keywords internal
.dim_payload_to_plot_df <- function(payload) {
  payload$score_plot |>
    dplyr::transmute(
      eje = as.character(.data$axis_label),
      grupo = as.character(.data$grupo),
      valor = as.numeric(.data$score_round),
      base = as.numeric(.data$base)
    )
}

#' @keywords internal
.dim_alias_radar_extra_args <- function(extra_args) {
  if (is.null(extra_args) || !is.list(extra_args) || !length(extra_args)) return(extra_args)

  alias_if_missing <- function(dst, src) {
    if (!is.null(extra_args[[src]]) && is.null(extra_args[[dst]])) {
      extra_args[[dst]] <<- extra_args[[src]]
    }
    extra_args[[src]] <<- NULL
  }

  alias_if_missing("canvas_h_header_in", "canvas_h_title")
  alias_if_missing("canvas_h_header_in", "canvas_h_header")
  alias_if_missing("canvas_h_legend_in", "canvas_h_legend")
  alias_if_missing("canvas_h_caption_in", "canvas_h_caption")
  extra_args
}

#' @keywords internal
.dim_payload_to_numeric_wide <- function(payload) {
  df_plot <- .dim_payload_to_plot_df(payload)
  grupos <- payload$group_order %||% unique(as.character(df_plot$grupo))
  if (!length(grupos)) grupos <- unique(as.character(df_plot$grupo))

  safe_name <- function(x) {
    x <- gsub("[^A-Za-z0-9]+", "_", as.character(x))
    x <- gsub("^_+|_+$", "", x)
    x <- gsub("_+", "_", x)
    paste0("serie_", ifelse(nzchar(x), x, "x"))
  }

  series_cols <- safe_name(grupos)
  make_unique <- function(x) {
    if (!length(x)) return(x)
    out <- x
    dup <- duplicated(out)
    if (any(dup)) {
      idx <- ave(seq_along(out), out, FUN = seq_along)
      out[dup] <- paste0(out[dup], "_", idx[dup])
    }
    out
  }
  series_cols <- make_unique(series_cols)
  map_cols <- stats::setNames(series_cols, grupos)

  wide <- df_plot |>
    dplyr::mutate(.serie_col = map_cols[.data$grupo]) |>
    dplyr::select(.data$eje, .data$.serie_col, .data$valor) |>
    tidyr::pivot_wider(names_from = ".serie_col", values_from = "valor")

  wide$categoria <- factor(
    wide$eje,
    levels = rev(payload$axis_order_plot %||% unique(as.character(wide$eje)))
  )
  wide <- wide[, c("categoria", series_cols), drop = FALSE]

  list(
    data = wide,
    vars_valor = series_cols,
    etiquetas_series = stats::setNames(grupos, series_cols)
  )
}

#' @keywords internal
.dim_make_table_df <- function(payload, titulo_left = "TOP TWO BOX", digits = 0L) {
  digits <- suppressWarnings(as.integer(digits))
  if (!is.finite(digits) || digits < 0L) digits <- 0L

  df_plot <- .dim_payload_to_plot_df(payload)
  ejes <- payload$axis_order_plot %||% unique(as.character(df_plot$eje))
  grupos <- payload$group_order %||% unique(as.character(df_plot$grupo))

  wide <- df_plot |>
    dplyr::transmute(
      eje = as.character(.data$eje),
      grupo = as.character(.data$grupo),
      valor = as.numeric(.data$valor)
    ) |>
    tidyr::complete(eje = ejes, grupo = grupos, fill = list(valor = 0)) |>
    tidyr::pivot_wider(names_from = "grupo", values_from = "valor")

  fmt_pct <- function(x) {
    x <- suppressWarnings(as.numeric(x))
    x[!is.finite(x) | is.na(x)] <- 0
    if (digits == 0L) sprintf("%.0f%%", x) else sprintf(paste0("%.", digits, "f%%"), x)
  }

  out <- as.data.frame(wide)
  out[[1]] <- as.character(out[[1]])
  for (j in 2:ncol(out)) out[[j]] <- fmt_pct(out[[j]])
  names(out)[1] <- as.character(titulo_left %||% "TOP TWO BOX")[1]
  out
}

#' @keywords internal
.dim_make_table_grob <- function(
    tb,
    header_fill = "#062A63",
    header_text = "white",
    body_fill = "#F2F2F2",
    grid_col = "white",
    text_blue = "#062A63",
    font_family = "Arial",
    header_size = 8,
    body_size = 7,
    firstcol_bold = TRUE,
    highlight_threshold = 60,
    highlight_col = "red",
    padding_mm = 3,
    firstcol_frac = 0.55,
    wrap_header = 14
) {
  if (!requireNamespace("gridExtra", quietly = TRUE)) stop("Requiere gridExtra.", call. = FALSE)

  n_data <- nrow(tb)
  n_cols <- ncol(tb)
  firstcol_frac <- suppressWarnings(as.numeric(firstcol_frac))
  if (!is.finite(firstcol_frac)) firstcol_frac <- 0.55
  firstcol_frac <- max(0.40, min(0.80, firstcol_frac))

  if (requireNamespace("stringr", quietly = TRUE) && is.finite(wrap_header) && wrap_header > 0) {
    nms <- names(tb)
    if (length(nms) >= 2) {
      nms[-1] <- stringr::str_wrap(nms[-1], width = as.integer(wrap_header))
      names(tb) <- nms
    }
  }

  tg <- gridExtra::tableGrob(
    tb,
    rows = NULL,
    theme = gridExtra::ttheme_minimal(
      base_size = body_size,
      base_family = font_family,
      padding = grid::unit(rep(padding_mm, 2), "mm"),
      colhead = list(
        fg_params = list(col = header_text, fontface = "bold"),
        bg_params = list(fill = header_fill, col = grid_col, lwd = 2)
      ),
      core = list(
        fg_params = list(col = text_blue),
        bg_params = list(fill = body_fill, col = grid_col, lwd = 2)
      )
    )
  )

  if (n_cols >= 2) {
    rest <- (1 - firstcol_frac) / (n_cols - 1)
    tg$widths <- grid::unit(c(firstcol_frac, rep(rest, n_cols - 1)), "npc")
  } else {
    tg$widths <- grid::unit(1, "npc")
  }

  for (j in seq_len(n_cols)) {
    k <- which(tg$layout$t == 1 & tg$layout$l == j & tg$layout$name == "colhead-fg")
    if (length(k)) {
      tg$grobs[[k]]$just <- "center"
      tg$grobs[[k]]$x <- grid::unit(0.5, "npc")
      tg$grobs[[k]]$gp <- grid::gpar(col = header_text, fontface = "bold", fontsize = header_size)
    }
  }

  for (i in seq_len(n_data)) {
    r <- i + 1
    k1 <- which(tg$layout$t == r & tg$layout$l == 1 & tg$layout$name == "core-fg")
    if (length(k1)) {
      tg$grobs[[k1]]$just <- "center"
      tg$grobs[[k1]]$x <- grid::unit(0.5, "npc")
      tg$grobs[[k1]]$y <- grid::unit(0.5, "npc")
      tg$grobs[[k1]]$gp <- grid::gpar(
        col = text_blue,
        fontface = if (isTRUE(firstcol_bold)) "bold" else "plain",
        fontsize = body_size,
        lineheight = 0.95
      )
    }

    if (n_cols >= 2) {
      for (j in 2:n_cols) {
        kj <- which(tg$layout$t == r & tg$layout$l == j & tg$layout$name == "core-fg")
        if (length(kj)) {
          tg$grobs[[kj]]$just <- "center"
          tg$grobs[[kj]]$x <- grid::unit(0.5, "npc")
          tg$grobs[[kj]]$y <- grid::unit(0.5, "npc")
          tg$grobs[[kj]]$gp <- grid::gpar(col = text_blue, fontface = "plain", fontsize = body_size)
        }
      }
    }
  }

  parse_pct <- function(x) suppressWarnings(as.numeric(gsub("%", "", x)))
  if (n_cols >= 2) {
    for (j in 2:n_cols) {
      vals <- parse_pct(tb[[j]])
      idx_low <- which(is.finite(vals) & !is.na(vals) & vals <= highlight_threshold)
      if (length(idx_low)) {
        for (ii in idx_low) {
          r <- ii + 1
          kj <- which(tg$layout$t == r & tg$layout$l == j & tg$layout$name == "core-fg")
          if (length(kj)) {
            tg$grobs[[kj]]$gp <- grid::gpar(col = highlight_col, fontface = "bold", fontsize = body_size)
            tg$grobs[[kj]]$just <- "center"
            tg$grobs[[kj]]$x <- grid::unit(0.5, "npc")
          }
        }
      }
    }
  }

  tg
}

#' @keywords internal
.dim_compose_plot_table_canvas <- function(
    plot_obj,
    table_grob,
    tabla_ph_ancho = 0.40,
    tabla_ph_gap = 0.03,
    tabla_auto_fit = FALSE,
    tabla_fit_pad = 0.98,
    tabla_allow_upscale = FALSE,
    debug_ph_bordes = FALSE,
    debug_ph_col = "#FF00FF",
    debug_ph_lwd = 0.6
) {
  tabla_ph_ancho <- suppressWarnings(as.numeric(tabla_ph_ancho))
  if (!is.finite(tabla_ph_ancho) || tabla_ph_ancho <= 0 || tabla_ph_ancho >= 0.8) tabla_ph_ancho <- 0.40
  tabla_ph_gap <- suppressWarnings(as.numeric(tabla_ph_gap))
  if (!is.finite(tabla_ph_gap) || tabla_ph_gap < 0 || tabla_ph_gap >= 0.2) tabla_ph_gap <- 0.03

  w_plot <- 1 - tabla_ph_ancho - tabla_ph_gap
  x_table <- w_plot + tabla_ph_gap

  scale_tab <- 1
  if (isTRUE(tabla_auto_fit)) {
    gw_in <- suppressWarnings(grid::convertWidth(sum(table_grob$widths), "in", valueOnly = TRUE))
    gh_in <- suppressWarnings(grid::convertHeight(sum(table_grob$heights), "in", valueOnly = TRUE))
    if (is.finite(gw_in) && gw_in > 0 && is.finite(gh_in) && gh_in > 0) {
      s_w <- tabla_ph_ancho / gw_in
      s_h <- 1 / gh_in
      scale_tab <- min(s_w, s_h)
      if (!isTRUE(tabla_allow_upscale)) scale_tab <- min(1, scale_tab)
      scale_tab <- scale_tab * tabla_fit_pad
      if (!is.finite(scale_tab) || scale_tab <= 0) scale_tab <- 1
    }
  }

  canvas <- cowplot::ggdraw() +
    cowplot::draw_plot(plot_obj, x = 0, y = 0, width = w_plot, height = 1) +
    cowplot::draw_grob(
      table_grob,
      x = x_table + (tabla_ph_ancho * 0.5),
      y = 0.5,
      width = tabla_ph_ancho,
      height = 1,
      hjust = 0.5,
      vjust = 0.5,
      scale = scale_tab
    )

  if (isTRUE(debug_ph_bordes)) {
    canvas <- canvas +
      cowplot::draw_grob(grid::rectGrob(gp = grid::gpar(col = debug_ph_col, fill = NA, lwd = debug_ph_lwd)), x = 0, y = 0, width = w_plot, height = 1) +
      cowplot::draw_grob(grid::rectGrob(gp = grid::gpar(col = debug_ph_col, fill = NA, lwd = debug_ph_lwd)), x = x_table, y = 0, width = tabla_ph_ancho, height = 1)
  }

  canvas
}

#' Heatmap semafórico de dimensiones en canvas
#'
#' @param data Base recodificada e indexada.
#' @param modo `"general"` o `"indicadores"`.
#' @param objetivo Id técnico del catálogo.
#' @param instrumento Instrumento opcional. Si es `NULL`, se usa `attr(data, "instrumento_reporte")`.
#' @param cruce Variable de comparación opcional.
#' @param incluir_total Si es `NULL`, usa el default de la configuración interna.
#' @param filtros Lista nombrada de filtros por variable.
#' @param iter_var Variable opcional de iteración.
#' @param iter_level Nivel opcional de iteración.
#' @param titulo,subtitulo,nota_pie Textos del gráfico.
#' @param usar_canvas Si `TRUE`, compone encabezado, panel, leyenda y pie con `cowplot`.
#' @param debug_ph_bordes,debug_ph_col,debug_ph_lwd Borde de depuración del canvas.
#' @param exportar Tipo de exportación.
#' @param path_salida Ruta de salida cuando `exportar != "rplot"`.
#' @param ancho,alto,dpi Tamaño y resolución.
#'
#' @return Objeto gráfico o exportación invisible.
#' @export
graficar_heatmap_dimensiones <- function(
    data,
    modo = c("general", "indicadores"),
    objetivo,
    instrumento = NULL,
    cruce = NULL,
    incluir_total = NULL,
    filtros = list(),
    iter_var = NULL,
    iter_level = NULL,
    titulo = NULL,
    subtitulo = NULL,
    nota_pie = NULL,
    color_titulo = "#004B8D",
    size_titulo = 12,
    color_subtitulo = "#004B8D",
    size_subtitulo = 9,
    color_nota_pie = "#004B8D",
    size_nota_pie = 8,
    color_leyenda = "#004B8D",
    size_leyenda = 9,
    color_ejes = "#20324d",
    size_ejes = 11,
    color_texto_celdas = "#122842",
    size_texto_celdas = 11,
    color_fondo = NA,
    angle_x = -18,
    mostrar_leyenda = TRUE,
    usar_canvas = TRUE,
    canvas_h_title = 0.15,
    canvas_h_legend = 0.10,
    canvas_h_caption = 0.08,
    canvas_pad_top = 0.01,
    debug_ph_bordes = FALSE,
    debug_ph_col = "#FF00FF",
    debug_ph_lwd = 0.6,
    exportar = c("rplot", "png", "ppt", "word"),
    path_salida = NULL,
    ancho = 8.5,
    alto = 5.6,
    dpi = 300,
    ppt_append = TRUE,
    ppt_layout = "Blank",
    ppt_master = "Office Theme"
) {
  if (!requireNamespace("ggplot2", quietly = TRUE)) stop("Requiere ggplot2.", call. = FALSE)
  if (!requireNamespace("cowplot", quietly = TRUE)) stop("Requiere cowplot.", call. = FALSE)

  modo <- match.arg(modo)
  exportar <- match.arg(exportar)
  mostrar_leyenda <- TRUE

  ctx <- .dim_build_context(data, instrumento = instrumento)
  payload <- .dim_build_payload(
    ctx,
    modo = modo,
    objetivo = objetivo,
    cruce = cruce,
    incluir_total = incluir_total,
    filtros = filtros,
    iter_var = iter_var,
    iter_level = iter_level
  )

  if (!nrow(payload$score_heat)) {
    return(.dim_export_canvas(
      .dim_blank_canvas(
        mensaje = "Sin datos para mostrar",
        debug_ph_bordes = debug_ph_bordes,
        debug_ph_col = debug_ph_col,
        debug_ph_lwd = debug_ph_lwd
      ),
      exportar = exportar,
      path_salida = path_salida,
      ancho = ancho,
      alto = alto,
      dpi = dpi,
      ppt_append = ppt_append,
      ppt_layout = ppt_layout,
      ppt_master = ppt_master
    ))
  }

  sem <- payload$semaforo
  cuts_lab <- .dim_range_labels(sem$cortes[1], sem$cortes[2])
  legend_breaks <- cuts_lab
  legend_limits <- c(cuts_lab[1], cuts_lab[2], cuts_lab[3], "Sin dato")
  sc <- payload$score_heat

  sc$grupo <- factor(sc$grupo, levels = payload$group_order %||% unique(as.character(sc$grupo)))
  sc$axis_label <- factor(sc$axis_label, levels = rev(payload$axis_order_heat %||% unique(as.character(sc$axis_label))))
  sc$estado <- dplyr::case_when(
    is.na(sc$score_raw) ~ "Sin dato",
    sc$score_raw < sem$cortes[1] ~ cuts_lab[1],
    sc$score_raw < sem$cortes[2] ~ cuts_lab[2],
    TRUE ~ cuts_lab[3]
  )
  sc$estado <- factor(sc$estado, levels = c(cuts_lab[1], cuts_lab[2], cuts_lab[3], "Sin dato"))
  sc$label <- ifelse(is.na(sc$score_raw), "", .dim_fmt_int(sc$score_round))

  max_chars <- max(nchar(as.character(payload$axis_order_heat), type = "width"), na.rm = TRUE)
  left_margin <- .dim_clamp(36 + 7 * max_chars, 130, 320)

  p_panel <- ggplot2::ggplot(
    sc,
    ggplot2::aes(x = .data$grupo, y = .data$axis_label, fill = .data$estado)
  ) +
    ggplot2::geom_tile(colour = "white", linewidth = 0.7) +
    ggplot2::geom_text(
      ggplot2::aes(label = .data$label),
      size = size_texto_celdas / 3,
      colour = color_texto_celdas,
      fontface = "bold"
    ) +
    ggplot2::scale_fill_manual(
      values = c(
        stats::setNames(sem$rojo, cuts_lab[1]),
        stats::setNames(sem$ambar, cuts_lab[2]),
        stats::setNames(sem$verde, cuts_lab[3]),
        "Sin dato" = sem$na
      ),
      limits = legend_limits,
      breaks = legend_breaks,
      drop = FALSE
    ) +
    ggplot2::theme_minimal(base_size = 11) +
    ggplot2::theme(
      panel.grid = ggplot2::element_blank(),
      axis.title = ggplot2::element_blank(),
      axis.text.x = ggplot2::element_text(size = size_ejes, colour = color_ejes, angle = angle_x, hjust = 0, vjust = 1),
      axis.text.y = ggplot2::element_text(size = size_ejes, colour = color_ejes),
      legend.title = ggplot2::element_blank(),
      legend.text = ggplot2::element_text(size = size_leyenda, colour = color_leyenda),
      legend.background = ggplot2::element_rect(fill = color_fondo, color = NA),
      legend.key = ggplot2::element_rect(fill = color_fondo, color = NA),
      plot.background = ggplot2::element_rect(fill = color_fondo, color = NA),
      panel.background = ggplot2::element_rect(fill = color_fondo, color = NA),
      plot.margin = ggplot2::margin(8, 12, 8, 8)
    )

  if (!isTRUE(usar_canvas)) {
    return(.dim_export_canvas(
      p_panel,
      exportar = exportar,
      path_salida = path_salida,
      ancho = ancho,
      alto = alto,
      dpi = dpi,
      ppt_append = ppt_append,
      ppt_layout = ppt_layout,
      ppt_master = ppt_master
    ))
  }

  p_panel <- p_panel + ggplot2::theme(legend.position = "none")

  title_block <- cowplot::ggdraw() +
    cowplot::draw_label(
      label = titulo %||% "",
      x = 0.5, y = if (!is.null(subtitulo) && nzchar(subtitulo)) 0.62 else 0.5,
      hjust = 0.5, vjust = 0.5,
      size = size_titulo,
      colour = color_titulo,
      fontface = "bold"
    ) +
    cowplot::draw_label(
      label = subtitulo %||% "",
      x = 0.5, y = if (!is.null(subtitulo) && nzchar(subtitulo)) 0.28 else 0.5,
      hjust = 0.5, vjust = 0.5,
      size = size_subtitulo,
      colour = color_subtitulo
    )

  legend_block <- if (isTRUE(mostrar_leyenda)) {
    .dim_heat_legend_block(
      labels = legend_breaks,
      colors = c(sem$rojo, sem$ambar, sem$verde),
      size = size_leyenda,
      colour = color_leyenda
    )
  } else {
    cowplot::ggdraw() + cowplot::theme_nothing()
  }
  caption_block <- cowplot::ggdraw() +
    cowplot::draw_label(
      label = nota_pie %||% "",
      x = 1, y = 0.5,
      hjust = 1, vjust = 0.5,
      size = size_nota_pie,
      colour = color_nota_pie
    )

  h_title <- canvas_h_title
  h_legend <- if (isTRUE(mostrar_leyenda)) canvas_h_legend else 0.01
  h_caption <- if (!is.null(nota_pie) && nzchar(nota_pie)) canvas_h_caption else 0.01
  h_panel <- max(0.01, 1 - (h_title + h_legend + h_caption) - canvas_pad_top)

  canvas <- cowplot::plot_grid(
    .dim_wrap_debug_canvas(title_block, debug_ph_bordes, debug_ph_col, debug_ph_lwd),
    .dim_wrap_debug_canvas(p_panel, debug_ph_bordes, debug_ph_col, debug_ph_lwd),
    .dim_wrap_debug_canvas(legend_block, debug_ph_bordes, debug_ph_col, debug_ph_lwd),
    .dim_wrap_debug_canvas(caption_block, debug_ph_bordes, debug_ph_col, debug_ph_lwd),
    ncol = 1,
    rel_heights = c(h_title, h_panel, h_legend, h_caption)
  )

  .dim_export_canvas(
    canvas,
    exportar = exportar,
    path_salida = path_salida,
    ancho = ancho,
    alto = alto,
    dpi = dpi,
    ppt_append = ppt_append,
    ppt_layout = ppt_layout,
    ppt_master = ppt_master
  )
}

#' Radar o barras de dimensiones en canvas
#'
#' @param data Base recodificada e indexada.
#' @param modo `"general"` o `"indicadores"`.
#' @param objetivo Id técnico del catálogo.
#' @param instrumento Instrumento opcional.
#' @param cruce Variable de comparación opcional.
#' @param incluir_total Si es `NULL`, usa el default interno.
#' @param filtros Lista nombrada de filtros por variable.
#' @param iter_var,iter_level Control opcional de iteración.
#' @param titulo,subtitulo,nota_pie Textos del gráfico.
#' @param ... Argumentos adicionales para `graficar_radar()` o `graficar_barras_numericas()`.
#'
#' @return Objeto gráfico o exportación invisible.
#' @export
graficar_radar_dimensiones <- function(
    data,
    modo = c("general", "indicadores"),
    objetivo,
    instrumento = NULL,
    cruce = NULL,
    incluir_total = NULL,
    filtros = list(),
    iter_var = NULL,
    iter_level = NULL,
    titulo = NULL,
    subtitulo = NULL,
    nota_pie = NULL,
    ...
) {
  modo <- match.arg(modo)
  ctx <- .dim_build_context(data, instrumento = instrumento)
  payload <- .dim_build_payload(
    ctx,
    modo = modo,
    objetivo = objetivo,
    cruce = cruce,
    incluir_total = incluir_total,
    filtros = filtros,
    iter_var = iter_var,
    iter_level = iter_level
  )

  if (!nrow(payload$score_plot)) {
    blank <- .dim_blank_canvas("Sin datos para mostrar")
    return(.dim_export_canvas(
      blank,
      exportar = extra_args$exportar %||% "rplot",
      path_salida = extra_args$path_salida %||% NULL,
      ancho = extra_args$ancho %||% 8.5,
      alto = extra_args$alto %||% 6.0,
      dpi = extra_args$dpi %||% 300,
      ppt_append = extra_args$ppt_append %||% TRUE,
      ppt_layout = extra_args$ppt_layout %||% "Blank",
      ppt_master = extra_args$ppt_master %||% "Office Theme"
    ))
  }

  extra_args <- list(...)
  extra_args <- .dim_alias_radar_extra_args(extra_args)

  if (identical(payload$visual_mode, "radar")) {
    if (!exists("graficar_radar", mode = "function", inherits = TRUE)) {
      stop("No existe `graficar_radar()`.", call. = FALSE)
    }

    df_plot <- .dim_payload_to_plot_df(payload)
    base_args <- list(
      data = df_plot,
      var_eje = "eje",
      var_grupo = "grupo",
      var_valor = "valor",
      escala_valor = "proporcion_100",
      colores_series = payload$group_colors,
      titulo = titulo,
      subtitulo = subtitulo,
      nota_pie = nota_pie,
      usar_canvas = TRUE,
      mostrar_radios = FALSE,
      mostrar_niveles = FALSE
    )

    args <- .merge_args(base_args, extra_args)
    args <- .keep_formals(graficar_radar, args)
    return(suppressWarnings(do.call(graficar_radar, args)))
  }

  if (!exists("graficar_barras_numericas", mode = "function", inherits = TRUE)) {
    stop("No existe `graficar_barras_numericas()`.", call. = FALSE)
  }

  wide <- .dim_payload_to_numeric_wide(payload)
  base_args <- list(
    data = wide$data,
    var_categoria = "categoria",
    vars_valor = wide$vars_valor,
    etiquetas_series = wide$etiquetas_series,
    orientacion = "horizontal",
    formato_valor = "numero",
    decimales = 0,
    colores_series = payload$group_colors,
    mostrar_n_sobre_barras = FALSE,
    titulo = titulo,
    subtitulo = subtitulo,
    nota_pie = nota_pie,
    usar_canvas = TRUE
  )

  args <- .merge_args(base_args, extra_args)
  args <- .keep_formals(graficar_barras_numericas, args)
  suppressWarnings(do.call(graficar_barras_numericas, args))
}

#' Radar + tabla de dimensiones en canvas
#'
#' @param data Base recodificada e indexada.
#' @param modo `"general"` o `"indicadores"`.
#' @param objetivo Id técnico del catálogo.
#' @param instrumento Instrumento opcional.
#' @param cruce Variable de comparación opcional.
#' @param incluir_total Si es `NULL`, usa el default interno.
#' @param filtros Lista nombrada de filtros por variable.
#' @param iter_var,iter_level Control opcional de iteración.
#' @param titulo,subtitulo,nota_pie Textos del gráfico.
#' @param titulo_tabla Título de la primera columna de la tabla.
#' @param ... Argumentos adicionales del radar, barras y tabla.
#'
#' @return Objeto gráfico o exportación invisible.
#' @export
graficar_radar_tabla_dimensiones <- function(
    data,
    modo = c("general", "indicadores"),
    objetivo,
    instrumento = NULL,
    cruce = NULL,
    incluir_total = NULL,
    filtros = list(),
    iter_var = NULL,
    iter_level = NULL,
    titulo = NULL,
    subtitulo = NULL,
    nota_pie = NULL,
    titulo_tabla = "TOP TWO BOX",
    ...
) {
  modo <- match.arg(modo)
  ctx <- .dim_build_context(data, instrumento = instrumento)
  payload <- .dim_build_payload(
    ctx,
    modo = modo,
    objetivo = objetivo,
    cruce = cruce,
    incluir_total = incluir_total,
    filtros = filtros,
    iter_var = iter_var,
    iter_level = iter_level
  )

  if (!nrow(payload$score_plot)) {
    blank <- .dim_blank_canvas("Sin datos para mostrar")
    return(.dim_export_canvas(
      blank,
      exportar = extra_args$exportar %||% "rplot",
      path_salida = extra_args$path_salida %||% NULL,
      ancho = extra_args$ancho %||% 8.5,
      alto = extra_args$alto %||% 6.0,
      dpi = extra_args$dpi %||% 300,
      ppt_append = extra_args$ppt_append %||% TRUE,
      ppt_layout = extra_args$ppt_layout %||% "Blank",
      ppt_master = extra_args$ppt_master %||% "Office Theme"
    ))
  }

  extra_args <- list(...)
  extra_args <- .dim_alias_radar_extra_args(extra_args)

  if (identical(payload$visual_mode, "radar")) {
    if (!exists("graficar_radar", mode = "function", inherits = TRUE)) {
      stop("No existe `graficar_radar()`.", call. = FALSE)
    }

    df_plot <- .dim_payload_to_plot_df(payload)
    base_args <- list(
      data = df_plot,
      var_eje = "eje",
      var_grupo = "grupo",
      var_valor = "valor",
      escala_valor = "proporcion_100",
      colores_series = payload$group_colors,
      titulo = titulo,
      subtitulo = subtitulo,
      nota_pie = nota_pie,
      usar_canvas = TRUE,
      mostrar_radios = FALSE,
      mostrar_niveles = FALSE,
      mostrar_tabla_derecha = TRUE,
      titulo_tabla = titulo_tabla
    )

    args <- .merge_args(base_args, extra_args)
    args <- .keep_formals(graficar_radar, args)
    return(suppressWarnings(do.call(graficar_radar, args)))
  }

  if (!exists("graficar_barras_numericas", mode = "function", inherits = TRUE)) {
    stop("No existe `graficar_barras_numericas()`.", call. = FALSE)
  }

  wide <- .dim_payload_to_numeric_wide(payload)
  args_bars <- .merge_args(
    list(
      data = wide$data,
      var_categoria = "categoria",
      vars_valor = wide$vars_valor,
      etiquetas_series = wide$etiquetas_series,
      orientacion = "horizontal",
      formato_valor = "numero",
      decimales = 0,
      colores_series = payload$group_colors,
      mostrar_n_sobre_barras = FALSE,
      titulo = titulo,
      subtitulo = subtitulo,
      nota_pie = NULL,
      usar_canvas = TRUE,
      exportar = "rplot"
    ),
    extra_args
  )
  args_bars <- .keep_formals(graficar_barras_numericas, args_bars)
  p_bars <- suppressWarnings(do.call(graficar_barras_numericas, args_bars))

  tb <- .dim_make_table_df(
    payload,
    titulo_left = titulo_tabla,
    digits = extra_args$tabla_digits %||% 0L
  )
  tg <- .dim_make_table_grob(
    tb,
    header_fill = extra_args$tabla_header_fill %||% "#062A63",
    body_fill = extra_args$tabla_body_fill %||% "#F2F2F2",
    grid_col = extra_args$tabla_grid_col %||% "white",
    text_blue = extra_args$tabla_text_blue %||% "#062A63",
    font_family = extra_args$tabla_font_family %||% "Arial",
    header_size = extra_args$tabla_header_size %||% 8,
    body_size = extra_args$tabla_body_size %||% 7,
    firstcol_bold = extra_args$tabla_firstcol_bold %||% TRUE,
    highlight_threshold = extra_args$umbral_rojo_pct %||% 60,
    padding_mm = extra_args$tabla_padding_mm %||% 3,
    firstcol_frac = extra_args$tabla_firstcol_frac %||% 0.55,
    wrap_header = extra_args$tabla_wrap_header %||% 14
  )

  canvas <- .dim_compose_plot_table_canvas(
    p_bars,
    tg,
    tabla_ph_ancho = extra_args$tabla_ph_ancho %||% 0.40,
    tabla_ph_gap = extra_args$tabla_ph_gap %||% 0.03,
    tabla_auto_fit = extra_args$tabla_auto_fit %||% FALSE,
    tabla_fit_pad = extra_args$tabla_fit_pad %||% 0.98,
    tabla_allow_upscale = extra_args$tabla_allow_upscale %||% FALSE,
    debug_ph_bordes = extra_args$debug_ph_bordes %||% FALSE,
    debug_ph_col = extra_args$debug_ph_col %||% "#FF00FF",
    debug_ph_lwd = extra_args$debug_ph_lwd %||% 0.6
  )

  if (!is.null(nota_pie) && nzchar(nota_pie)) {
    canvas <- cowplot::plot_grid(
      canvas,
      cowplot::ggdraw() +
        cowplot::draw_label(
          label = nota_pie,
          x = 1, y = 0.5,
          hjust = 1, vjust = 0.5,
          size = extra_args$size_nota_pie %||% 8,
          colour = extra_args$color_nota_pie %||% "#004B8D"
        ),
      ncol = 1,
      rel_heights = c(1, 0.08)
    )
  }

  if (identical(extra_args$exportar %||% "rplot", "rplot")) return(canvas)
  .dim_export_canvas(
    canvas,
    exportar = extra_args$exportar %||% "rplot",
    path_salida = extra_args$path_salida %||% NULL,
    ancho = extra_args$ancho %||% 8.5,
    alto = extra_args$alto %||% 6.0,
    dpi = extra_args$dpi %||% 300,
    ppt_append = extra_args$ppt_append %||% TRUE,
    ppt_layout = extra_args$ppt_layout %||% "Blank",
    ppt_master = extra_args$ppt_master %||% "Office Theme"
  )
}
