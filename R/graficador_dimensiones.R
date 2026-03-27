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
#' Visualiza indices y subindices de dimensiones como heatmap semafórico (rojo /
#' ámbar / verde). Requiere que `data` sea la salida encadenada de
#' [reporte_dimensiones()] y [reporte_dimensiones_indices()].
#'
#' @param data Base recodificada e indexada; salida de
#'   `reporte_dimensiones() |> reporte_dimensiones_indices()`.
#' @param modo `"general"` (índices por subindices) o `"indicadores"` (subindices por ítem).
#' @param objetivo Id técnico del catálogo (columna `idx_*` o clave de bloque).
#' @param instrumento Instrumento opcional. Si es `NULL`, se usa `attr(data, "instrumento_reporte")`.
#' @param cruce Variable de comparación opcional (columna en `data`).
#' @param incluir_total Si es `NULL`, usa el default de la configuración interna.
#' @param brecha_filas,brecha_cols Si `TRUE`, agrega fila/columna de brecha
#'   calculada como `(max - min)` en la matriz del heatmap.
#' @param etiq_brecha_filas,etiq_brecha_cols Etiquetas para fila/columna de brecha.
#' @param aplicar_gradiente_brecha Si `TRUE`, colorea celdas de brecha con gradiente.
#' @param brecha_colores Vector de colores para brecha (`bajo`, `alto`).
#' @param brecha_cortes Cortes numéricos (mín, máx) para escalar el gradiente de brecha.
#' @param filtros Lista nombrada de filtros por variable.
#' @param iter_var Variable opcional de iteración (columna en `data`).
#' @param iter_level Nivel específico de iteración.
#' @param titulo,subtitulo,nota_pie Textos del gráfico.
#' @param size_ejes_x Tamaño opcional de etiquetas del eje X. Si es `NULL`, usa `size_ejes`.
#' @param titulo_total_x,titulo_total_y Etiquetas para los totales en eje X y eje Y.
#' @param mostrar_n_cruce_x Si `TRUE`, agrega `(N=...)` en etiquetas del eje X por cruce.
#' @param usar_canvas Si `TRUE`, compone encabezado, panel, leyenda y pie con `cowplot`.
#' @param debug_ph_bordes,debug_ph_col,debug_ph_lwd Borde de depuración del canvas.
#' @param exportar Tipo de exportación: `"rplot"`, `"png"`, `"ppt"` o `"word"`.
#' @param path_salida Ruta de salida cuando `exportar != "rplot"`.
#' @param ancho,alto,dpi Tamaño y resolución de exportación.
#'
#' @return Objeto gráfico (canvas cowplot) o exportación invisible.
#' @family graficador
#' @seealso [reporte_dimensiones()], [reporte_dimensiones_indices()],
#'   [reporte_dimensiones_config()], [graficar_radar_dimensiones()]
#' @export
graficar_heatmap_dimensiones <- function(
    data,
    modo = c("general", "indicadores"),
    objetivo,
    instrumento = NULL,
    cruce = NULL,
    incluir_total = NULL,
    brecha_filas = FALSE,
    etiq_brecha_filas = "Brecha",
    brecha_cols = FALSE,
    etiq_brecha_cols = "Brecha",
    aplicar_gradiente_brecha = TRUE,
    brecha_colores = c(bajo = "#FFFFFF", alto = "#F4B183"),
    brecha_cortes = c(0, 30),
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
    size_ejes = 10,
    size_ejes_x = NULL,
    color_texto_celdas = "#122842",
    size_texto_celdas = 10,
    color_fondo = NA,
    angle_x = 0,
    titulo_total_x = "Total",
    titulo_total_y = "Total cruce",
    mostrar_n_cruce_x = FALSE,
    mostrar_leyenda = TRUE,
    usar_canvas = TRUE,
    canvas_h_title = 0.13,
    canvas_h_legend = 0.09,
    canvas_h_caption = 0.06,
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
  mostrar_leyenda <- isTRUE(mostrar_leyenda)
  mostrar_n_cruce_x <- isTRUE(mostrar_n_cruce_x)

  brecha_filas <- isTRUE(brecha_filas)
  brecha_cols <- isTRUE(brecha_cols)
  aplicar_gradiente_brecha <- isTRUE(aplicar_gradiente_brecha)

  size_ejes_x <- suppressWarnings(as.numeric(size_ejes_x))
  if (!length(size_ejes_x) || !is.finite(size_ejes_x[1]) || is.na(size_ejes_x[1]) || size_ejes_x[1] <= 0) {
    size_ejes_x <- size_ejes
  } else {
    size_ejes_x <- size_ejes_x[1]
  }

  titulo_total_x <- as.character(titulo_total_x %||% "Total")[1]
  if (!nzchar(trimws(titulo_total_x))) titulo_total_x <- "Total"
  titulo_total_y <- as.character(titulo_total_y %||% "Total cruce")[1]
  if (!nzchar(trimws(titulo_total_y))) titulo_total_y <- "Total cruce"

  etiq_brecha_filas <- as.character(etiq_brecha_filas %||% "Brecha")[1]
  if (!nzchar(trimws(etiq_brecha_filas))) etiq_brecha_filas <- "Brecha"
  etiq_brecha_cols <- as.character(etiq_brecha_cols %||% "Brecha")[1]
  if (!nzchar(trimws(etiq_brecha_cols))) etiq_brecha_cols <- "Brecha"

  brecha_colores <- as.character(brecha_colores)
  nmbc <- names(brecha_colores %||% character(0))
  if (is.null(nmbc)) nmbc <- character(0)
  col_brecha_bajo <- if ("bajo" %in% nmbc) brecha_colores[["bajo"]] else if (length(brecha_colores) >= 1L) brecha_colores[1] else "#FFFFFF"
  col_brecha_alto <- if ("alto" %in% nmbc) brecha_colores[["alto"]] else if (length(brecha_colores) >= 2L) brecha_colores[2] else "#F4B183"

  brecha_cortes <- suppressWarnings(as.numeric(brecha_cortes))
  brecha_cortes <- brecha_cortes[is.finite(brecha_cortes) & !is.na(brecha_cortes)]
  if (length(brecha_cortes) < 2L) brecha_cortes <- c(0, 30)
  brecha_cortes <- sort(brecha_cortes)[1:2]
  brecha_corte_min <- brecha_cortes[1]
  brecha_corte_max <- brecha_cortes[2]

  .to_rgb <- function(col, fallback = "#FFFFFF") {
    x <- tryCatch(grDevices::col2rgb(col), error = function(e) NULL)
    if (is.null(x)) x <- grDevices::col2rgb(fallback)
    as.numeric(x[, 1])
  }
  .mix_color <- function(col_bajo, col_alto, t) {
    t <- pmax(0, pmin(1, as.numeric(t)))
    if (!is.finite(t) || is.na(t)) return(col_bajo)
    r0 <- .to_rgb(col_bajo, "#FFFFFF")
    r1 <- .to_rgb(col_alto, "#F4B183")
    rr <- round(r0 + (r1 - r0) * t)
    grDevices::rgb(rr[1], rr[2], rr[3], maxColorValue = 255)
  }
  .calc_brecha <- function(x) {
    v <- suppressWarnings(as.numeric(x))
    v <- v[is.finite(v) & !is.na(v)]
    if (length(v) < 2L) return(NA_real_)
    max(v, na.rm = TRUE) - min(v, na.rm = TRUE)
  }
  .fmt_n_x <- function(x) {
    x <- .dim_round_half_up(x, 0)
    if (!is.finite(x) || is.na(x)) return(NA_character_)
    format(as.integer(x), trim = TRUE, big.mark = ",", scientific = FALSE)
  }

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
  sc_base <- payload$score_heat
  sc <- sc_base

  if (isTRUE(brecha_cols)) {
    sc_bc <- sc_base |>
      dplyr::group_by(.data$axis_label) |>
      dplyr::summarise(score_raw = .calc_brecha(.data$score_raw), .groups = "drop") |>
      dplyr::mutate(
        axis_var = "__brecha_cols__",
        grupo = etiq_brecha_cols,
        tipo = "brecha_cols",
        base = NA_real_,
        score_round = .dim_round_half_up(.data$score_raw, 0)
      ) |>
      dplyr::select("axis_var", "axis_label", "grupo", "tipo", "score_raw", "base", "score_round")
    sc <- dplyr::bind_rows(sc, sc_bc)
  }

  if (isTRUE(brecha_filas)) {
    sc_bf <- sc_base |>
      dplyr::group_by(.data$grupo) |>
      dplyr::summarise(score_raw = .calc_brecha(.data$score_raw), .groups = "drop") |>
      dplyr::mutate(
        axis_var = "__brecha_filas__",
        axis_label = etiq_brecha_filas,
        tipo = "brecha_filas",
        base = NA_real_,
        score_round = .dim_round_half_up(.data$score_raw, 0)
      ) |>
      dplyr::select("axis_var", "axis_label", "grupo", "tipo", "score_raw", "base", "score_round")
    sc <- dplyr::bind_rows(sc, sc_bf)
  }

  if (isTRUE(brecha_cols) && isTRUE(brecha_filas)) {
    sc_corner <- data.frame(
      axis_var = "__brecha_corner__",
      axis_label = etiq_brecha_filas,
      grupo = etiq_brecha_cols,
      tipo = "brecha_corner",
      score_raw = NA_real_,
      base = NA_real_,
      score_round = NA_real_,
      stringsAsFactors = FALSE
    )
    sc <- dplyr::bind_rows(sc, sc_corner)
  }

  sc$grupo <- as.character(sc$grupo)
  sc$axis_label <- as.character(sc$axis_label)
  sc$grupo[sc$grupo == "Total"] <- titulo_total_x
  sc$axis_label[sc$axis_label == "Total cruce"] <- titulo_total_y

  group_order <- payload$group_order %||% unique(as.character(sc_base$grupo))
  group_order <- as.character(group_order)
  group_order[group_order == "Total"] <- titulo_total_x
  if (isTRUE(brecha_cols) && !(etiq_brecha_cols %in% group_order)) {
    group_order <- c(group_order, etiq_brecha_cols)
  }
  axis_order_heat <- payload$axis_order_heat %||% unique(as.character(sc_base$axis_label))
  axis_order_heat <- as.character(axis_order_heat)
  axis_order_heat[axis_order_heat == "Total cruce"] <- titulo_total_y
  if (isTRUE(brecha_filas) && !(etiq_brecha_filas %in% axis_order_heat)) {
    axis_order_heat <- c(axis_order_heat, etiq_brecha_filas)
  }

  if (isTRUE(mostrar_n_cruce_x)) {
    bases_grupo <- sc |>
      dplyr::group_by(.data$grupo) |>
      dplyr::summarise(
        base_plot = {
          b <- suppressWarnings(as.numeric(.data$base))
          b <- b[is.finite(b) & !is.na(b)]
          if (length(b)) b[1] else NA_real_
        },
        .groups = "drop"
      )

    map_n <- stats::setNames(
      vapply(seq_len(nrow(bases_grupo)), function(i) {
        lab <- as.character(bases_grupo$grupo[i])
        n_txt <- .fmt_n_x(as.numeric(bases_grupo$base_plot[i]))
        if (is.na(n_txt) || !nzchar(n_txt)) lab else paste0(lab, " (N=", n_txt, ")")
      }, character(1)),
      as.character(bases_grupo$grupo)
    )

    sc_groups_new <- unname(map_n[sc$grupo])
    keep_old_sc <- is.na(sc_groups_new) | !nzchar(sc_groups_new)
    sc_groups_new[keep_old_sc] <- sc$grupo[keep_old_sc]
    sc$grupo <- sc_groups_new

    group_order_new <- unname(map_n[group_order])
    keep_old_ord <- is.na(group_order_new) | !nzchar(group_order_new)
    group_order_new[keep_old_ord] <- group_order[keep_old_ord]
    group_order <- unique(group_order_new)
  }

  sc$grupo <- factor(sc$grupo, levels = group_order)
  sc$axis_label <- factor(sc$axis_label, levels = rev(axis_order_heat))
  sc$is_brecha <- as.character(sc$tipo %||% "") %in% c("brecha_cols", "brecha_filas", "brecha_corner")
  sc$estado <- dplyr::case_when(
    is.na(sc$score_raw) ~ "Sin dato",
    sc$score_raw < sem$cortes[1] ~ cuts_lab[1],
    sc$score_raw < sem$cortes[2] ~ cuts_lab[2],
    TRUE ~ cuts_lab[3]
  )
  sc$estado <- factor(sc$estado, levels = c(cuts_lab[1], cuts_lab[2], cuts_lab[3], "Sin dato"))

  fill_std <- dplyr::case_when(
    sc$estado == cuts_lab[1] ~ sem$rojo,
    sc$estado == cuts_lab[2] ~ sem$ambar,
    sc$estado == cuts_lab[3] ~ sem$verde,
    TRUE ~ sem$na
  )
  fill_brecha <- rep(col_brecha_alto, nrow(sc))
  if (isTRUE(aplicar_gradiente_brecha)) {
    vals_b <- suppressWarnings(as.numeric(sc$score_raw))
    tt <- if (isTRUE(brecha_corte_max > brecha_corte_min)) {
      pmax(0, pmin(1, (vals_b - brecha_corte_min) / (brecha_corte_max - brecha_corte_min)))
    } else {
      rep(0.5, length(vals_b))
    }
    fill_brecha <- vapply(tt, function(ti) .mix_color(col_brecha_bajo, col_brecha_alto, ti), character(1))
  }
  fill_brecha[!is.finite(suppressWarnings(as.numeric(sc$score_raw))) | is.na(sc$score_raw)] <- sem$na
  sc$fill_hex <- ifelse(sc$is_brecha, fill_brecha, as.character(fill_std))
  sc$label <- ifelse(is.na(sc$score_raw), "", .dim_fmt_int(sc$score_round))

  max_chars <- max(nchar(as.character(axis_order_heat), type = "width"), na.rm = TRUE)
  left_margin <- .dim_clamp(36 + 7 * max_chars, 130, 320)

  p_panel <- ggplot2::ggplot(
    sc,
    ggplot2::aes(x = .data$grupo, y = .data$axis_label, fill = .data$fill_hex)
  ) +
    ggplot2::geom_tile(colour = "#F2F5F9", linewidth = 0.45) +
    ggplot2::geom_text(
      ggplot2::aes(label = .data$label),
      size = size_texto_celdas / 3,
      colour = color_texto_celdas,
      fontface = "bold"
    ) +
    ggplot2::scale_fill_identity() +
    ggplot2::theme_minimal(base_size = 11) +
    ggplot2::theme(
      panel.grid = ggplot2::element_blank(),
      axis.title = ggplot2::element_blank(),
      axis.text.x = ggplot2::element_text(
        size = size_ejes_x,
        colour = color_ejes,
        angle = angle_x,
        hjust = if (abs(angle_x) < 1e-6) 0.5 else 0,
        vjust = if (abs(angle_x) < 1e-6) 0.5 else 1
      ),
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
#' Visualiza indices y bloques de dimensiones como radar (cuando hay 3+ ejes)
#' o barras numéricas comparativas. Requiere que `data` sea la salida encadenada
#' de [reporte_dimensiones()] y [reporte_dimensiones_indices()].
#'
#' @param data Base recodificada e indexada; salida de
#'   `reporte_dimensiones() |> reporte_dimensiones_indices()`.
#' @param modo `"general"` (índices por subindices) o `"indicadores"` (subindices por ítem).
#' @param objetivo Id técnico del catálogo (columna `idx_*` o clave de bloque).
#' @param instrumento Instrumento opcional. Si es `NULL`, se usa `attr(data, "instrumento_reporte")`.
#' @param cruce Variable de comparación opcional (columna en `data`).
#' @param incluir_total Si es `NULL`, usa el default interno.
#' @param inicio_eje_pct Piso visual del eje radial en porcentaje (0-99). Si se
#'   define, se mapea internamente a `limites = c(inicio_eje_pct/100, 1)` y
#'   falla con error si hay valores observados por debajo de ese piso.
#' @param filtros Lista nombrada de filtros por variable.
#' @param iter_var,iter_level Variable y nivel opcionales de iteración.
#' @param titulo,subtitulo,nota_pie Textos del gráfico.
#' @param ... Argumentos adicionales para `graficar_radar()` o `graficar_barras_numericas()`.
#'
#' @return Objeto gráfico (canvas cowplot) o exportación invisible.
#' @family graficador
#' @seealso [reporte_dimensiones()], [reporte_dimensiones_indices()],
#'   [reporte_dimensiones_config()], [graficar_heatmap_dimensiones()]
#' @export
graficar_radar_dimensiones <- function(
    data,
    modo = c("general", "indicadores"),
    objetivo,
    instrumento = NULL,
    cruce = NULL,
    incluir_total = NULL,
    inicio_eje_pct = NULL,
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

  extra_args <- list(...)
  extra_args <- .dim_alias_radar_extra_args(extra_args)

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

  if (identical(payload$visual_mode, "radar")) {
    if (!exists("graficar_radar", mode = "function", inherits = TRUE)) {
      stop("No existe `graficar_radar()`.", call. = FALSE)
    }

    df_plot <- .dim_payload_to_plot_df(payload)
    inicio_eje_pct <- .dim_or(inicio_eje_pct, extra_args$inicio_eje_pct)
    if (!is.null(inicio_eje_pct)) {
      inicio_eje_pct <- suppressWarnings(as.numeric(inicio_eje_pct)[1])
      if (!is.finite(inicio_eje_pct) || inicio_eje_pct < 0 || inicio_eje_pct >= 100) {
        stop("`inicio_eje_pct` debe ser NULL o un número en [0, 100).", call. = FALSE)
      }
      vals <- suppressWarnings(as.numeric(df_plot$valor))
      vals <- vals[is.finite(vals) & !is.na(vals)]
      if (length(vals)) {
        min_obs <- suppressWarnings(min(vals, na.rm = TRUE))
        if (is.finite(min_obs) && min_obs < inicio_eje_pct) {
          stop(
            "`inicio_eje_pct`=", format(inicio_eje_pct, trim = TRUE),
            " no es válido: el mínimo observado es ",
            format(round(min_obs, 1), trim = TRUE),
            ". Ajuste el piso o revise los datos.",
            call. = FALSE
          )
        }
      }
      if (is.null(extra_args$limites)) {
        extra_args$limites <- c(inicio_eje_pct / 100, 1)
      }
    }
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
      color_grilla = "#D9E1EA",
      color_radios = "#E4EAF1",
      cortes_grilla = 4,
      wrap_ejes = 22,
      eje_label_mult = 1.03,
      leyenda_posicion = "abajo",
      legend_n_por_fila = 4,
      legend_key_cm = 0.45,
      legend_espaciado = 12,
      canvas_h_header_in = 0.58,
      canvas_h_legend_in = 0.20,
      canvas_h_caption_in = 0.08
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
#' Visualiza indices y bloques de dimensiones como radar o barras, con una
#' tabla adjunta de valores numéricos. Requiere que `data` sea la salida
#' encadenada de [reporte_dimensiones()] y [reporte_dimensiones_indices()].
#'
#' @param data Base recodificada e indexada; salida de
#'   `reporte_dimensiones() |> reporte_dimensiones_indices()`.
#' @param modo `"general"` (índices por subindices) o `"indicadores"` (subindices por ítem).
#' @param objetivo Id técnico del catálogo (columna `idx_*` o clave de bloque).
#' @param instrumento Instrumento opcional. Si es `NULL`, se usa `attr(data, "instrumento_reporte")`.
#' @param cruce Variable de comparación opcional (columna en `data`).
#' @param incluir_total Si es `NULL`, usa el default interno.
#' @param filtros Lista nombrada de filtros por variable.
#' @param iter_var,iter_level Variable y nivel opcionales de iteración.
#' @param titulo,subtitulo,nota_pie Textos del gráfico.
#' @param titulo_tabla Título de la primera columna de la tabla adjunta.
#' @param ... Argumentos adicionales del radar, barras y tabla.
#'
#' @return Objeto gráfico (canvas cowplot) o exportación invisible.
#' @family graficador
#' @seealso [reporte_dimensiones()], [reporte_dimensiones_indices()],
#'   [reporte_dimensiones_config()], [graficar_heatmap_dimensiones()]
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

  extra_args <- list(...)
  extra_args <- .dim_alias_radar_extra_args(extra_args)

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

# =============================================================================
# FODA dimensiones
# =============================================================================

#' Matriz FODA automática de dimensiones
#'
#' Clasifica indicadores o subíndices en una matriz 2x2 FODA (Fortalezas,
#' Oportunidades, Debilidades, Amenazas) basándose en dos ejes derivados de
#' los datos: el puntaje promedio (0-100) y la variabilidad (desviación
#' estándar) entre respondentes.
#'
#' La clasificación es automática:
#' \itemize{
#'   \item \strong{Fortaleza}: puntaje alto + SD baja (consolidado y fuerte).
#'   \item \strong{Oportunidad}: puntaje alto + SD alta (fuerte pero desigual).
#'   \item \strong{Debilidad}: puntaje bajo + SD baja (rezago estructural).
#'   \item \strong{Amenaza}: puntaje bajo + SD alta (rezago con alta dispersión).
#' }
#'
#' @param data Base recodificada e indexada; salida de
#'   \code{reporte_dimensiones() |> reporte_dimensiones_indices()}.
#' @param nivel \code{"subindices"} para clasificar columnas \verb{sub_*},
#'   o \code{"indicadores"} para clasificar columnas \verb{r100_*} de un bloque.
#' @param objetivo Para \code{nivel = "indicadores"}: clave del bloque en el
#'   catálogo de indicadores. Para \code{nivel = "subindices"} es opcional (si
#'   se provee, usa solo los axis_vars de ese índice).
#' @param modo_foda Modo de visualización: \code{"matriz"} (cuadrantes con
#'   tarjetas) o \code{"dispersion"} (scatter con coordenadas reales).
#' @param instrumento Instrumento opcional. Si es \code{NULL}, se usa
#'   \code{attr(data, "instrumento_reporte")}.
#' @param cruce Variable opcional de cruce. Solo se aplica cuando
#'   \code{modo_foda = "dispersion"}.
#' @param incluir_total Si \code{TRUE}, agrega el grupo total en modo
#'   dispersión combinada con cruce.
#' @param solo_indice_general_cruce Reservado para compatibilidad. En
#'   \code{nivel = "subindices"} con \code{modo_foda = "dispersion"} no se usa
#'   para evitar reemplazar el desglose por cruce.
#' @param filtros Lista nombrada de filtros por variable.
#' @param usar_pesos Si \code{TRUE} y existe columna de peso, calcula media y
#'   SD ponderadas para clasificar cuadrantes.
#' @param corte_score Corte de puntaje para clasificar alto/bajo. Si es
#'   \code{NULL}, usa \code{semaforo$cortes[2]} (default 75), salvo en
#'   \code{modo_foda = "dispersion"} donde es obligatorio.
#' @param corte_sd Corte de SD para clasificar variabilidad alta/baja. Si es
#'   \code{NULL}, usa la mediana de las SDs calculadas.
#' @param colores_foda Vector nombrado de 4 colores de fondo para los
#'   cuadrantes (\code{fortaleza}, \code{oportunidad}, \code{debilidad},
#'   \code{amenaza}).
#' @param titulo,subtitulo,nota_pie Textos del gráfico.
#' @param color_titulo,size_titulo Estilo del título.
#' @param color_subtitulo,size_subtitulo Estilo del subtítulo.
#' @param color_nota_pie,size_nota_pie Estilo del pie de página.
#' @param color_cuadrante_titulo,size_cuadrante_titulo Estilo de los títulos
#'   de cuadrante (FORTALEZAS, OPORTUNIDADES, etc.).
#' @param titulos_areas_foda Vector nombrado con los títulos por cuadrante
#'   (\code{fortaleza}, \code{oportunidad}, \code{debilidad}, \code{amenaza}).
#'   Si se omite, usa los títulos tradicionales por defecto.
#' @param mostrar_subtitulo_area Si \code{TRUE}, muestra el subtítulo
#'   descriptivo dentro de cada área del FODA.
#' @param sd_tecnico Si \code{TRUE}, muestra el corte técnico de SD y eje X
#'   numérico. Si \code{FALSE}, oculta el corte SD y usa extremos narrativos
#'   en X (\code{"Menor dispersión"} / \code{"Mayor dispersión"}).
#' @param color_indice_total Color de la tarjeta del \code{"Indice"} total
#'   global en modo dispersión con cruce.
#' @param disposicion_recuadro Disposición del texto en tarjetas de dispersión
#'   con cruce: \code{"dos_lineas"} (\code{subíndice} arriba y \code{cruce}
#'   abajo), \code{"una_linea"} (ambos en una línea) o \code{"sin_cruce"}
#'   (solo subíndice). El \code{Indice} total global siempre se muestra en una
#'   sola línea con su nombre agregado.
#' @param etiqueta_cruce_en_dos_lineas Alias legado opcional para
#'   compatibilidad. Si se define, \code{TRUE} equivale a
#'   \code{disposicion_recuadro = "dos_lineas"} y \code{FALSE} a
#'   \code{"una_linea"}.
#' @param color_items,size_items Estilo de los ítems dentro de cada cuadrante.
#' @param ancho_tarjeta_base_rel Ancho base relativo de tarjetas.
#' @param factor_ancho_matriz Multiplicador del ancho de tarjeta para modo
#'   \code{"matriz"}.
#' @param factor_ancho_dispersion Multiplicador del ancho de tarjeta para modo
#'   \code{"dispersion"}.
#' @param ancho_recuadro_rel Ajuste relativo opcional del ancho de recuadro
#'   respecto al cálculo automático (\code{1 = automático}).
#' @param ancho_recuadro_auto Si \code{TRUE}, ajusta automáticamente el ancho
#'   de cada recuadro según su contenido para reducir espacio sobrante entre
#'   texto y chip.
#' @param ancho_chip_rel Ancho relativo del chip de puntaje dentro del recuadro.
#'   Se acota internamente para preservar legibilidad.
#' @param sufijo_puntaje Sufijo del puntaje mostrado en el chip (por ejemplo
#'   \code{" pts"}).
#' @param cortes_chip Cortes del semáforo para los chips de puntaje
#'   (vector numérico de longitud 2). Si es \code{NULL}, usa
#'   \code{c(semaforo$cortes[1], corte_score)} para mantener consistencia con
#'   el corte principal de puntaje.
#' @param tamano_texto_tarjeta Tamaño base del texto del subíndice dentro de la
#'   tarjeta.
#' @param tamano_letra_recuadro Alias opcional en español para
#'   \code{tamano_texto_tarjeta}. Si se define, tiene prioridad.
#' @param tamano_texto_chip Tamaño base del texto del chip de puntaje.
#' @param tarjetas_color_solido Si \code{TRUE}, aplica color sólido a las
#'   tarjetas según cuadrante FODA y texto blanco.
#' @param jitter_x_rel Intensidad relativa del jitter horizontal en dispersión.
#' @param jitter_y_rel Intensidad relativa del jitter vertical en dispersión.
#' @param iter_separacion Número de iteraciones de separación para reducir
#'   solapes de tarjetas en dispersión.
#' @param factor_reduccion_tarjeta_dispersion Factor de reducción de tamaño de
#'   tarjetas en modo dispersión.
#' @param color_fondo Color de fondo del gráfico.
#' @param mostrar_leyenda Si \code{TRUE}, muestra leyenda explicativa.
#' @param usar_canvas Si \code{TRUE}, compone con \code{cowplot}
#'   (título/panel/leyenda/pie).
#' @param canvas_h_title,canvas_h_legend,canvas_h_caption,canvas_pad_top
#'   Proporciones de altura del canvas.
#' @param debug_ph_bordes,debug_ph_col,debug_ph_lwd Bordes de depuración.
#' @param exportar Tipo de exportación: \code{"rplot"}, \code{"png"},
#'   \code{"ppt"} o \code{"word"}.
#' @param path_salida Ruta de salida para \code{exportar != "rplot"}.
#' @param ancho,alto,dpi Tamaño y resolución de exportación.
#' @param ppt_append,ppt_layout,ppt_master Parámetros de exportación PPT.
#'
#' @return Objeto gráfico (canvas cowplot) o exportación invisible.
#' @family indicador
#' @family graficador
#' @seealso [reporte_dimensiones()], [reporte_dimensiones_indices()],
#'   [reporte_dimensiones_config()], [graficar_heatmap_dimensiones()],
#'   [graficar_radar_dimensiones()]
#' @export
graficar_foda_dimensiones <- function(
    data,
    nivel = c("subindices", "indicadores"),
    objetivo = NULL,
    modo_foda = c("matriz", "dispersion"),
    instrumento = NULL,
    cruce = NULL,
    incluir_total = TRUE,
    solo_indice_general_cruce = FALSE,
    filtros = list(),
    usar_pesos = TRUE,
    corte_score = NULL,
    corte_sd = NULL,
    colores_foda = c(
      fortaleza = "#E8F5E9", oportunidad = "#E3F2FD",
      debilidad = "#FFEBEE", amenaza = "#FFF3E0"
    ),
    titulo = NULL,
    subtitulo = NULL,
    nota_pie = NULL,
    color_titulo = "#004B8D",
    size_titulo = 12,
    color_subtitulo = "#004B8D",
    size_subtitulo = 9,
    color_nota_pie = "#004B8D",
    size_nota_pie = 8,
    color_cuadrante_titulo = "#20324d",
    size_cuadrante_titulo = 11,
    color_items = "#122842",
    size_items = 9,
    titulos_areas_foda = c(
      fortaleza = "FORTALEZAS",
      oportunidad = "OPORTUNIDADES",
      debilidad = "DEBILIDADES",
      amenaza = "AMENAZAS"
    ),
    mostrar_subtitulo_area = TRUE,
    sd_tecnico = TRUE,
    color_indice_total = "#FF6A00",
    disposicion_recuadro = c("dos_lineas", "una_linea", "sin_cruce"),
    etiqueta_cruce_en_dos_lineas = NULL,
    ancho_tarjeta_base_rel = 0.72,
    factor_ancho_matriz = 1.00,
    factor_ancho_dispersion = 0.72,
    ancho_recuadro_rel = NULL,
    ancho_recuadro_auto = FALSE,
    ancho_chip_rel = 0.18,
    sufijo_puntaje = " pts",
    cortes_chip = NULL,
    tamano_texto_tarjeta = NULL,
    tamano_letra_recuadro = NULL,
    tamano_texto_chip = NULL,
    tarjetas_color_solido = TRUE,
    jitter_x_rel = 0.06,
    jitter_y_rel = 0.03,
    iter_separacion = 12L,
    factor_reduccion_tarjeta_dispersion = 0.85,
    chip_width_rel = NULL,
    score_suffix = NULL,
    color_fondo = NA,
    mostrar_leyenda = TRUE,
    usar_canvas = TRUE,
    canvas_h_title = 0,
    canvas_h_legend = 0.09,
    canvas_h_caption = 0.06,
    canvas_pad_top = 0.01,
    debug_ph_bordes = FALSE,
    debug_ph_col = "#FF00FF",
    debug_ph_lwd = 0.6,
    exportar = c("rplot", "png", "ppt", "word"),
    path_salida = NULL,
    ancho = 10,
    alto = 7,
    dpi = 300,
    ppt_append = TRUE,
    ppt_layout = "Blank",
    ppt_master = "Office Theme"
) {
  if (!requireNamespace("ggplot2", quietly = TRUE)) stop("Requiere ggplot2.", call. = FALSE)
  if (!requireNamespace("cowplot", quietly = TRUE)) stop("Requiere cowplot.", call. = FALSE)

  nivel    <- match.arg(nivel)
  modo_foda <- match.arg(modo_foda)
  exportar <- match.arg(exportar)
  usar_pesos <- isTRUE(usar_pesos)
  incluir_total <- isTRUE(incluir_total)
  mostrar_subtitulo_area <- isTRUE(mostrar_subtitulo_area)
  sd_tecnico <- isTRUE(sd_tecnico)
  if (!is.null(etiqueta_cruce_en_dos_lineas)) {
    if (!is.logical(etiqueta_cruce_en_dos_lineas) || length(etiqueta_cruce_en_dos_lineas) != 1L || is.na(etiqueta_cruce_en_dos_lineas)) {
      stop("`etiqueta_cruce_en_dos_lineas` debe ser NULL o logical(1).", call. = FALSE)
    }
    disposicion_recuadro <- if (isTRUE(etiqueta_cruce_en_dos_lineas)) "dos_lineas" else "una_linea"
  }
  disposicion_recuadro <- as.character(disposicion_recuadro %||% "dos_lineas")[1]
  if (!nzchar(disposicion_recuadro) || is.na(disposicion_recuadro)) disposicion_recuadro <- "dos_lineas"
  disposicion_recuadro <- match.arg(disposicion_recuadro, c("dos_lineas", "una_linea", "sin_cruce"))
  cruce <- as.character(cruce %||% "")[1]
  color_indice_total <- as.character(color_indice_total %||% "#FF6A00")[1]
  if (!nzchar(trimws(color_indice_total)) || is.na(color_indice_total)) color_indice_total <- "#FF6A00"
  if (inherits(try(grDevices::col2rgb(color_indice_total), silent = TRUE), "try-error")) color_indice_total <- "#FF6A00"

  titulos_default <- c(
    fortaleza = "FORTALEZAS",
    oportunidad = "OPORTUNIDADES",
    debilidad = "DEBILIDADES",
    amenaza = "AMENAZAS"
  )
  titulos_in <- as.character(titulos_areas_foda)
  titulos_out <- titulos_default
  if (length(titulos_in)) {
    nms_t <- names(titulos_areas_foda %||% character(0))
    if (is.null(nms_t)) nms_t <- character(0)
    if (!length(nms_t) || !any(nzchar(trimws(nms_t)))) {
      n_take <- min(length(titulos_default), length(titulos_in))
      if (n_take > 0L) titulos_out[seq_len(n_take)] <- titulos_in[seq_len(n_take)]
    } else {
      nms_t <- tolower(trimws(as.character(nms_t)))
      for (k in names(titulos_default)) {
        hit <- which(nms_t == k)
        if (length(hit)) titulos_out[k] <- titulos_in[hit[1]]
      }
    }
  }
  titulos_out <- vapply(
    names(titulos_out),
    function(k) {
      x <- as.character(titulos_out[[k]] %||% titulos_default[[k]])[1]
      x <- trimws(x)
      if (!nzchar(x) || is.na(x)) titulos_default[[k]] else x
    },
    character(1)
  )
  names(titulos_out) <- names(titulos_default)

  ancho_tarjeta_base_rel <- suppressWarnings(as.numeric(ancho_tarjeta_base_rel)[1])
  if (!is.finite(ancho_tarjeta_base_rel) || is.na(ancho_tarjeta_base_rel)) ancho_tarjeta_base_rel <- 0.72
  ancho_tarjeta_base_rel <- .dim_clamp(ancho_tarjeta_base_rel, 0.50, 0.90)
  factor_ancho_matriz <- suppressWarnings(as.numeric(factor_ancho_matriz)[1])
  if (!is.finite(factor_ancho_matriz) || is.na(factor_ancho_matriz)) factor_ancho_matriz <- 1.00
  factor_ancho_matriz <- .dim_clamp(factor_ancho_matriz, 0.60, 1.40)
  factor_ancho_dispersion <- suppressWarnings(as.numeric(factor_ancho_dispersion)[1])
  if (!is.finite(factor_ancho_dispersion) || is.na(factor_ancho_dispersion)) factor_ancho_dispersion <- 0.72
  factor_ancho_dispersion <- .dim_clamp(factor_ancho_dispersion, 0.40, 1.20)
  ancho_tarjeta_matriz <- .dim_clamp(ancho_tarjeta_base_rel * factor_ancho_matriz, 0.48, 0.88)
  ancho_tarjeta_disp <- .dim_clamp(ancho_tarjeta_base_rel * factor_ancho_dispersion, 0.36, 0.78)
  ancho_recuadro_auto <- isTRUE(ancho_recuadro_auto)
  ancho_recuadro_rel <- suppressWarnings(as.numeric(ancho_recuadro_rel)[1])
  if (is.finite(ancho_recuadro_rel) && !is.na(ancho_recuadro_rel) && ancho_recuadro_rel > 0) {
    ancho_recuadro_rel <- .dim_clamp(ancho_recuadro_rel, 0.55, 1.45)
    ancho_tarjeta_matriz <- .dim_clamp(ancho_tarjeta_matriz * ancho_recuadro_rel, 0.36, 0.88)
    ancho_tarjeta_disp <- .dim_clamp(ancho_tarjeta_disp * ancho_recuadro_rel, 0.26, 0.78)
  }

  if (!is.null(chip_width_rel)) ancho_chip_rel <- chip_width_rel
  if (!is.null(score_suffix)) sufijo_puntaje <- score_suffix
  ancho_chip_rel <- suppressWarnings(as.numeric(ancho_chip_rel)[1])
  if (!is.finite(ancho_chip_rel) || is.na(ancho_chip_rel)) ancho_chip_rel <- 0.18
  ancho_chip_rel <- .dim_clamp(ancho_chip_rel, 0.10, 0.36)
  sufijo_puntaje <- as.character(sufijo_puntaje %||% " pts")[1]
  if (is.na(sufijo_puntaje)) sufijo_puntaje <- " pts"
  if (!is.null(tamano_letra_recuadro)) tamano_texto_tarjeta <- tamano_letra_recuadro
  tamano_texto_tarjeta <- suppressWarnings(as.numeric(tamano_texto_tarjeta)[1])
  if (!is.finite(tamano_texto_tarjeta) || is.na(tamano_texto_tarjeta) || tamano_texto_tarjeta <= 0) {
    tamano_texto_tarjeta <- size_items
  }
  tamano_texto_chip <- suppressWarnings(as.numeric(tamano_texto_chip)[1])
  if (!is.finite(tamano_texto_chip) || is.na(tamano_texto_chip) || tamano_texto_chip <= 0) {
    tamano_texto_chip <- max(8, size_items + 1.0)
  }
  tarjetas_color_solido <- isTRUE(tarjetas_color_solido)
  jitter_x_rel <- suppressWarnings(as.numeric(jitter_x_rel)[1])
  if (!is.finite(jitter_x_rel) || is.na(jitter_x_rel) || jitter_x_rel < 0) jitter_x_rel <- 0.06
  jitter_x_rel <- .dim_clamp(jitter_x_rel, 0, 0.20)
  jitter_y_rel <- suppressWarnings(as.numeric(jitter_y_rel)[1])
  if (!is.finite(jitter_y_rel) || is.na(jitter_y_rel) || jitter_y_rel < 0) jitter_y_rel <- 0.045
  jitter_y_rel <- .dim_clamp(jitter_y_rel, 0, 0.20)
  iter_separacion <- suppressWarnings(as.integer(iter_separacion)[1])
  if (!is.finite(iter_separacion) || is.na(iter_separacion) || iter_separacion < 0L) iter_separacion <- 12L
  iter_separacion <- max(0L, min(30L, iter_separacion))
  factor_reduccion_tarjeta_dispersion <- suppressWarnings(as.numeric(factor_reduccion_tarjeta_dispersion)[1])
  if (!is.finite(factor_reduccion_tarjeta_dispersion) || is.na(factor_reduccion_tarjeta_dispersion)) {
    factor_reduccion_tarjeta_dispersion <- 0.85
  }
  factor_reduccion_tarjeta_dispersion <- .dim_clamp(factor_reduccion_tarjeta_dispersion, 0.55, 1.00)

  ctx <- .dim_build_context(data, instrumento = instrumento)
  obj <- NULL
  sem <- ctx$semaforo

  # --- Resolver variables y etiquetas segun nivel ---
  if (identical(nivel, "subindices")) {
    if (!is.null(objetivo) && nzchar(as.character(objetivo)[1])) {
      obj <- ctx$catalog_general[[as.character(objetivo)[1]]]
      if (is.null(obj)) {
        stop(
          "graficar_foda_dimensiones(): `objetivo` '", objetivo,
          "' no encontrado en catalog_general.",
          call. = FALSE
        )
      }
      vars   <- obj$axis_vars
      labels <- obj$axis_labels
    } else {
      vars   <- character(0)
      labels <- character(0)
      for (entry in ctx$catalog_general) {
        new_vars   <- setdiff(entry$axis_vars, vars)
        new_labels <- entry$axis_labels[entry$axis_vars %in% new_vars]
        vars   <- c(vars, new_vars)
        labels <- c(labels, new_labels)
      }
      if (!length(vars)) {
        sub_cols <- grep("^sub_", names(data), value = TRUE)
        vars   <- sub_cols
        labels <- vapply(sub_cols, ctx$label_sub, character(1))
      }
    }
  } else {
    if (is.null(objetivo) || !nzchar(as.character(objetivo)[1])) {
      stop(
        "graficar_foda_dimensiones(): `objetivo` es requerido para nivel='indicadores'.",
        call. = FALSE
      )
    }
    obj <- ctx$catalog_indicadores[[as.character(objetivo)[1]]]
    if (is.null(obj)) {
      stop(
        "graficar_foda_dimensiones(): `objetivo` '", objetivo,
        "' no encontrado en catalog_indicadores.",
        call. = FALSE
      )
    }
    vars   <- obj$axis_vars
    labels <- obj$axis_labels
  }

  vars <- as.character(vars)
  labels <- as.character(labels)
  if (length(labels) < length(vars)) {
    labels <- c(labels, vars[seq.int(length(labels) + 1L, length(vars))])
  }
  labels <- labels[seq_along(vars)]
  keep_vars <- vars %in% names(data)
  vars <- vars[keep_vars]
  labels <- labels[keep_vars]

  .resolve_indice_total <- function() {
    out_var <- NA_character_
    out_lbl <- NA_character_

    .norm_txt <- function(x) {
      x <- tolower(trimws(as.character(x %||% "")))
      x <- suppressWarnings(iconv(x, to = "ASCII//TRANSLIT"))
      x[is.na(x)] <- ""
      x <- gsub("[^a-z0-9]+", " ", x)
      trimws(gsub("\\s+", " ", x))
    }

    if (is.na(out_var) || !nzchar(out_var)) {
      best_score <- -Inf
      best_var <- NA_character_
      best_lbl <- NA_character_
      for (nm in names(ctx$catalog_general %||% list())) {
        cc <- ctx$catalog_general[[nm]]
        if (!is.list(cc)) next
        id_cc <- as.character(cc$id %||% nm %||% NA_character_)[1]
        if (is.na(id_cc) || !nzchar(id_cc) || !(id_cc %in% names(data))) next
        key_cc <- as.character(cc$key %||% nm %||% "")[1]
        lbl_cc <- as.character(cc$label %||% id_cc)[1]
        axis_cc <- unique(as.character(cc$axis_vars %||% character(0)))
        cov_cc <- if (length(vars)) mean(vars %in% axis_cc) else 0
        score_cc <- length(axis_cc) + (10 * cov_cc)
        txt_cc <- .norm_txt(paste(id_cc, key_cc, lbl_cc))
        if (grepl("\\bindice general\\b|\\bindice_general\\b", txt_cc)) {
          score_cc <- score_cc + 1000
        } else if (grepl("\\bgeneral\\b", txt_cc)) {
          score_cc <- score_cc + 200
        }
        if (is.finite(score_cc) && score_cc > best_score) {
          best_score <- score_cc
          best_var <- id_cc
          best_lbl <- lbl_cc
        }
      }
      if (is.finite(best_score) && nzchar(best_var)) {
        out_var <- best_var
        out_lbl <- best_lbl
      }
    }

    if ((is.na(out_var) || !nzchar(out_var)) && ("idx_indice_general" %in% names(data))) {
      out_var <- "idx_indice_general"
      out_lbl <- "Índice General"
    }
    if ((is.na(out_var) || !nzchar(out_var)) && !is.null(obj) && identical(nivel, "subindices")) {
      id_obj <- as.character(obj$id %||% NA_character_)[1]
      if (!is.na(id_obj) && nzchar(id_obj) && id_obj %in% names(data)) {
        out_var <- id_obj
        out_lbl <- as.character(obj$label %||% id_obj)[1]
      }
    }
    if ((is.na(out_var) || !nzchar(out_var))) {
      idx_guess <- grep("^idx_", names(data), value = TRUE)
      if (length(idx_guess)) {
        out_var <- idx_guess[1]
        out_lbl <- .dim_pretty_label(out_var)
      }
    }
    list(var = out_var, label = out_lbl)
  }

  if (!length(vars)) {
    return(.dim_export_canvas(
      .dim_blank_canvas(
        mensaje = "Sin variables disponibles para FODA",
        debug_ph_bordes = debug_ph_bordes,
        debug_ph_col = debug_ph_col,
        debug_ph_lwd = debug_ph_lwd
      ),
      exportar = exportar,
      path_salida = path_salida,
      ancho = ancho, alto = alto, dpi = dpi,
      ppt_append = ppt_append,
      ppt_layout = ppt_layout,
      ppt_master = ppt_master
    ))
  }

  # --- Aplicar filtros ---
  cruce <- trimws(as.character(cruce %||% "")[1])
  if (identical(cruce, "NA")) cruce <- ""
  if (identical(modo_foda, "matriz") && nzchar(cruce)) {
    stop("`cruce` solo está disponible cuando `modo_foda = 'dispersion'`.", call. = FALSE)
  }

  df <- .dim_apply_filters(ctx$data, filters = filtros)
  if (!nrow(df)) {
    return(.dim_export_canvas(
      .dim_blank_canvas(
        mensaje = "Sin datos luego de aplicar filtros para FODA",
        debug_ph_bordes = debug_ph_bordes,
        debug_ph_col = debug_ph_col,
        debug_ph_lwd = debug_ph_lwd
      ),
      exportar = exportar,
      path_salida = path_salida,
      ancho = ancho, alto = alto, dpi = dpi,
      ppt_append = ppt_append,
      ppt_layout = ppt_layout,
      ppt_master = ppt_master
    ))
  }

  total_group_label <- "Indice"

  # --- Calcular stats (base o por cruce) ---
  if (identical(modo_foda, "dispersion") && nzchar(cruce)) {
    if (!cruce %in% names(df)) {
      stop("`cruce` no existe en `data`: ", cruce, call. = FALSE)
    }

    w_cruce <- .dim_safe_weights(df, weight_col = ctx$weight_col)
    lev <- .dim_categorias_var(
      df = df,
      var = cruce,
      w = w_cruce,
      data_ref = ctx$data,
      instrumento = ctx$instrumento,
      max_levels = max(1000L, as.integer(nrow(df) + 1L))
    )
    lev_df <- lev$rows %||% data.frame()
    if (nrow(lev_df)) {
      is_total_like <- function(x) {
        x <- tolower(trimws(as.character(x %||% "")))
        x <- suppressWarnings(iconv(x, to = "ASCII//TRANSLIT"))
        x[is.na(x)] <- ""
        x <- gsub("[^a-z0-9]+", " ", x)
        x <- trimws(gsub("\\s+", " ", x))
        x %in% c(
          "total", "totales",
          "indice total", "indice totales",
          "indice general total", "indice general totales"
        )
      }
      keep_lev <- !(is_total_like(lev_df$value) | is_total_like(lev_df$label))
      lev_df <- lev_df[keep_lev, , drop = FALSE]
    }
    if (!nrow(lev_df)) {
      return(.dim_export_canvas(
        .dim_blank_canvas(
          mensaje = "Sin niveles válidos para el cruce en modo dispersión",
          debug_ph_bordes = debug_ph_bordes,
          debug_ph_col = debug_ph_col,
          debug_ph_lwd = debug_ph_lwd
        ),
        exportar = exportar,
        path_salida = path_salida,
        ancho = ancho, alto = alto, dpi = dpi,
        ppt_append = ppt_append,
        ppt_layout = ppt_layout,
        ppt_master = ppt_master
      ))
    }

    x_cruce <- trimws(as.character(df[[cruce]]))
    add_total_idx <- .resolve_indice_total()
    has_total_idx <- isTRUE(incluir_total) &&
      !is.na(add_total_idx$var) &&
      nzchar(as.character(add_total_idx$var))

    stats_list <- vector("list", nrow(lev_df) + if (isTRUE(has_total_idx)) 1L else 0L)
    k <- 1L
    for (i in seq_len(nrow(lev_df))) {
      key_i <- trimws(as.character(lev_df$value[i] %||% ""))
      if (!nzchar(key_i)) next
      mask_i <- !is.na(x_cruce) & nzchar(x_cruce) & (x_cruce == key_i)
      if (!any(mask_i)) next

      st_i <- .foda_compute_stats(
        data = df[mask_i, , drop = FALSE],
        vars = vars,
        labels = labels,
        usar_pesos = usar_pesos,
        weight_col = ctx$weight_col
      )
      st_i$grupo_key <- key_i
      st_i$grupo <- as.character(lev_df$label[i] %||% key_i)
      st_i$is_total_global <- FALSE
      stats_list[[k]] <- st_i
      k <- k + 1L
    }

    if (isTRUE(has_total_idx)) {
      st_total <- .foda_compute_stats(
        data = df,
        vars = as.character(add_total_idx$var)[1],
        labels = as.character(add_total_idx$label %||% add_total_idx$var)[1],
        usar_pesos = usar_pesos,
        weight_col = ctx$weight_col
      )
      st_total$grupo_key <- "__total__"
      st_total$grupo <- total_group_label
      st_total$is_total_global <- TRUE
      stats_list[[k]] <- st_total
    }

    stats_list <- stats_list[vapply(stats_list, function(x) !is.null(x), logical(1))]
    stats_df <- if (length(stats_list)) {
      do.call(rbind, stats_list)
    } else {
      data.frame()
    }
  } else {
    stats_df <- .foda_compute_stats(
      data = df,
      vars = vars,
      labels = labels,
      usar_pesos = usar_pesos,
      weight_col = ctx$weight_col
    )
    stats_df$grupo_key <- "__total__"
    stats_df$grupo <- total_group_label
    stats_df$is_total_global <- FALSE
  }

  if (!nrow(stats_df)) {
    return(.dim_export_canvas(
      .dim_blank_canvas(
        mensaje = "Sin datos suficientes para FODA",
        debug_ph_bordes = debug_ph_bordes,
        debug_ph_col = debug_ph_col,
        debug_ph_lwd = debug_ph_lwd
      ),
      exportar = exportar,
      path_salida = path_salida,
      ancho = ancho, alto = alto, dpi = dpi,
      ppt_append = ppt_append,
      ppt_layout = ppt_layout,
      ppt_master = ppt_master
    ))
  }

  stats_df$grupo <- as.character(stats_df$grupo)
  stats_df$grupo_key <- as.character(stats_df$grupo_key)
  if (!("is_total_global" %in% names(stats_df))) stats_df$is_total_global <- FALSE
  stats_df$is_total_global <- !is.na(stats_df$is_total_global) & as.logical(stats_df$is_total_global)
  stats_df <- stats_df[!is.na(stats_df$score_mean), , drop = FALSE]
  if (!nrow(stats_df)) {
    return(.dim_export_canvas(
      .dim_blank_canvas(
        mensaje = "Sin datos clasificables para FODA",
        debug_ph_bordes = debug_ph_bordes,
        debug_ph_col = debug_ph_col,
        debug_ph_lwd = debug_ph_lwd
      ),
      exportar = exportar,
      path_salida = path_salida,
      ancho = ancho, alto = alto, dpi = dpi,
      ppt_append = ppt_append,
      ppt_layout = ppt_layout,
      ppt_master = ppt_master
    ))
  }

  corte_score_raw <- suppressWarnings(as.numeric(corte_score)[1])
  if (identical(modo_foda, "dispersion") &&
      (is.na(corte_score_raw) || !is.finite(corte_score_raw))) {
    stop("En `modo_foda='dispersion'` el argumento `corte_score` es obligatorio.", call. = FALSE)
  }
  corte_score_val <- if (is.na(corte_score_raw) || !is.finite(corte_score_raw)) sem$cortes[2] else corte_score_raw

  sd_vals <- stats_df$score_sd[!is.na(stats_df$score_sd) & is.finite(stats_df$score_sd)]
  corte_sd_val <- suppressWarnings(as.numeric(corte_sd)[1])
  if (is.na(corte_sd_val) || !is.finite(corte_sd_val)) {
    corte_sd_val <- if (length(sd_vals)) stats::median(sd_vals) else 25
  }

  stats_df <- .foda_classify(stats_df, corte_score_val, corte_sd_val)
  stats_df <- stats_df[!is.na(stats_df$cuadrante), , drop = FALSE]
  if (!nrow(stats_df)) {
    return(.dim_export_canvas(
      .dim_blank_canvas(
        mensaje = "Sin datos clasificables para FODA",
        debug_ph_bordes = debug_ph_bordes,
        debug_ph_col = debug_ph_col,
        debug_ph_lwd = debug_ph_lwd
      ),
      exportar = exportar,
      path_salida = path_salida,
      ancho = ancho, alto = alto, dpi = dpi,
      ppt_append = ppt_append,
      ppt_layout = ppt_layout,
      ppt_master = ppt_master
    ))
  }

  # --- Resolver colores FODA ---
  colores_foda <- as.character(colores_foda)
  nms_cf <- names(colores_foda)
  if (is.null(nms_cf)) nms_cf <- character(0)
  col_f <- if ("fortaleza"   %in% nms_cf) colores_foda[["fortaleza"]]   else "#E8F5E9"
  col_o <- if ("oportunidad" %in% nms_cf) colores_foda[["oportunidad"]] else "#E3F2FD"
  col_d <- if ("debilidad"   %in% nms_cf) colores_foda[["debilidad"]]   else "#FFEBEE"
  col_a <- if ("amenaza"     %in% nms_cf) colores_foda[["amenaza"]]     else "#FFF3E0"

  .is_light_col <- function(col, threshold = 0.62) {
    rgb <- tryCatch(grDevices::col2rgb(col) / 255, error = function(e) NULL)
    if (is.null(rgb) || !ncol(rgb)) return(FALSE)
    lum <- 0.2126 * rgb[1, 1] + 0.7152 * rgb[2, 1] + 0.0722 * rgb[3, 1]
    is.finite(lum) && !is.na(lum) && lum >= threshold
  }

  # --- Color semáforo por score (independiente del corte de cuadrantes) ---
  chip_cortes <- suppressWarnings(as.numeric(cortes_chip))
  chip_cortes <- chip_cortes[is.finite(chip_cortes)]
  if (length(chip_cortes) >= 2L) {
    chip_cortes <- sort(unique(chip_cortes))[1:2]
  } else {
    chip_lo <- suppressWarnings(as.numeric(sem$cortes[1])[1])
    chip_hi <- corte_score_val
    if (!is.finite(chip_lo) || is.na(chip_lo)) chip_lo <- chip_hi - 15
    if (!is.finite(chip_hi) || is.na(chip_hi)) chip_hi <- suppressWarnings(as.numeric(sem$cortes[2])[1])
    if (!is.finite(chip_hi) || is.na(chip_hi)) chip_hi <- chip_lo + 15
    if (chip_lo >= chip_hi) chip_lo <- chip_hi - 10
    chip_cortes <- c(chip_lo, chip_hi)
  }
  chip_cortes <- pmax(-Inf, pmin(Inf, chip_cortes))
  if (length(chip_cortes) < 2L || chip_cortes[1] >= chip_cortes[2]) {
    chip_cortes <- c(60, max(75, corte_score_val))
  }

  sem_keys <- c("rojo", "ambar", "verde")
  stats_df$score_round <- .dim_round_half_up(stats_df$score_mean, 0)
  stats_df$sem_key <- ifelse(
    stats_df$score_mean >= chip_cortes[2], sem_keys[3],
    ifelse(stats_df$score_mean >= chip_cortes[1], sem_keys[2], sem_keys[1])
  )
  stats_df$sem_color <- ifelse(
    stats_df$score_mean >= chip_cortes[2], sem$verde,
    ifelse(stats_df$score_mean >= chip_cortes[1], sem$ambar, sem$rojo)
  )
  stats_df$score_sd_plot <- ifelse(is.na(stats_df$score_sd), 0, stats_df$score_sd)

  # --- Utilidades tipográficas ---
  .foda_trunc <- function(x, max_chars = 28L) {
    x <- trimws(as.character(x %||% ""))
    max_chars <- max(6L, as.integer(max_chars)[1])
    ifelse(
      nchar(x, type = "width") <= max_chars,
      x,
      paste0(substr(x, 1L, max_chars - 3L), "...")
    )
  }
  .wrap_item_label <- function(x, width = 24L, max_lines = 2L) {
    width <- max(10L, as.integer(width)[1])
    max_lines <- max(1L, as.integer(max_lines)[1])
    x <- as.character(x %||% "")
    wrapped <- if (requireNamespace("stringr", quietly = TRUE)) {
      stringr::str_wrap(x, width = width)
    } else {
      vapply(x, function(xx) paste(strwrap(xx, width = width), collapse = "\n"), character(1))
    }
    out <- vapply(wrapped, function(xx) {
      ln <- strsplit(xx, "\n", fixed = TRUE)[[1]]
      if (length(ln) <= max_lines) return(paste(ln, collapse = "\n"))
      ln <- ln[seq_len(max_lines)]
      ln[max_lines] <- .foda_trunc(ln[max_lines], max_chars = width)
      paste(ln, collapse = "\n")
    }, character(1))
    out
  }
  stats_df <- stats_df[order(-stats_df$score_mean, stats_df$grupo, stats_df$label), , drop = FALSE]
  legend_cruce_labels <- character(0)
  legend_cruce_colors <- character(0)

  # --- Construcción del panel según modo ---
  if (identical(modo_foda, "matriz")) {
    cuadrantes_cfg <- data.frame(
      cuadrante = c("fortaleza", "oportunidad", "debilidad", "amenaza"),
      titulo    = c(
        titulos_out[["fortaleza"]],
        titulos_out[["oportunidad"]],
        titulos_out[["debilidad"]],
        titulos_out[["amenaza"]]
      ),
      subtexto  = c("Puntaje alto + consistente",
                     "Puntaje alto + disperso",
                     "Puntaje bajo + consistente",
                     "Puntaje bajo + disperso"),
      xmin = c(0, 1, 0, 1),
      xmax = c(1, 2, 1, 2),
      ymin = c(1, 1, 0, 0),
      ymax = c(2, 2, 1, 1),
      fill = c(col_f, col_o, col_d, col_a),
      tx   = c(0.5, 1.5, 0.5, 1.5),
      ty   = c(1.92, 1.92, 0.92, 0.92),
      stx  = c(0.5, 1.5, 0.5, 1.5),
      sty  = c(1.84, 1.84, 0.84, 0.84),
      stringsAsFactors = FALSE
    )

    items_df <- data.frame(
      x = numeric(0), y = numeric(0),
      w = numeric(0), h = numeric(0),
      title_x = numeric(0), title_y = numeric(0),
      title_txt = character(0),
      score_x = numeric(0), score_y = numeric(0),
      score_txt = character(0),
      chip_x = numeric(0), chip_w = numeric(0), chip_h = numeric(0),
      chip_fill = character(0), chip_text_col = character(0),
      card_fill = character(0), card_border = character(0), title_col = character(0),
      stringsAsFactors = FALSE
    )

    for (q in cuadrantes_cfg$cuadrante) {
      q_items <- stats_df[stats_df$cuadrante == q, , drop = FALSE]
      if (!nrow(q_items)) next

      cfg <- cuadrantes_cfg[cuadrantes_cfg$cuadrante == q, , drop = FALSE]
      n_cols <- if (nrow(q_items) <= 4L) 1L else 2L
      max_rows <- if (n_cols == 1L) 6L else 5L
      max_show <- n_cols * max_rows
      n_show <- min(nrow(q_items), max_show)

      idx_show <- seq_len(n_show)
      row_ids <- ((idx_show - 1L) %/% n_cols) + 1L
      col_ids <- ((idx_show - 1L) %% n_cols) + 1L

      if (n_cols == 1L) {
        card_w_base <- min(ancho_tarjeta_matriz, 0.82)
        card_w_max <- card_w_base
        x_centers <- cfg$xmin + 0.5
      } else {
        gap_x <- 0.028
        side_min <- 0.045
        card_w_base <- min(ancho_tarjeta_matriz, (1 - (2 * side_min) - gap_x) / 2)
        card_w_max <- card_w_base
        side_pad <- max(side_min, (1 - (2 * card_w_base) - gap_x) / 2)
        x_centers <- c(cfg$xmin + side_pad + card_w_base / 2, cfg$xmax - side_pad - card_w_base / 2)
      }

      y_top <- cfg$ymax - 0.22
      y_bottom <- cfg$ymin + 0.08
      n_rows_show <- max(row_ids)
      gap_y <- if (n_rows_show <= 2L) 0.028 else 0.020
      area_h <- max(0.22, y_top - y_bottom)
      card_h <- min(0.22, (area_h - gap_y * (n_rows_show - 1L)) / n_rows_show)
      card_h <- max(0.115, card_h)
      stack_h <- n_rows_show * card_h + (n_rows_show - 1L) * gap_y
      y_mid <- (y_top + y_bottom) / 2
      y_start <- y_mid + stack_h / 2 - card_h / 2
      y_vals <- y_start - (row_ids - 1L) * (card_h + gap_y)
      x_vals <- x_centers[col_ids]

      score_txt_vec <- paste0(.dim_fmt_int(q_items$score_round[idx_show]), sufijo_puntaje)
      pad_left <- 0.026
      pad_right <- pad_left
      gap_title_chip <- 0.012
      card_w <- rep(card_w_base, n_show)
      if (isTRUE(ancho_recuadro_auto)) {
        lbl_vec <- as.character(q_items$label[idx_show] %||% "")
        chars_lbl <- nchar(lbl_vec, type = "width")
        chip_need_auto <- 0.028 + pmax(1, nchar(score_txt_vec, type = "width")) * 0.0085
        txt_need_auto <- 0.030 + pmin(chars_lbl, 44) * 0.0062
        w_need <- pad_left + txt_need_auto + gap_title_chip + chip_need_auto + pad_right
        card_w <- pmin(card_w_max, pmax(0.235, w_need))
      }
      chip_h <- card_h * 0.72
      chip_w_target <- card_w * .dim_clamp(ancho_chip_rel, 0.10, 0.55)
      chip_w_need <- 0.028 + pmax(1, nchar(score_txt_vec, type = "width")) * 0.0085
      chip_w <- pmin(card_w * 0.52, pmax(chip_w_target, chip_w_need))
      left_edge <- x_vals - (card_w / 2)
      right_edge <- x_vals + (card_w / 2)
      chip_x <- right_edge - (chip_w / 2) - pad_right
      title_x <- left_edge + pad_left
      score_x <- chip_x
      score_y <- y_vals
      title_y <- y_vals
      text_w <- pmax(0.18, (chip_x - (chip_w / 2)) - title_x - gap_title_chip)
      wrap_w <- pmax(11L, as.integer(floor(text_w * 64)))

      q_df <- data.frame(
        x = x_vals,
        y = y_vals,
        w = card_w,
        h = rep(card_h, n_show),
        title_x = title_x,
        title_y = title_y,
        title_txt = mapply(
          function(tt, ww) .wrap_item_label(tt, width = ww, max_lines = 2L),
          q_items$label[idx_show], wrap_w, USE.NAMES = FALSE
        ),
        score_x = score_x,
        score_y = score_y,
        score_txt = score_txt_vec,
        chip_x = chip_x,
        chip_w = chip_w,
        chip_h = rep(chip_h, n_show),
        chip_fill = as.character(q_items$sem_color[idx_show]),
        chip_text_col = rep("#FFFFFF", n_show),
        card_fill = rep("#FFFFFF", n_show),
        card_border = rep("#7C90A6", n_show),
        title_col = rep("#0D243E", n_show),
        stringsAsFactors = FALSE
      )
      if (nrow(q_items) > n_show && nrow(q_df)) {
        j <- nrow(q_df)
        q_df$title_txt[j] <- paste0("+", nrow(q_items) - n_show, " más")
        q_df$score_txt[j] <- ""
        q_df$chip_fill[j] <- "#E3EAF3"
        q_df$chip_text_col[j] <- "#4D6178"
        q_df$card_fill[j] <- "#F8FAFD"
        q_df$card_border[j] <- "#C2CFDC"
        q_df$title_col[j] <- "#546678"
      }

      items_df <- rbind(items_df, q_df)
    }

    p_panel <- ggplot2::ggplot() +
      ggplot2::geom_rect(
        data = cuadrantes_cfg,
        ggplot2::aes(xmin = .data$xmin, xmax = .data$xmax, ymin = .data$ymin, ymax = .data$ymax),
        fill = cuadrantes_cfg$fill,
        colour = "#D4DEE9",
        linewidth = 0.55
      ) +
      ggplot2::annotate(
        "segment",
        x = 1, xend = 1, y = 0, yend = 2,
        colour = "#7C8FA4",
        linewidth = 0.6,
        linetype = "22"
      ) +
      ggplot2::annotate(
        "segment",
        x = 0, xend = 2, y = 1, yend = 1,
        colour = "#7C8FA4",
        linewidth = 0.6,
        linetype = "22"
      ) +
      ggplot2::annotate(
        "text",
        x = cuadrantes_cfg$tx,
        y = cuadrantes_cfg$ty,
        label = cuadrantes_cfg$titulo,
        fontface = "bold",
        size = size_cuadrante_titulo / 3,
        colour = color_cuadrante_titulo,
        hjust = 0.5, vjust = 1
      )

    if (isTRUE(sd_tecnico)) {
      p_panel <- p_panel +
        ggplot2::annotate(
          "label",
          x = 1,
          y = 1.985,
          label = paste0("Corte SD = ", format(round(corte_sd_val, 1), trim = TRUE)),
          fill = "#EEF3F8",
          colour = "#4A5F75",
          linewidth = 0,
          size = (size_items - 1) / 3,
          vjust = 1
        )
    }

    p_panel <- p_panel +
      ggplot2::annotate(
        "label",
        x = 0.012,
        y = 1,
        label = paste0("Corte puntaje = ", format(round(corte_score_val, 1), trim = TRUE)),
        fill = "#EEF3F8",
        colour = "#4A5F75",
        linewidth = 0,
        size = (size_items - 1) / 3,
        hjust = 0,
        vjust = -0.2
      )

    if (isTRUE(mostrar_subtitulo_area)) {
      p_panel <- p_panel +
        ggplot2::annotate(
          "text",
          x = cuadrantes_cfg$stx,
          y = cuadrantes_cfg$sty,
          label = cuadrantes_cfg$subtexto,
          fontface = "plain",
          size = (size_items - 1) / 3.2,
          colour = "#5D6F82",
          hjust = 0.5, vjust = 1
        )
    }

    if (nrow(items_df)) {
      p_panel <- p_panel +
        ggplot2::geom_tile(
          data = items_df,
          ggplot2::aes(
            x = .data$x, y = .data$y,
            width = .data$w, height = .data$h,
            fill = .data$card_fill, colour = .data$card_border
          ),
          linewidth = 0.62
        ) +
        ggplot2::geom_tile(
          data = items_df,
          ggplot2::aes(
            x = .data$chip_x, y = .data$y,
            width = .data$chip_w, height = .data$chip_h,
            fill = .data$chip_fill
          ),
          colour = "#2E425A",
          linewidth = 0.25
        ) +
        ggplot2::geom_text(
          data = items_df,
          ggplot2::aes(x = .data$title_x, y = .data$title_y, label = .data$title_txt, colour = .data$title_col),
          size = tamano_texto_tarjeta / 1.95,
          hjust = 0,
          vjust = 0.5,
          lineheight = 0.92,
          fontface = "bold"
        ) +
        ggplot2::geom_text(
          data = items_df,
          ggplot2::aes(x = .data$score_x, y = .data$score_y, label = .data$score_txt, colour = .data$chip_text_col),
          size = tamano_texto_chip / 2.55,
          hjust = 0.5,
          vjust = 0.5,
          lineheight = 1,
          fontface = "bold"
        ) +
        ggplot2::scale_fill_identity() +
        ggplot2::scale_colour_identity()
    }

    p_panel <- p_panel +
      ggplot2::coord_cartesian(xlim = c(0, 2), ylim = c(0, 2), expand = FALSE)
  } else {
    plot_df <- stats_df[order(stats_df$score_mean, stats_df$score_sd_plot, stats_df$grupo, stats_df$label), , drop = FALSE]
    n_pts <- nrow(plot_df)

    if (nzchar(cruce)) {
      grp_ref <- unique(plot_df[, c("grupo", "grupo_key"), drop = FALSE])
      grp_cols <- .dim_group_colors(
        groups = as.character(grp_ref$grupo),
        paleta_radar = ctx$paleta_radar,
        total_color = color_indice_total,
        palette_override = ctx$paletas_cruce[[cruce]] %||% NULL,
        group_keys = as.character(grp_ref$grupo_key)
      )
      plot_df$is_total <- !is.na(plot_df$is_total_global) & as.logical(plot_df$is_total_global)
      plot_df$card_fill <- as.character(grp_cols[plot_df$grupo])
      plot_df$card_fill[!nzchar(plot_df$card_fill) | is.na(plot_df$card_fill)] <- "#2F4A66"
      idx_total <- which(!is.na(plot_df$is_total) & plot_df$is_total)
      if (length(idx_total)) plot_df$card_fill[idx_total] <- color_indice_total

      grp_ref$is_total <- (as.character(grp_ref$grupo_key) == "__total__") |
        (tolower(trimws(as.character(grp_ref$grupo))) %in% c("indice", "índice"))
      grp_ref$color_leg <- as.character(grp_cols[as.character(grp_ref$grupo)])
      keep_legend <- !grp_ref$is_total
      legend_cruce_labels <- as.character(grp_ref$grupo[keep_legend])
      legend_cruce_colors <- as.character(grp_ref$color_leg[keep_legend])
      ok_leg <- nzchar(legend_cruce_labels) & !is.na(legend_cruce_colors) & nzchar(legend_cruce_colors)
      legend_cruce_labels <- legend_cruce_labels[ok_leg]
      legend_cruce_colors <- legend_cruce_colors[ok_leg]
    } else {
      plot_df$card_fill <- if (isTRUE(tarjetas_color_solido)) "#2F4A66" else "#FFFFFF"
    }
    plot_df$card_border <- grDevices::adjustcolor(plot_df$card_fill, alpha.f = 0.92)
    plot_df$title_col <- ifelse(
      vapply(plot_df$card_fill, .is_light_col, logical(1)),
      "#0D243E",
      "#FFFFFF"
    )
    plot_df$chip_text_col <- "#FFFFFF"
    plot_df$score_txt <- paste0(.dim_fmt_int(plot_df$score_round), sufijo_puntaje)

    # --- Normalización centrada en cortes para cuadrantes visualmente equivalentes ---
    x_raw <- plot_df$score_sd_plot
    y_raw <- plot_df$score_mean
    x_min_obs <- suppressWarnings(min(c(x_raw, corte_sd_val), na.rm = TRUE))
    x_max_obs <- suppressWarnings(max(c(x_raw, corte_sd_val), na.rm = TRUE))
    y_min_obs <- suppressWarnings(min(c(y_raw, corte_score_val), na.rm = TRUE))
    y_max_obs <- suppressWarnings(max(c(y_raw, corte_score_val), na.rm = TRUE))

    x_span_lo <- max(corte_sd_val - x_min_obs, 0.35)
    x_span_hi <- max(x_max_obs - corte_sd_val, 0.35)
    y_span_lo <- max(corte_score_val - y_min_obs, 3.0)
    y_span_hi <- max(y_max_obs - corte_score_val, 3.0)

    x_min_ref <- max(0, corte_sd_val - x_span_lo * 1.18)
    x_max_ref <- corte_sd_val + x_span_hi * 1.18
    y_min_ref <- max(0, corte_score_val - y_span_lo * 1.15)
    y_max_ref <- min(100, corte_score_val + y_span_hi * 1.15)

    .map_norm <- function(v, cut, lo, hi) {
      v <- as.numeric(v)
      out <- rep(NA_real_, length(v))
      d_lo <- max(cut - lo, 1e-6)
      d_hi <- max(hi - cut, 1e-6)
      i_lo <- which(v <= cut)
      i_hi <- which(v > cut)
      if (length(i_lo)) out[i_lo] <- -1 + (v[i_lo] - lo) / d_lo
      if (length(i_hi)) out[i_hi] <- (v[i_hi] - cut) / d_hi
      pmax(-1, pmin(1, out))
    }
    .x_map <- function(v) .map_norm(v, cut = corte_sd_val, lo = x_min_ref, hi = x_max_ref)
    .y_map <- function(v) .map_norm(v, cut = corte_score_val, lo = y_min_ref, hi = y_max_ref)

    plot_df$x_base <- .x_map(x_raw)
    plot_df$y_base <- .y_map(y_raw)
    x_lim <- c(-1, 1)
    y_lim <- c(-1, 1)

    cuadrantes_cfg <- data.frame(
      cuadrante = c("fortaleza", "oportunidad", "debilidad", "amenaza"),
      titulo    = c(
        titulos_out[["fortaleza"]],
        titulos_out[["oportunidad"]],
        titulos_out[["debilidad"]],
        titulos_out[["amenaza"]]
      ),
      subtexto  = c("Puntaje alto + consistente",
                    "Puntaje alto + disperso",
                    "Puntaje bajo + consistente",
                    "Puntaje bajo + disperso"),
      xmin = c(-1, 0, -1, 0),
      xmax = c(0, 1, 0, 1),
      ymin = c(0, 0, -1, -1),
      ymax = c(1, 1, 0, 0),
      fill = c(col_f, col_o, col_d, col_a),
      stringsAsFactors = FALSE
    )
    cuadrantes_cfg$tx <- (cuadrantes_cfg$xmin + cuadrantes_cfg$xmax) / 2
    cuadrantes_cfg$ty <- cuadrantes_cfg$ymin + (cuadrantes_cfg$ymax - cuadrantes_cfg$ymin) * 0.92
    cuadrantes_cfg$stx <- cuadrantes_cfg$tx
    cuadrantes_cfg$sty <- cuadrantes_cfg$ymin + (cuadrantes_cfg$ymax - cuadrantes_cfg$ymin) * 0.80

    card_w_base <- .dim_clamp(ancho_tarjeta_disp * factor_reduccion_tarjeta_dispersion * 0.42, 0.11, 0.31)
    pad_left <- 0.012
    pad_right <- 0.006
    gap_title_chip <- 0.010
    card_w <- rep(card_w_base, n_pts)
    if (isTRUE(ancho_recuadro_auto) && n_pts > 0L) {
      label_seed <- as.character(plot_df$label %||% "")
      if (nzchar(cruce)) {
        is_total_lbl <- !is.na(plot_df$is_total) & plot_df$is_total
        if (identical(disposicion_recuadro, "dos_lineas")) {
          label_seed <- ifelse(
            is_total_lbl,
            as.character(plot_df$label),
            ifelse(
              nchar(as.character(plot_df$grupo), type = "width") > nchar(as.character(plot_df$label), type = "width"),
              as.character(plot_df$grupo), as.character(plot_df$label)
            )
          )
        } else if (identical(disposicion_recuadro, "sin_cruce")) {
          label_seed <- as.character(plot_df$label)
        } else {
          label_seed <- ifelse(is_total_lbl, as.character(plot_df$label), paste0(as.character(plot_df$label), " · ", as.character(plot_df$grupo)))
        }
      }
      chars_lbl <- nchar(label_seed, type = "width")
      chip_need_auto <- 0.028 + pmax(1, nchar(plot_df$score_txt, type = "width")) * 0.0080
      txt_need_auto <- 0.026 + pmin(chars_lbl, 52) * 0.0058
      w_need <- pad_left + txt_need_auto + gap_title_chip + chip_need_auto + pad_right
      card_w <- pmax(0.12, pmin(card_w_base, w_need))
    }
    card_h <- .dim_clamp((if (n_pts > 26L) 0.082 else 0.095) * factor_reduccion_tarjeta_dispersion, 0.055, 0.12)

    if (n_pts > 0) {
      idx <- seq_len(n_pts)
      plot_df$x_card <- plot_df$x_base + jitter_x_rel * sin(idx * 2.399 + 0.7)
      plot_df$y_card <- plot_df$y_base + jitter_y_rel * cos(idx * 1.913 + 0.2)
    } else {
      plot_df$x_card <- numeric(0)
      plot_df$y_card <- numeric(0)
    }

    pad_x <- card_w / 2 + 0.012
    pad_y <- card_h / 2 + 0.012
    title_band_q <- 0.20
    q_ymax <- stats::setNames(cuadrantes_cfg$ymax, cuadrantes_cfg$cuadrante)
    y_cap_by_q <- as.numeric(q_ymax[as.character(plot_df$cuadrante)]) - title_band_q - (card_h / 2)
    y_cap_by_q[!is.finite(y_cap_by_q)] <- y_lim[2] - pad_y
    y_cap_by_q <- pmin(y_cap_by_q, y_lim[2] - pad_y)
    if (n_pts > 0) {
      plot_df$x_card <- pmin(pmax(plot_df$x_card, x_lim[1] + pad_x), x_lim[2] - pad_x)
      plot_df$y_card <- pmin(pmax(plot_df$y_card, y_lim[1] + pad_y), y_cap_by_q)
    }

    # Empuje iterativo simple para reducir colisiones.
    if (n_pts > 1L && iter_separacion > 0L) {
      xs <- as.numeric(plot_df$x_card)
      ys <- as.numeric(plot_df$y_card)
      for (it in seq_len(iter_separacion)) {
        for (i in seq_len(n_pts - 1L)) {
          for (j in seq.int(i + 1L, n_pts)) {
            dx <- xs[j] - xs[i]
            dy <- ys[j] - ys[i]
            min_dx <- (card_w[i] + card_w[j]) / 2 + 0.014
            min_dy <- card_h + 0.010
            if (abs(dx) < min_dx && abs(dy) < min_dy) {
              sx <- ifelse(dx >= 0, 1, -1)
              sy <- ifelse(dy >= 0, 1, -1)
              move_x <- (min_dx - abs(dx)) * 0.33
              move_y <- (min_dy - abs(dy)) * 0.46
              xs[i] <- xs[i] - sx * move_x
              xs[j] <- xs[j] + sx * move_x
              ys[i] <- ys[i] - sy * move_y
              ys[j] <- ys[j] + sy * move_y
            }
          }
        }
        xs <- pmin(pmax(xs, x_lim[1] + pad_x), x_lim[2] - pad_x)
        ys <- pmin(pmax(ys, y_lim[1] + pad_y), y_cap_by_q)
      }
      plot_df$x_card <- xs
      plot_df$y_card <- ys
    }

    chip_h <- card_h * 0.66
    chip_w_target <- card_w * .dim_clamp(ancho_chip_rel, 0.10, 0.58)
    chip_w_need <- 0.028 + pmax(1, nchar(plot_df$score_txt, type = "width")) * 0.0080
    chip_w <- pmin(card_w * 0.56, pmax(chip_w_target, chip_w_need))
    plot_df$card_w <- card_w
    plot_df$chip_w <- chip_w
    plot_df$chip_h <- chip_h
    plot_df$chip_x <- plot_df$x_card + (card_w / 2) - (chip_w / 2) - pad_right
    plot_df$title_x <- plot_df$x_card - (card_w / 2) + pad_left
    txt_w <- pmax(0.12, (plot_df$chip_x - chip_w / 2) - plot_df$title_x - gap_title_chip)
    wrap_w <- pmax(9L, as.integer(floor(txt_w * 82)))

    if (nzchar(cruce)) {
      is_total_lbl <- !is.na(plot_df$is_total) & plot_df$is_total
      if (identical(disposicion_recuadro, "dos_lineas")) {
        line_1 <- mapply(function(tt, ww) .foda_trunc(tt, max_chars = max(12L, ww + 1L)), plot_df$label, wrap_w, USE.NAMES = FALSE)
        line_2 <- mapply(function(tt, ww) .foda_trunc(tt, max_chars = max(9L, ww)), plot_df$grupo, wrap_w, USE.NAMES = FALSE)
        title_txt <- paste0(line_1, "\n", line_2)
      } else if (identical(disposicion_recuadro, "sin_cruce")) {
        title_txt <- mapply(
          function(tt, ww) .foda_trunc(tt, max_chars = max(14L, ww + 3L)),
          plot_df$label, wrap_w,
          USE.NAMES = FALSE
        )
      } else {
        title_txt <- mapply(
          function(tt, gg, ww) .foda_trunc(paste0(tt, " · ", gg), max_chars = max(14L, ww + 3L)),
          plot_df$label, plot_df$grupo, wrap_w,
          USE.NAMES = FALSE
        )
      }
      if (any(is_total_lbl)) {
        title_txt[is_total_lbl] <- mapply(
          function(tt, ww) .foda_trunc(tt, max_chars = max(14L, ww + 3L)),
          plot_df$label[is_total_lbl], wrap_w[is_total_lbl],
          USE.NAMES = FALSE
        )
      }
      plot_df$title_txt <- title_txt
    } else {
      plot_df$title_txt <- mapply(
        function(tt, ww) .wrap_item_label(tt, width = ww, max_lines = 2L),
        plot_df$label, wrap_w, USE.NAMES = FALSE
      )
    }
    size_title_eff <- pmax(2.0, pmin(tamano_texto_tarjeta / 2.30, card_h * 28))
    size_chip_eff <- pmax(2.4, pmin(tamano_texto_chip / 2.45, chip_h * 27))

    if (isTRUE(sd_tecnico)) {
      x_break_vals <- sort(unique(c(pretty(c(x_min_ref, x_max_ref), n = 5), corte_sd_val)))
      x_break_vals <- x_break_vals[is.finite(x_break_vals) & x_break_vals >= x_min_ref & x_break_vals <= x_max_ref]
      x_breaks <- .x_map(x_break_vals)
      x_labels <- format(round(x_break_vals, 1), trim = TRUE)
      x_expand_mult <- c(0, 0)
    } else {
      x_breaks <- c(x_lim[1] + 0.03, x_lim[2] - 0.03)
      x_labels <- c("Menor\ndispersión", "Mayor\ndispersión")
      x_expand_mult <- c(0.10, 0.10)
    }

    y_break_vals <- sort(unique(c(pretty(c(y_min_ref, y_max_ref), n = 5), corte_score_val)))
    y_break_vals <- y_break_vals[is.finite(y_break_vals) & y_break_vals >= y_min_ref & y_break_vals <= y_max_ref]
    y_breaks <- .y_map(y_break_vals)
    y_labels <- format(round(y_break_vals, 1), trim = TRUE)

    p_panel <- ggplot2::ggplot() +
      ggplot2::geom_rect(
        data = cuadrantes_cfg,
        ggplot2::aes(xmin = .data$xmin, xmax = .data$xmax, ymin = .data$ymin, ymax = .data$ymax),
        fill = cuadrantes_cfg$fill,
        colour = "#D4DEE9",
        linewidth = 0.55
      ) +
      ggplot2::annotate(
        "segment",
        x = 0, xend = 0, y = y_lim[1], yend = y_lim[2],
        colour = "#7C8FA4", linewidth = 0.6, linetype = "22"
      ) +
      ggplot2::annotate(
        "segment",
        x = x_lim[1], xend = x_lim[2], y = 0, yend = 0,
        colour = "#7C8FA4", linewidth = 0.6, linetype = "22"
      ) +
      ggplot2::annotate(
        "text",
        x = cuadrantes_cfg$tx,
        y = cuadrantes_cfg$ty,
        label = cuadrantes_cfg$titulo,
        fontface = "bold",
        size = size_cuadrante_titulo / 3,
        colour = color_cuadrante_titulo,
        hjust = 0.5, vjust = 1
      )

    if (isTRUE(sd_tecnico)) {
      p_panel <- p_panel +
        ggplot2::annotate(
          "label",
          x = 0,
          y = y_lim[2] - 0.015,
          label = paste0("Corte SD = ", format(round(corte_sd_val, 1), trim = TRUE)),
          fill = "#EEF3F8",
          colour = "#4A5F75",
          linewidth = 0,
          size = (size_items - 1) / 3,
          vjust = 1
        )
    }

    p_panel <- p_panel +
      ggplot2::annotate(
        "label",
        x = x_lim[1] + 0.012,
        y = 0,
        label = paste0("Corte puntaje = ", format(round(corte_score_val, 1), trim = TRUE)),
        fill = "#EEF3F8",
        colour = "#4A5F75",
        linewidth = 0,
        size = (size_items - 1) / 3,
        hjust = 0,
        vjust = -0.2
      )

    if (isTRUE(mostrar_subtitulo_area)) {
      p_panel <- p_panel +
        ggplot2::annotate(
          "text",
          x = cuadrantes_cfg$stx,
          y = cuadrantes_cfg$sty,
          label = cuadrantes_cfg$subtexto,
          fontface = "plain",
          size = (size_items - 1) / 3.2,
          colour = "#5D6F82",
          hjust = 0.5, vjust = 1
        )
    }

    if (nrow(plot_df)) {
      p_panel <- p_panel +
        ggplot2::geom_tile(
          data = plot_df,
          ggplot2::aes(x = .data$x_card, y = .data$y_card, width = .data$card_w),
          height = card_h,
          fill = plot_df$card_fill,
          colour = plot_df$card_border,
          linewidth = 0.56
        ) +
        ggplot2::geom_tile(
          data = plot_df,
          ggplot2::aes(
            x = .data$chip_x, y = .data$y_card,
            fill = .data$sem_color,
            width = .data$chip_w,
            height = .data$chip_h
          ),
          colour = "#24394F",
          linewidth = 0.23
        ) +
        ggplot2::geom_text(
          data = plot_df,
          ggplot2::aes(x = .data$title_x, y = .data$y_card, label = .data$title_txt, colour = .data$title_col),
          size = size_title_eff,
          hjust = 0,
          vjust = 0.5,
          lineheight = 0.86,
          fontface = "bold"
        ) +
        ggplot2::geom_text(
          data = plot_df,
          ggplot2::aes(x = .data$chip_x, y = .data$y_card, label = .data$score_txt, colour = .data$chip_text_col),
          size = size_chip_eff,
          hjust = 0.5,
          vjust = 0.5,
          lineheight = 1,
          fontface = "bold"
        ) +
        ggplot2::scale_fill_identity() +
        ggplot2::scale_colour_identity()
    }

    p_panel <- p_panel +
      ggplot2::coord_cartesian(xlim = x_lim, ylim = y_lim, expand = FALSE, clip = "on") +
      ggplot2::scale_x_continuous(
        breaks = x_breaks,
        labels = x_labels,
        limits = x_lim,
        expand = ggplot2::expansion(mult = x_expand_mult)
      ) +
      ggplot2::scale_y_continuous(
        breaks = y_breaks,
        labels = y_labels,
        limits = y_lim,
        expand = ggplot2::expansion(mult = 0)
      ) +
      ggplot2::labs(x = if (isTRUE(sd_tecnico)) "Desviación estándar" else NULL, y = "Puntaje")
  }

  if (identical(modo_foda, "matriz")) {
    p_panel <- p_panel +
      ggplot2::theme_void() +
      ggplot2::theme(
        plot.background = ggplot2::element_rect(fill = color_fondo, colour = NA),
        panel.background = ggplot2::element_rect(fill = color_fondo, colour = NA),
        plot.margin = ggplot2::margin(4, 8, 4, 8)
      )
  } else {
    x_text_size <- max(7, size_items - 1)
    x_text_margin <- if (!isTRUE(sd_tecnico)) ggplot2::margin(t = 7, r = 6, b = 0, l = 6) else ggplot2::margin(t = 4)
    pm_right <- if (!isTRUE(sd_tecnico)) 30 else 8
    pm_left <- if (!isTRUE(sd_tecnico)) 30 else 8
    p_panel <- p_panel +
      ggplot2::theme_minimal(base_size = max(8, size_items)) +
      ggplot2::theme(
        plot.background = ggplot2::element_rect(fill = color_fondo, colour = NA),
        panel.background = ggplot2::element_rect(fill = color_fondo, colour = NA),
        panel.grid.minor = ggplot2::element_blank(),
        panel.grid.major = ggplot2::element_line(colour = "#DCE5EF", linewidth = 0.35),
        axis.title = ggplot2::element_text(colour = "#1B314A", size = max(8, size_items)),
        axis.text.y = ggplot2::element_text(colour = "#334A63", size = max(7, size_items - 1)),
        axis.text.x = ggplot2::element_text(
          colour = "#334A63",
          size = x_text_size,
          margin = x_text_margin,
          lineheight = if (!isTRUE(sd_tecnico)) 0.95 else 1
        ),
        axis.ticks = ggplot2::element_line(colour = "#8AA0B7", linewidth = 0.25),
        plot.margin = ggplot2::margin(4, pm_right, 6, pm_left)
      )
  }

  # --- Non-canvas mode ---
  if (!isTRUE(usar_canvas)) {
    return(.dim_export_canvas(
      p_panel,
      exportar = exportar,
      path_salida = path_salida,
      ancho = ancho, alto = alto, dpi = dpi,
      ppt_append = ppt_append,
      ppt_layout = ppt_layout,
      ppt_master = ppt_master
    ))
  }

  # --- Canvas mode ---
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
    if (identical(modo_foda, "dispersion") && nzchar(cruce) && length(legend_cruce_labels)) {
      .dim_heat_legend_block(
        labels = legend_cruce_labels,
        colors = legend_cruce_colors,
        size = max(7, size_items),
        colour = "#4A5F75"
      )
    } else {
      .dim_heat_legend_block(
        labels = c("Rojo", "Ambar", "Verde"),
        colors = c(sem$rojo, sem$ambar, sem$verde),
        size = max(7, size_items),
        colour = "#4A5F75"
      )
    }
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

  h_title   <- canvas_h_title
  h_legend  <- if (isTRUE(mostrar_leyenda)) canvas_h_legend else 0.01
  h_caption <- if (!is.null(nota_pie) && nzchar(nota_pie)) canvas_h_caption else 0.01
  h_panel   <- max(0.01, 1 - (h_title + h_legend + h_caption) - canvas_pad_top)

  canvas <- cowplot::plot_grid(
    .dim_wrap_debug_canvas(title_block,   debug_ph_bordes, debug_ph_col, debug_ph_lwd),
    .dim_wrap_debug_canvas(p_panel,       debug_ph_bordes, debug_ph_col, debug_ph_lwd),
    .dim_wrap_debug_canvas(legend_block,  debug_ph_bordes, debug_ph_col, debug_ph_lwd),
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
