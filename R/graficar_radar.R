# =============================================================================
# graficar_radar() — plot-ready estilo prosecnur (canvas + export) + TABLA DERECHA
# FIXES:
# A) PPT (rvg) NO ABORTA:
#    - En el objeto que va a rvg::dml() NO se usa geom_polygon() (ni en malla ni en fill)
#    - clip="on" + límites recalculados para que NADA quede fuera del viewport
#    - sanitización: se eliminan coords no finitas antes de dibujar
#
# B) TABLA (canvas):
#    - Auto-fit REAL (scale = min(w/gw, h/gh) * pad)
#    - Placeholder con clip=on para que nunca se salga
#    - Headers centrados (incluye 1ra celda) y cuerpo: 1ra col izquierda, demás centradas
# =============================================================================

#' @export
graficar_radar <- function(
    data,
    var_eje   = "eje",
    var_grupo = "grupo",
    var_valor = "valor",

    escala_valor  = c("proporcion_1", "proporcion_100"),
    limites       = NULL,
    cortes_grilla = 5L,

    mostrar_tela    = TRUE,
    mostrar_radios  = TRUE,
    mostrar_niveles = TRUE,

    wrap_ejes = 24,
    eje_label_mult = 1.14,

    mostrar_puntos = TRUE,
    size_linea     = 0.9,
    alpha_relleno  = 0.18,
    size_punto     = 2.2,

    rellenar_poligono = FALSE,   # ojo: en PPT se forzará FALSE (segfault rvg+polygon)

    etiquetas_series = NULL,     # named: old -> new
    colores_series   = NULL,     # named por etiqueta final

    mostrar_leyenda   = TRUE,
    leyenda_posicion  = c("abajo", "derecha"),
    legend_n_por_fila = 6L,

    legend_key_cm           = 0.35,
    legend_espaciado        = 0.25, # pt
    legend_key_spacing_x_cm = 0.10,

    titulo       = NULL,
    subtitulo    = NULL,
    nota_pie     = NULL,
    pos_titulo   = c("centro","izquierda","derecha"),
    pos_nota_pie = c("derecha","izquierda","centro"),
    textos_negrita = NULL,

    color_titulo    = "#004B8D",
    size_titulo     = 12,
    color_subtitulo = "#004B8D",
    size_subtitulo  = 9,
    color_nota_pie  = "#004B8D",
    size_nota_pie   = 8,
    color_leyenda   = "#004B8D",
    size_leyenda    = 8,

    color_ejes = "#004B8D",
    size_ejes  = 10,

    color_grilla = "#DDDDDD",
    color_radios = "#DDDDDD",
    color_fondo  = NA,

    # -------------------------------------------------------------------------
    # TABLA (derecha)
    # -------------------------------------------------------------------------
    mostrar_tabla_derecha = FALSE,
    titulo_tabla = "TOP TWO BOX",
    umbral_rojo_pct = 60,
    tabla_digits = 0L,

    tabla_header_fill = "#062A63",
    tabla_body_fill   = "#F2F2F2",
    tabla_grid_col    = "white",
    tabla_text_blue   = "#062A63",
    tabla_font_family = "Arial",

    tabla_header_size = 14,
    tabla_body_size   = 12,
    tabla_firstcol_bold = TRUE,

    tabla_padding_mm = 3,

    tabla_ph_ancho = 0.40,
    tabla_ph_gap   = 0.03,
    tabla_ph_margin_top = 0.04,
    tabla_ph_margin_bot = 0.06,

    tabla_auto_fit = FALSE,
    tabla_fit_pad   = 0.98,
    tabla_allow_upscale = FALSE,
    tabla_clip      = TRUE,

    # -------------------------------------------------------------------------
    # CANVAS
    # -------------------------------------------------------------------------
    usar_canvas = FALSE,
    canvas_h_header_in  = 0.75,
    canvas_h_legend_in  = 0.75,
    canvas_h_caption_in = 0.40,
    canvas_h_panel_in   = NULL,
    alto_por_eje        = 0.32,
    encabezado_desplazamiento_in = 0,
    encabezado_separacion_in     = 0.14,
    leyenda_desplazamiento_in    = 0,
    centro_cowplot              = NA_real_,

    debug_ph_bordes = FALSE,
    debug_ph_col    = "#FF00FF",
    debug_ph_lwd    = 0.6,

    exportar    = c("rplot", "png", "ppt", "word"),
    path_salida = NULL,
    ancho       = 8.5,
    alto        = 6.5,
    dpi         = 300,

    ppt_append = TRUE,
    ppt_layout = "Blank",
    ppt_master = "Office Theme",

    # -------------------------------------------------------------------------
    # DEBUG PPT (callr / Rscript)
    # -------------------------------------------------------------------------
    debug_ppt = FALSE,
    debug_ppt_log = "radar_ppt_export_debug.log"
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y
  hjust_from_pos <- function(x) switch(x, "izquierda"=0, "centro"=0.5, "derecha"=1, 0.5)

  textos_negrita <- textos_negrita %||% character(0)

  # deps base
  if (!requireNamespace("ggplot2", quietly = TRUE)) stop("Requiere ggplot2.", call. = FALSE)
  if (!requireNamespace("dplyr", quietly = TRUE))  stop("Requiere dplyr.",  call. = FALSE)
  if (!requireNamespace("tidyr", quietly = TRUE))  stop("Requiere tidyr.",  call. = FALSE)
  if (!requireNamespace("grid", quietly = TRUE))   stop("Requiere grid.",   call. = FALSE)
  if (!requireNamespace("tibble", quietly = TRUE)) stop("Requiere tibble.", call. = FALSE)

  escala_valor     <- match.arg(escala_valor)
  exportar         <- match.arg(exportar)
  leyenda_posicion <- match.arg(leyenda_posicion)
  pos_titulo       <- match.arg(pos_titulo)
  pos_nota_pie     <- match.arg(pos_nota_pie)
  ppt_safe <- exportar %in% c("ppt","word", "rplot")

  if (!is.data.frame(data)) stop("`data` debe ser data.frame/tibble.", call. = FALSE)
  if (!all(c(var_eje, var_grupo, var_valor) %in% names(data))) {
    faltan <- setdiff(c(var_eje, var_grupo, var_valor), names(data))
    stop("Faltan columnas en `data`: ", paste(faltan, collapse = ", "), call. = FALSE)
  }

  # normalizaciones
  legend_n_por_fila <- suppressWarnings(as.integer(legend_n_por_fila))
  if (!is.finite(legend_n_por_fila) || legend_n_por_fila < 1L) legend_n_por_fila <- 6L

  legend_key_cm <- suppressWarnings(as.numeric(legend_key_cm))
  if (!is.finite(legend_key_cm) || legend_key_cm <= 0) legend_key_cm <- 0.35

  legend_espaciado <- suppressWarnings(as.numeric(legend_espaciado))
  if (!is.finite(legend_espaciado) || legend_espaciado < 0) legend_espaciado <- 0.25

  legend_key_spacing_x_cm <- suppressWarnings(as.numeric(legend_key_spacing_x_cm))
  if (!is.finite(legend_key_spacing_x_cm) || legend_key_spacing_x_cm < 0) legend_key_spacing_x_cm <- 0.10

  cortes_grilla <- suppressWarnings(as.integer(cortes_grilla))
  if (!is.finite(cortes_grilla) || cortes_grilla < 2L) cortes_grilla <- 5L

  wrap_ejes <- suppressWarnings(as.integer(wrap_ejes))
  if (!is.finite(wrap_ejes) || wrap_ejes < 0L) wrap_ejes <- 24L

  eje_label_mult <- suppressWarnings(as.numeric(eje_label_mult))
  if (!is.finite(eje_label_mult) || eje_label_mult <= 0) eje_label_mult <- 1.14

  # clamps tabla
  tabla_header_size <- suppressWarnings(as.numeric(tabla_header_size))
  if (!is.finite(tabla_header_size) || tabla_header_size <= 0) tabla_header_size <- 14
  tabla_body_size <- suppressWarnings(as.numeric(tabla_body_size))
  if (!is.finite(tabla_body_size) || tabla_body_size <= 0) tabla_body_size <- 12
  tabla_padding_mm <- suppressWarnings(as.numeric(tabla_padding_mm))
  if (!is.finite(tabla_padding_mm) || tabla_padding_mm < 0) tabla_padding_mm <- 3
  tabla_fit_pad <- suppressWarnings(as.numeric(tabla_fit_pad))
  if (!is.finite(tabla_fit_pad) || tabla_fit_pad <= 0 || tabla_fit_pad > 1.2) tabla_fit_pad <- 0.98

  hjust_titulo  <- hjust_from_pos(pos_titulo)
  hjust_caption <- hjust_from_pos(pos_nota_pie)

  # ---------------------------------------------------------------------------
  # Helpers: tabla Top Two Box
  # ---------------------------------------------------------------------------

  # A) construir data.frame (texto) para la tabla
  .make_tabla_ttb_df <- function(df_plot, ejes, grupos, digits = 0L, titulo_left = "TOP TWO BOX") {
    digits <- suppressWarnings(as.integer(digits))
    if (!is.finite(digits) || digits < 0L) digits <- 0L

    wide <- df_plot |>
      dplyr::transmute(
        eje   = as.character(.data$.eje),
        grupo = as.character(.data$.grupo),
        valor = as.numeric(.data$.valor)
      ) |>
      tidyr::complete(eje = ejes, grupo = grupos, fill = list(valor = 0)) |>
      tidyr::pivot_wider(names_from = "grupo", values_from = "valor")

    fmt_pct <- function(x) {
      x <- suppressWarnings(as.numeric(x))
      x[!is.finite(x) | is.na(x)] <- 0
      p <- round(x * 100, digits)
      if (digits == 0L) sprintf("%.0f%%", p) else sprintf(paste0("%.", digits, "f%%"), p)
    }

    out <- as.data.frame(wide)
    out[[1]] <- as.character(out[[1]])
    for (j in 2:ncol(out)) out[[j]] <- fmt_pct(out[[j]])
    names(out)[1] <- titulo_left
    out
  }

  # B) construir grob con estilo (tableGrob)
  .make_table_grob_ttb_style <- function(
    tb,
    header_fill = "#062A63",
    header_text = "white",
    body_fill   = "#F2F2F2",
    grid_col    = "white",
    text_blue   = "#062A63",
    font_family = "Arial",
    header_size = 14,
    body_size   = 12,
    firstcol_bold = TRUE,
    highlight_threshold = 60,
    highlight_col = "red",
    padding_mm = 3,
    firstcol_frac = 0.62
  ) {
    if (!requireNamespace("gridExtra", quietly = TRUE)) stop("Requiere gridExtra.", call. = FALSE)

    n_data <- nrow(tb)
    n_cols <- ncol(tb)

    firstcol_frac <- suppressWarnings(as.numeric(firstcol_frac))
    if (!is.finite(firstcol_frac)) firstcol_frac <- 0.62
    firstcol_frac <- max(0.40, min(0.80, firstcol_frac))

    tg <- gridExtra::tableGrob(
      tb,
      rows = NULL,
      theme = gridExtra::ttheme_minimal(
        base_size   = body_size,
        base_family = font_family,
        padding     = grid::unit(rep(padding_mm, 2), "mm"),
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

    # widths post-creation
    if (n_cols >= 2) {
      rest <- (1 - firstcol_frac) / (n_cols - 1)
      tg$widths <- grid::unit(c(firstcol_frac, rep(rest, n_cols - 1)), "npc")
    } else {
      tg$widths <- grid::unit(1, "npc")
    }

    # header centered
    for (j in seq_len(n_cols)) {
      k <- which(tg$layout$t == 1 & tg$layout$l == j & tg$layout$name == "colhead-fg")
      if (length(k)) {
        tg$grobs[[k]]$just <- "center"
        tg$grobs[[k]]$x <- grid::unit(0.5, "npc")
        tg$grobs[[k]]$gp <- grid::gpar(col = header_text, fontface = "bold", fontsize = header_size)
      }
    }

    # body centered; first col bold
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

    # highlight <= threshold
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

  .wrap_clip <- function(g) {
    grid::grobTree(
      g,
      vp = grid::viewport(
        x = 0.5, y = 0.5, width = 1, height = 1,
        just = c("center","center"),
        clip = "on"
      )
    )
  }

  # ---------------------------------------------------------------------------
  # 1) Preparar data plot-ready
  # ---------------------------------------------------------------------------
  df0 <- data |>
    dplyr::transmute(
      .eje   = as.character(.data[[var_eje]]),
      .grupo = as.character(.data[[var_grupo]]),
      .valor = suppressWarnings(as.numeric(.data[[var_valor]]))
    ) |>
    dplyr::filter(
      !is.na(.data$.eje), nzchar(.data$.eje),
      !is.na(.data$.grupo), nzchar(.data$.grupo)
    )

  if (!nrow(df0)) stop("`data` no tiene filas válidas para radar.", call. = FALSE)

  df0$.valor[!is.finite(df0$.valor) | is.na(df0$.valor)] <- 0
  if (escala_valor == "proporcion_100") df0$.valor <- df0$.valor / 100
  df0$.valor <- pmax(0, pmin(1, df0$.valor))

  if (!is.null(etiquetas_series) && length(etiquetas_series) > 0) {
    if (is.null(names(etiquetas_series))) stop("`etiquetas_series` debe ser nombrado: old -> new.", call. = FALSE)
    mp <- as.character(etiquetas_series)
    names(mp) <- as.character(names(etiquetas_series))
    df0$.grupo <- dplyr::recode(df0$.grupo, !!!mp)
  }

  ejes   <- unique(df0$.eje)
  grupos <- unique(df0$.grupo)

  if (length(ejes) < 3) stop("Radar requiere al menos 3 ejes.", call. = FALSE)
  if (length(grupos) < 1) stop("Radar requiere al menos 1 grupo.", call. = FALSE)

  df_plot <- df0 |>
    dplyr::mutate(
      .eje   = factor(.data$.eje,   levels = ejes),
      .grupo = factor(.data$.grupo, levels = grupos)
    ) |>
    tidyr::complete(.eje, .grupo, fill = list(.valor = 0)) |>
    dplyr::arrange(.grupo, .eje)

  lab_ejes <- levels(df_plot$.eje)
  if (!is.null(wrap_ejes) && is.finite(wrap_ejes) && wrap_ejes > 0) {
    if (requireNamespace("stringr", quietly = TRUE)) {
      lab_ejes <- stringr::str_wrap(lab_ejes, width = as.integer(wrap_ejes))
    }
  }

  # ---------------------------------------------------------------------------
  # 2) Geometría (x,y)
  # ---------------------------------------------------------------------------
  K <- length(levels(df_plot$.eje))
  theta0 <- -pi/2

  angle_tbl <- tibble::tibble(
    .eje = factor(levels(df_plot$.eje), levels = levels(df_plot$.eje)),
    .idx = seq_len(K),
    .ang = theta0 + 2*pi*(seq_len(K)-1)/K
  )

  df_xy <- df_plot |>
    dplyr::left_join(angle_tbl, by = ".eje") |>
    dplyr::mutate(
      x = .data$.valor * cos(.data$.ang),
      y = .data$.valor * sin(.data$.ang)
    )

  df_poly <- df_xy |>
    dplyr::arrange(.data$.grupo, .data$.idx) |>
    dplyr::group_by(.data$.grupo) |>
    dplyr::group_modify(function(g, ...) dplyr::bind_rows(g, g[1, , drop = FALSE])) |>
    dplyr::ungroup()

  # ---------------------------------------------------------------------------
  # 3) Límites radiales
  # ---------------------------------------------------------------------------
  if (is.null(limites)) {
    r_lim <- c(0, 1)
  } else {
    r_lim <- suppressWarnings(as.numeric(limites))
    if (length(r_lim) != 2 || any(!is.finite(r_lim))) r_lim <- c(0, 1)
    r_lim <- sort(r_lim)
    r_lim[1] <- max(0, r_lim[1])
    r_lim[2] <- min(1, r_lim[2])
    if (r_lim[2] <= r_lim[1]) r_lim <- c(0, 1)
  }

  rings <- unique(seq(r_lim[1], r_lim[2], length.out = cortes_grilla))

  grid_df <- NULL
  if (isTRUE(mostrar_tela)) {
    grid_df <- lapply(rings, function(rr) {
      lvl <- angle_tbl |>
        dplyr::mutate(.r = rr, x = rr * cos(.data$.ang), y = rr * sin(.data$.ang)) |>
        dplyr::arrange(.data$.idx)
      dplyr::bind_rows(lvl, lvl[1, , drop = FALSE])
    }) |> dplyr::bind_rows()
  }

  axes_df <- NULL
  if (isTRUE(mostrar_radios)) {
    axes_df <- angle_tbl |>
      dplyr::mutate(x0 = 0, y0 = 0, x1 = r_lim[2] * cos(.data$.ang), y1 = r_lim[2] * sin(.data$.ang))
  }

  level_lab <- NULL
  if (isTRUE(mostrar_niveles)) level_lab <- tibble::tibble(.nivel = rings, x = rings, y = 0)

  label_ring <- r_lim[2] * eje_label_mult
  lab_axes <- angle_tbl |>
    dplyr::mutate(
      eje = lab_ejes[.data$.idx],
      x   = label_ring * cos(.data$.ang),
      y   = label_ring * sin(.data$.ang)
    )

  # ---------------------------------------------------------------------------
  # 4) Paleta
  # ---------------------------------------------------------------------------
  pal <- NULL
  if (!is.null(colores_series)) {
    cs <- as.character(colores_series)
    if (is.null(names(cs))) {
      cs <- cs[seq_len(min(length(cs), length(grupos)))]
      cs <- stats::setNames(cs, as.character(grupos)[seq_along(cs)])
    } else {
      names(cs) <- trimws(as.character(names(cs)))
    }
    g_chr <- as.character(grupos)
    pal <- cs[g_chr]
    if (all(is.na(pal)) || length(pal) == 0) pal <- NULL
  } else if (requireNamespace("scales", quietly = TRUE)) {
    pal <- stats::setNames(scales::hue_pal()(length(grupos)), as.character(grupos))
  }

  # ---------------------------------------------------------------------------
  # 5) Plot (base)
  # ---------------------------------------------------------------------------
  leg_pos <- if (!isTRUE(mostrar_leyenda)) "none" else if (leyenda_posicion == "derecha") "right" else "bottom"

  p <- ggplot2::ggplot() +
    ggplot2::theme_minimal(base_size = 9) +
    ggplot2::theme(
      panel.grid       = ggplot2::element_blank(),
      axis.title       = ggplot2::element_blank(),
      axis.text        = ggplot2::element_blank(),
      axis.ticks       = ggplot2::element_blank(),
      plot.margin      = ggplot2::margin(0,0,0,0),
      panel.spacing    = grid::unit(0, "pt"),
      legend.position  = leg_pos,
      legend.title     = ggplot2::element_blank(),
      legend.text      = ggplot2::element_text(
        color  = color_leyenda,
        size   = size_leyenda,
        margin = ggplot2::margin(l = legend_espaciado/2, r = legend_espaciado/2, unit = "pt")
      ),
      legend.key.width      = grid::unit(legend_key_cm, "cm"),
      legend.key.height     = grid::unit(legend_key_cm, "cm"),
      legend.key.spacing.x  = grid::unit(legend_key_spacing_x_cm, "cm"),
      plot.title = ggplot2::element_text(
        color = color_titulo, size = size_titulo,
        face  = if ("titulo" %in% textos_negrita) "bold" else "plain",
        hjust = hjust_titulo
      ),
      plot.subtitle = ggplot2::element_text(
        color = color_subtitulo, size = size_subtitulo,
        face  = if ("subtitulo" %in% textos_negrita) "bold" else "plain",
        hjust = hjust_titulo
      ),
      plot.caption = ggplot2::element_text(
        color = color_nota_pie, size = size_nota_pie,
        face  = if ("nota_pie" %in% textos_negrita) "bold" else "plain",
        hjust = hjust_caption
      ),
      plot.background  = ggplot2::element_rect(fill = color_fondo, color = NA),
      panel.background = ggplot2::element_rect(fill = color_fondo, color = NA)
    ) +
    ggplot2::labs(title = titulo, subtitle = subtitulo, caption = nota_pie)

  # ---------------------------------------------------------------------------
  # Capas “normales” (para rplot/png/word).
  # ---------------------------------------------------------------------------
  if (isTRUE(mostrar_tela) && !is.null(grid_df)) {

    grid_df2 <- grid_df |>
      dplyr::filter(is.finite(.data$x), is.finite(.data$y), !is.na(.data$x), !is.na(.data$y))

    if (ppt_safe) {
      # Importante: Es PPT SAFE, NO polygon (evita C_polygon segfault)
      p <- p + ggplot2::geom_path(
        data = grid_df2,
        ggplot2::aes(x = .data$x, y = .data$y, group = .data$.r),
        color = color_grilla, linewidth = 0.5
      )
    } else {
      p <- p + ggplot2::geom_polygon(
        data = grid_df2,
        ggplot2::aes(x = .data$x, y = .data$y, group = .data$.r),
        fill = NA, color = color_grilla, linewidth = 0.5
      )
    }
  }
  if (isTRUE(mostrar_radios) && !is.null(axes_df)) {
    p <- p + ggplot2::geom_segment(
      data = axes_df,
      ggplot2::aes(x = .data$x0, y = .data$y0, xend = .data$x1, yend = .data$y1),
      color = color_radios, linewidth = 0.5
    )
  }
  if (isTRUE(rellenar_poligono)) {
    p <- p + ggplot2::geom_polygon(
      data = df_poly,
      ggplot2::aes(x = .data$x, y = .data$y, group = .data$.grupo, fill = .data$.grupo),
      color = NA, alpha = alpha_relleno
    )
  }

  p <- p + ggplot2::geom_path(
    data = df_poly,
    ggplot2::aes(x = .data$x, y = .data$y, group = .data$.grupo, color = .data$.grupo),
    linewidth = size_linea
  )

  if (isTRUE(mostrar_puntos)) {
    p <- p + ggplot2::geom_point(
      data = df_xy,
      ggplot2::aes(x = .data$x, y = .data$y, color = .data$.grupo),
      size = size_punto
    )
  }

  p <- p + ggplot2::geom_text(
    data = lab_axes,
    ggplot2::aes(x = .data$x, y = .data$y, label = .data$eje),
    size = size_ejes / 3,
    colour = color_ejes,
    fontface = if ("ejes" %in% textos_negrita) "bold" else "plain",
    lineheight = 0.95
  )

  if (isTRUE(mostrar_niveles) && !is.null(level_lab)) {
    p <- p + ggplot2::geom_text(
      data = level_lab,
      ggplot2::aes(x = .data$x, y = .data$y, label = paste0(round(.data$.nivel * 100), "%")),
      size = 3,
      color = "grey40",
      fontface = if ("niveles" %in% textos_negrita) "bold" else "plain",
      vjust = -0.2
    )
  }

  lim_xy <- r_lim[2] * max(1.28, eje_label_mult * 1.10)

  clip_mode <- if (ppt_safe) "on" else "off"

  p <- p +
    ggplot2::coord_equal(clip = clip_mode) +
    ggplot2::scale_x_continuous(limits = c(-lim_xy, lim_xy), expand = ggplot2::expansion(mult = 0, add = 0)) +
    ggplot2::scale_y_continuous(limits = c(-lim_xy, lim_xy), expand = ggplot2::expansion(mult = 0, add = 0))

  if (!is.null(pal)) {
    p <- p + ggplot2::scale_color_manual(values = pal, breaks = as.character(grupos), drop = FALSE)
    if (isTRUE(rellenar_poligono)) p <- p + ggplot2::scale_fill_manual(values = pal, breaks = as.character(grupos), drop = FALSE)
  } else {
    p <- p + ggplot2::scale_color_discrete(drop = FALSE)
    if (isTRUE(rellenar_poligono)) p <- p + ggplot2::scale_fill_discrete(drop = FALSE)
  }

  p <- p + ggplot2::guides(
    color = ggplot2::guide_legend(
      ncol  = if (leyenda_posicion == "abajo") legend_n_por_fila else 1,
      byrow = TRUE,
      keywidth  = grid::unit(legend_key_cm, "cm"),
      keyheight = grid::unit(legend_key_cm, "cm")
    ),
    fill  = if (isTRUE(rellenar_poligono)) ggplot2::guide_legend(
      ncol  = if (leyenda_posicion == "abajo") legend_n_por_fila else 1,
      byrow = TRUE,
      keywidth  = grid::unit(legend_key_cm, "cm"),
      keyheight = grid::unit(legend_key_cm, "cm")
    ) else "none"
  )

  # ---------------------------------------------------------------------------
  # CANVAS (radar + tabla opcional)
  # ---------------------------------------------------------------------------
  if (isTRUE(usar_canvas)) {
    if (!requireNamespace("cowplot", quietly = TRUE)) stop("Para `usar_canvas=TRUE` se requiere cowplot.", call. = FALSE)

    has_header  <- (!is.null(titulo) && nzchar(titulo)) || (!is.null(subtitulo) && nzchar(subtitulo))
    has_caption <- (!is.null(nota_pie) && nzchar(nota_pie))
    has_legend  <- isTRUE(mostrar_leyenda) && leg_pos != "none" && length(grupos) > 0

    p_panel <- p +
      ggplot2::labs(title = NULL, subtitle = NULL, caption = NULL) +
      ggplot2::theme(legend.position = "none", plot.margin = ggplot2::margin(0,0,0,0))

    leg_grob <- NULL
    if (has_legend) {
      p_for_legend <- p +
        ggplot2::theme(
          legend.position  = "bottom",
          legend.direction = "horizontal",
          legend.box       = "horizontal",
          legend.title     = ggplot2::element_blank(),
          legend.text = ggplot2::element_text(
            color  = color_leyenda,
            size   = size_leyenda,
            face   = if ("leyenda" %in% textos_negrita) "bold" else "plain",
            margin = ggplot2::margin(l = legend_espaciado/2, r = legend_espaciado/2, unit = "pt")
          ),
          legend.key.width     = grid::unit(legend_key_cm, "cm"),
          legend.key.height    = grid::unit(legend_key_cm, "cm"),
          legend.key.spacing.x = grid::unit(legend_key_spacing_x_cm, "cm"),
          plot.margin = ggplot2::margin(0,0,0,0)
        ) +
        ggplot2::guides(
          color = ggplot2::guide_legend(byrow = TRUE, ncol = legend_n_por_fila,
                                        keywidth  = grid::unit(legend_key_cm, "cm"),
                                        keyheight = grid::unit(legend_key_cm, "cm")),
          fill  = if (isTRUE(rellenar_poligono)) ggplot2::guide_legend(byrow = TRUE, ncol = legend_n_por_fila,
                                                                       keywidth  = grid::unit(legend_key_cm, "cm"),
                                                                       keyheight = grid::unit(legend_key_cm, "cm")) else "none"
        )
      leg_grob <- cowplot::get_legend(p_for_legend)
    }

    h_panel_in <- if (!is.null(canvas_h_panel_in) && is.finite(canvas_h_panel_in) && canvas_h_panel_in > 0) {
      canvas_h_panel_in
    } else {
      max(1, K) * alto_por_eje
    }

    h_header_in  <- if (has_header)  canvas_h_header_in  else 0
    h_legend_in  <- if (has_legend)  canvas_h_legend_in  else 0
    h_caption_in <- if (has_caption) canvas_h_caption_in else 0

    h_total_in <- h_header_in + h_panel_in + h_legend_in + h_caption_in
    if (h_total_in <= 0) h_total_in <- 1

    header_h  <- h_header_in  / h_total_in
    panel_h   <- h_panel_in   / h_total_in
    legend_h  <- h_legend_in  / h_total_in
    caption_h <- h_caption_in / h_total_in

    y_header0  <- 1 - header_h
    y_panel0   <- y_header0 - panel_h
    y_legend0  <- y_panel0  - legend_h
    y_caption0 <- y_legend0 - caption_h

    .ph_border <- function(x, y, w, h) {
      cowplot::draw_grob(
        grid::rectGrob(
          x = 0, y = 0, width = 1, height = 1,
          just = c("left","bottom"),
          gp = grid::gpar(col = debug_ph_col, fill = NA, lwd = debug_ph_lwd)
        ),
        x = x, y = y, width = w, height = h,
        hjust = 0, vjust = 0
      )
    }

    canvas <- cowplot::ggdraw()

    # Header
    if (has_header) {
      y_header_center <- y_header0 + header_h * 0.5
      dy_head <- encabezado_desplazamiento_in / h_total_in
      sep     <- encabezado_separacion_in     / h_total_in

      has_t <- (!is.null(titulo) && nzchar(titulo))
      has_s <- (!is.null(subtitulo) && nzchar(subtitulo))

      if (has_t && has_s) {
        y_title <- y_header_center + (sep * 0.5) + dy_head
        y_sub   <- y_header_center - (sep * 0.5) + dy_head
      } else if (has_t) {
        y_title <- y_header_center + dy_head
        y_sub   <- NA_real_
      } else {
        y_title <- NA_real_
        y_sub   <- y_header_center + dy_head
      }

      if (has_t) {
        canvas <- canvas + cowplot::draw_text(
          titulo, x = hjust_titulo, y = y_title,
          hjust = hjust_titulo, vjust = 0.5,
          size = size_titulo, colour = color_titulo,
          fontface = if ("titulo" %in% textos_negrita) "bold" else "plain"
        )
      }
      if (has_s) {
        canvas <- canvas + cowplot::draw_text(
          subtitulo,
          x = hjust_titulo, y = y_sub,
          hjust = hjust_titulo, vjust = 0.5,
          size = size_subtitulo, colour = color_subtitulo,
          fontface = if ("subtitulo" %in% textos_negrita) "bold" else "plain"
        )
      }
      if (debug_ph_bordes) canvas <- canvas + .ph_border(0, y_header0, 1, header_h)
    }

    # Panel: radar + tabla
    if (isTRUE(mostrar_tabla_derecha)) {
      tabla_ph_ancho <- suppressWarnings(as.numeric(tabla_ph_ancho))
      if (!is.finite(tabla_ph_ancho) || tabla_ph_ancho <= 0 || tabla_ph_ancho >= 0.85) tabla_ph_ancho <- 0.40
      tabla_ph_gap <- suppressWarnings(as.numeric(tabla_ph_gap))
      if (!is.finite(tabla_ph_gap) || tabla_ph_gap < 0) tabla_ph_gap <- 0.03
      tabla_ph_margin_top <- suppressWarnings(as.numeric(tabla_ph_margin_top))
      if (!is.finite(tabla_ph_margin_top) || tabla_ph_margin_top < 0) tabla_ph_margin_top <- 0.04
      tabla_ph_margin_bot <- suppressWarnings(as.numeric(tabla_ph_margin_bot))
      if (!is.finite(tabla_ph_margin_bot) || tabla_ph_margin_bot < 0) tabla_ph_margin_bot <- 0.06

      w_tab <- tabla_ph_ancho
      w_gap <- tabla_ph_gap
      w_radar <- 1 - w_tab - w_gap
      if (w_radar <= 0.10) {
        w_tab <- min(0.45, max(0.25, w_tab))
        w_gap <- min(0.05, max(0.01, w_gap))
        w_radar <- 1 - w_tab - w_gap
      }

      # Radar izquierda
      canvas <- canvas + cowplot::draw_plot(p_panel, x = 0, y = y_panel0, width = w_radar, height = panel_h)
      if (debug_ph_bordes) canvas <- canvas + .ph_border(0, y_panel0, w_radar, panel_h)

      # Tabla derecha con top/bot
      y_tab <- y_panel0 + tabla_ph_margin_bot
      h_tab <- panel_h - tabla_ph_margin_top - tabla_ph_margin_bot
      if (h_tab <= 0) {
        y_tab <- y_panel0
        h_tab <- panel_h
      }

      tb <- .make_tabla_ttb_df(
        df_plot,
        ejes   = levels(df_plot$.eje),
        grupos = levels(df_plot$.grupo),
        digits = tabla_digits,
        titulo_left = titulo_tabla
      )

      # ------------------------------------------------------------
      # WRAP 1RA COLUMNA (ejes) según el ancho real del PH de la tabla
      # ------------------------------------------------------------
      if (requireNamespace("stringr", quietly = TRUE)) {

        # ancho real disponible del PH de la tabla (en pulgadas)
        ph_w_in <- ancho * w_tab

        # porcentaje del PH que se quiere para la 1ra columna (ajustable)
        firstcol_frac <- 0.62
        firstcol_in   <- ph_w_in * firstcol_frac

        # estimación: caracteres por pulgada según tamaño de fuente
        # (0.55 es un factor práctico para fuentes tipo Arial)
        chars_per_in <- 72 / (tabla_body_size * 0.55)

        wrap_n <- floor(firstcol_in * chars_per_in)
        wrap_n <- max(12, min(60, wrap_n))  # clamps razonables

        tb[[1]] <- stringr::str_wrap(tb[[1]], width = wrap_n)
      }

      tab_grob <- .make_table_grob_ttb_style(
        tb,
        header_fill = tabla_header_fill,
        body_fill   = tabla_body_fill,
        grid_col    = tabla_grid_col,
        text_blue   = tabla_text_blue,
        font_family = tabla_font_family,
        header_size = tabla_header_size,
        body_size   = tabla_body_size,
        firstcol_bold = tabla_firstcol_bold,
        highlight_threshold = umbral_rojo_pct,
        highlight_col = "red",
        padding_mm = tabla_padding_mm
      )

      tab_draw <- if (isTRUE(tabla_clip)) .wrap_clip(tab_grob) else tab_grob

      # -----------------------------------------------------------------
      # AUTO-FIT (robusto): medir el grob en pulgadas y escalar contra el PH
      # -----------------------------------------------------------------
      scale_tab <- 1

      if (isTRUE(tabla_auto_fit)) {

        # OJO: esto es más estable que grobWidth() en algunos devices
        gw_in <- suppressWarnings(grid::convertWidth(sum(tab_grob$widths),  "in", valueOnly = TRUE))
        gh_in <- suppressWarnings(grid::convertHeight(sum(tab_grob$heights), "in", valueOnly = TRUE))

        # Tamaño disponible del PH (en pulgadas) usando el tamaño final del canvas
        ph_w_in <- ancho * w_tab
        ph_h_in <- alto  * h_tab

        if (is.finite(gw_in) && gw_in > 0 && is.finite(gh_in) && gh_in > 0) {

          s_w <- ph_w_in / gw_in
          s_h <- ph_h_in / gh_in

          scale_tab <- min(s_w, s_h)

          if (!isTRUE(tabla_allow_upscale)) scale_tab <- min(1, scale_tab)

          scale_tab <- scale_tab * tabla_fit_pad
          if (!is.finite(scale_tab) || scale_tab <= 0) scale_tab <- 1
        }
      }

      # IMPORTANTE: centrar el grob dentro del PH:
      # draw_grob con hjust/vjust = 0.5 y x/y al centro del PH
      canvas <- canvas + cowplot::draw_grob(
        tab_draw,
        x = (w_radar + w_gap) + (w_tab * 0.5),
        y = y_tab + (h_tab * 0.5),
        width  = w_tab,
        height = h_tab,
        hjust = 0.5, vjust = 0.5,
        scale = scale_tab
      )

      if (debug_ph_bordes) canvas <- canvas + .ph_border(w_radar + w_gap, y_tab, w_tab, h_tab)

    } else {
      canvas <- canvas + cowplot::draw_plot(p_panel, x = 0, y = y_panel0, width = 1, height = panel_h)
      if (debug_ph_bordes) canvas <- canvas + .ph_border(0, y_panel0, 1, panel_h)
    }

  # ---------------------------------------------------------------
  # LEYENDA CENTRADA SOLO EN EL PH DEL PANEL
  # ---------------------------------------------------------------
  if (has_legend && !is.null(leg_grob)) {

    # ancho del panel (izquierda)
    panel_w <- if (isTRUE(mostrar_tabla_derecha)) w_radar else 1

    # leyenda solo ocupa ancho del panel
    legend_ph_x <- 0
    legend_ph_w <- panel_w

    y_legend_center <- y_legend0 + legend_h * 0.5
    dy_leg <- leyenda_desplazamiento_in / h_total_in

    leg_w_npc <- suppressWarnings(
      grid::convertWidth(sum(leg_grob$widths), "npc", valueOnly = TRUE)
    )
    if (!is.finite(leg_w_npc) || leg_w_npc <= 0) leg_w_npc <- 1

    canvas <- canvas + cowplot::draw_grob(
      leg_grob,
      x = legend_ph_x + (legend_ph_w * 0.5),
      y = y_legend_center + dy_leg,
      width  = legend_ph_w,
      height = legend_h,
      hjust = 0.5,
      vjust = 0.5
    )

    if (debug_ph_bordes) {
      canvas <- canvas + .ph_border(legend_ph_x, y_legend0, legend_ph_w, legend_h)
    }
  }

    # Caption
    if (has_caption) {
      canvas <- canvas + cowplot::draw_text(
        nota_pie,
        x = hjust_caption,
        y = y_caption0 + caption_h * 0.35,
        hjust = hjust_caption,
        vjust = 0.5,
        size = size_nota_pie,
        colour = color_nota_pie,
        fontface = if ("nota_pie" %in% textos_negrita) "bold" else "plain"
      )
      if (debug_ph_bordes) canvas <- canvas + .ph_border(0, y_caption0, 1, caption_h)
    }

    # -------------------------------------------------------------------------
    # EXPORT desde CANVAS
    # -------------------------------------------------------------------------
    if (exportar == "rplot") return(canvas)

    if (is.null(path_salida) || !nzchar(path_salida)) stop("`path_salida` es requerido para exportar.", call. = FALSE)

    if (exportar == "png") {
      ggplot2::ggsave(path_salida, canvas, width = ancho, height = alto, units = "in", dpi = dpi, bg = "transparent")
      return(invisible(canvas))
    }

    # ============ PPT/WORD =============
    if (exportar %in% c("ppt","word")) {
      if (!requireNamespace("officer", quietly = TRUE)) stop("Para exportar a PPT/Word se requiere officer.", call. = FALSE)
      if (!requireNamespace("rvg", quietly = TRUE))     stop("Para exportar a PPT/Word se requiere rvg.", call. = FALSE)

      # ---- PPT SAFE OBJ (para rvg): NO polygons + clip on ----
      # Nota: exportamos el CANVAS (cowplot) tal cual; la estabilidad viene de:
      # - El radar base ya está sin “fill” si exportar=="ppt" (ver bloque abajo),
      # - Y la tabla es un grob (grid) sin polygons problemáticos.
      #
      # Aun así, si hay aborts, se recomienda exportar el radar a rvg y la tabla con officer como tabla nativa.
      if (exportar == "ppt") {

        # Debug por steps (Rscript) para aislar segfaults (si se activa)
        .run_ppt_step <- function(step = c("01_read", "02_slide", "03_size", "04_ph_with", "05_print"),
                                  plot_obj,
                                  path_out,
                                  ppt_layout = "Blank",
                                  ppt_master = "Office Theme") {
          step <- match.arg(step)

          f_plot <- tempfile(fileext = ".rds")
          f_err  <- tempfile(fileext = ".txt")
          f_scr  <- tempfile(fileext = ".R")

          saveRDS(plot_obj, f_plot)

          code <- c(
            "suppressPackageStartupMessages({library(officer); library(rvg); library(ggplot2); library(cowplot); library(grid)})",
            sprintf("p <- readRDS('%s')", gsub("\\\\", "/", f_plot)),
            sprintf("out <- '%s'", gsub("\\\\", "/", path_out)),
            "doc <- read_pptx()",
            if (step %in% c("02_slide","03_size","04_ph_with","05_print"))
              sprintf("doc <- add_slide(doc, layout = '%s', master = '%s')", ppt_layout, ppt_master)
            else "invisible(NULL)",
            if (step %in% c("03_size","04_ph_with","05_print"))
              "ss <- slide_size(doc); sw <- ss$width; sh <- ss$height"
            else "invisible(NULL)",
            if (step %in% c("04_ph_with","05_print"))
              "doc <- ph_with(doc, value = rvg::dml(ggobj = p), location = ph_location(left=0, top=0, width=sw, height=sh))"
            else "invisible(NULL)",
            if (step %in% c("05_print"))
              "print(doc, target = out)"
            else "invisible(NULL)",
            "cat('OK\\n')"
          )

          writeLines(code, f_scr)

          rscript <- Sys.which("Rscript")
          if (!nzchar(rscript)) stop("No se encontró Rscript en PATH.", call. = FALSE)

          suppressWarnings(system2(rscript, args = c(shQuote(f_scr)), stdout = FALSE, stderr = f_err))

          err <- if (file.exists(f_err)) paste(readLines(f_err, warn = FALSE), collapse = "\n") else ""
          list(stderr = err, out_exists = file.exists(path_out))
        }

        if (isTRUE(debug_ppt)) {
          cat("PPT EXPORT DEBUG START\n", file = debug_ppt_log)
          .log <- function(...) cat(..., "\n", file = debug_ppt_log, append = TRUE)

          steps <- c("01_read", "02_slide", "03_size", "04_ph_with", "05_print")
          last_ok <- NA_character_

          for (st in steps) {
            out_step <- tempfile(fileext = paste0("_", st, ".pptx"))
            res <- .run_ppt_step(
              step       = st,
              plot_obj   = canvas,
              path_out   = out_step,
              ppt_layout = ppt_layout,
              ppt_master = ppt_master
            )

            .log("[TRY] ", st)
            if (nzchar(res$stderr)) {
              .log("[STDERR] ")
              .log(res$stderr)
              .log("[WARN?] ", st, " (ver STDERR arriba)")
            } else {
              .log("[OK] ", st)
              last_ok <- st
            }

            if (st == "05_print") {
              if (isTRUE(res$out_exists)) {
                .log("[OK] 05_print (pptx creado)")
              } else {
                .log("[FAIL] 05_print")
                .log("STOP: aborta en print() o antes (no se creó pptx). Último OK: ", last_ok %||% "ninguno")
              }
            }
          }
          .log("PPT EXPORT DEBUG END")
          message("PPT export debug log -> ", normalizePath(debug_ppt_log, winslash = "/"))
        }

        doc <- if (ppt_append && file.exists(path_salida)) officer::read_pptx(path_salida) else officer::read_pptx()
        doc <- officer::add_slide(doc, layout = ppt_layout, master = ppt_master)
        doc <- officer::ph_with(doc, value = rvg::dml(ggobj = canvas), location = officer::ph_location_fullsize())
        print(doc, target = path_salida)
        return(invisible(canvas))
      }

      if (exportar == "word") {
        doc <- if (file.exists(path_salida)) officer::read_docx(path_salida) else officer::read_docx()
        doc <- officer::body_add_par(doc, value = "", style = "Normal")
        doc <- officer::body_add_dml(doc, value = rvg::dml(ggobj = canvas), width = ancho, height = alto)
        print(doc, target = path_salida)
        return(invisible(canvas))
      }
    }

    stop("Tipo de exportación no soportado.", call. = FALSE)
  }

  # ---------------------------------------------------------------------------
  # NO CANVAS
  # ---------------------------------------------------------------------------
  if (exportar == "rplot") return(p)

  if (is.null(path_salida) || !nzchar(path_salida)) stop("`path_salida` es requerido para exportar.", call. = FALSE)

  if (exportar == "png") {
    ggplot2::ggsave(path_salida, p, width = ancho, height = alto, units = "in", dpi = dpi, bg = "transparent")
    return(invisible(p))
  }

  if (exportar %in% c("ppt","word")) {
    if (!requireNamespace("officer", quietly = TRUE)) stop("Para exportar a PPT/Word se requiere officer.", call. = FALSE)
    if (!requireNamespace("rvg", quietly = TRUE))     stop("Para exportar a PPT/Word se requiere rvg.", call. = FALSE)

    if (exportar == "ppt") {
      # ---- PLOT PPT SAFE (reconstrucción SIN polygons) ----
      # 1) Forzar NO fill en PPT (aunque el usuario lo pida)
      rellenar_ppt <- FALSE

      # 2) Malla sin geom_polygon: se reemplaza por geom_path
      #    y se filtran coords no finitas por seguridad
      grid_df_ppt <- grid_df
      if (!is.null(grid_df_ppt)) {
        grid_df_ppt <- grid_df_ppt |>
          dplyr::filter(is.finite(.data$x), is.finite(.data$y), !is.na(.data$x), !is.na(.data$y))
      }
      axes_df_ppt <- axes_df
      if (!is.null(axes_df_ppt)) {
        axes_df_ppt <- axes_df_ppt |>
          dplyr::filter(is.finite(.data$x0), is.finite(.data$y0), is.finite(.data$x1), is.finite(.data$y1))
      }
      df_poly_ppt <- df_poly |>
        dplyr::filter(is.finite(.data$x), is.finite(.data$y), !is.na(.data$x), !is.na(.data$y))
      df_xy_ppt <- df_xy |>
        dplyr::filter(is.finite(.data$x), is.finite(.data$y), !is.na(.data$x), !is.na(.data$y))

      # 3) límites: incluir labels dentro del viewport
      lim_xy_ppt <- max(r_lim[2] * 1.25, (r_lim[2] * eje_label_mult) * 1.12)

      fondo_ppt <- if (is.na(color_fondo) || is.null(color_fondo)) "transparent" else color_fondo

      p_ppt <- ggplot2::ggplot() +
        ggplot2::theme_minimal(base_size = 9) +
        ggplot2::theme(
          panel.grid       = ggplot2::element_blank(),
          axis.title       = ggplot2::element_blank(),
          axis.text        = ggplot2::element_blank(),
          axis.ticks       = ggplot2::element_blank(),
          plot.margin      = ggplot2::margin(0,0,0,0),
          panel.spacing    = grid::unit(0, "pt"),
          legend.position  = leg_pos,
          legend.title     = ggplot2::element_blank(),
          legend.text      = ggplot2::element_text(
            color  = color_leyenda,
            size   = size_leyenda,
            family = "sans",
            margin = ggplot2::margin(l = legend_espaciado/2, r = legend_espaciado/2, unit = "pt")
          ),
          legend.key.width      = grid::unit(legend_key_cm, "cm"),
          legend.key.height     = grid::unit(legend_key_cm, "cm"),
          legend.key.spacing.x  = grid::unit(legend_key_spacing_x_cm, "cm"),
          plot.title = ggplot2::element_text(
            color = color_titulo, size = size_titulo, family = "sans",
            face  = if ("titulo" %in% textos_negrita) "bold" else "plain",
            hjust = hjust_titulo
          ),
          plot.subtitle = ggplot2::element_text(
            color = color_subtitulo, size = size_subtitulo, family = "sans",
            face  = if ("subtitulo" %in% textos_negrita) "bold" else "plain",
            hjust = hjust_titulo
          ),
          plot.caption = ggplot2::element_text(
            color = color_nota_pie, size = size_nota_pie, family = "sans",
            face  = if ("nota_pie" %in% textos_negrita) "bold" else "plain",
            hjust = hjust_caption
          ),
          plot.background  = ggplot2::element_rect(fill = fondo_ppt, color = NA),
          panel.background = ggplot2::element_rect(fill = fondo_ppt, color = NA)
        ) +
        ggplot2::labs(title = titulo, subtitle = subtitulo, caption = nota_pie)

      if (isTRUE(mostrar_tela) && !is.null(grid_df_ppt)) {
        p_ppt <- p_ppt + ggplot2::geom_path(
          data = grid_df_ppt,
          ggplot2::aes(x = .data$x, y = .data$y, group = .data$.r),
          color = color_grilla, linewidth = 0.5
        )
      }

      if (isTRUE(mostrar_radios) && !is.null(axes_df_ppt)) {
        p_ppt <- p_ppt + ggplot2::geom_segment(
          data = axes_df_ppt,
          ggplot2::aes(x = .data$x0, y = .data$y0, xend = .data$x1, yend = .data$y1),
          color = color_radios, linewidth = 0.5
        )
      }

      # NO geom_polygon() en PPT
      if (isTRUE(rellenar_ppt) && FALSE) {
        p_ppt <- p_ppt + ggplot2::geom_polygon(
          data = df_poly_ppt,
          ggplot2::aes(x = .data$x, y = .data$y, group = .data$.grupo, fill = .data$.grupo),
          color = NA, alpha = alpha_relleno
        )
      }

      p_ppt <- p_ppt + ggplot2::geom_path(
        data = df_poly_ppt,
        ggplot2::aes(x = .data$x, y = .data$y, group = .data$.grupo, color = .data$.grupo),
        linewidth = size_linea
      )

      if (isTRUE(mostrar_puntos)) {
        p_ppt <- p_ppt + ggplot2::geom_point(
          data = df_xy_ppt,
          ggplot2::aes(x = .data$x, y = .data$y, color = .data$.grupo),
          size = size_punto
        )
      }

      p_ppt <- p_ppt + ggplot2::geom_text(
        data = lab_axes,
        ggplot2::aes(x = .data$x, y = .data$y, label = .data$eje),
        size = size_ejes / 3,
        colour = color_ejes,
        family = "sans",
        fontface = if ("ejes" %in% textos_negrita) "bold" else "plain",
        lineheight = 1
      )

      if (isTRUE(mostrar_niveles) && !is.null(level_lab)) {
        p_ppt <- p_ppt + ggplot2::geom_text(
          data = level_lab,
          ggplot2::aes(x = .data$x, y = .data$y, label = paste0(round(.data$.nivel * 100), "%")),
          size = 3,
          color = "grey40",
          family = "sans",
          fontface = if ("niveles" %in% textos_negrita) "bold" else "plain",
          vjust = -0.2
        )
      }

      p_ppt <- p_ppt +
        ggplot2::coord_equal(clip = "on") +
        ggplot2::scale_x_continuous(limits = c(-lim_xy_ppt, lim_xy_ppt), expand = ggplot2::expansion(mult = 0, add = 0)) +
        ggplot2::scale_y_continuous(limits = c(-lim_xy_ppt, lim_xy_ppt), expand = ggplot2::expansion(mult = 0, add = 0))

      if (!is.null(pal)) {
        p_ppt <- p_ppt + ggplot2::scale_color_manual(values = pal, breaks = as.character(grupos), drop = FALSE)
      } else {
        p_ppt <- p_ppt + ggplot2::scale_color_discrete(drop = FALSE)
      }

      p_ppt <- p_ppt + ggplot2::guides(
        color = ggplot2::guide_legend(
          ncol  = if (leyenda_posicion == "abajo") legend_n_por_fila else 1,
          byrow = TRUE,
          keywidth  = grid::unit(legend_key_cm, "cm"),
          keyheight = grid::unit(legend_key_cm, "cm")
        )
      )

      # --- Debug steps (opcional) para p_ppt ---
      if (isTRUE(debug_ppt)) {
        cat("PPT EXPORT DEBUG START\n", file = debug_ppt_log)
        .log <- function(...) cat(..., "\n", file = debug_ppt_log, append = TRUE)

        .run_ppt_step <- function(step = c("01_read", "02_slide", "03_size", "04_ph_with", "05_print"),
                                  plot_obj,
                                  path_out,
                                  ppt_layout = "Blank",
                                  ppt_master = "Office Theme") {
          step <- match.arg(step)

          f_plot <- tempfile(fileext = ".rds")
          f_err  <- tempfile(fileext = ".txt")
          f_scr  <- tempfile(fileext = ".R")

          saveRDS(plot_obj, f_plot)

          code <- c(
            "suppressPackageStartupMessages({library(officer); library(rvg); library(ggplot2); library(grid)})",
            sprintf("p <- readRDS('%s')", gsub("\\\\", "/", f_plot)),
            sprintf("out <- '%s'", gsub("\\\\", "/", path_out)),
            "doc <- read_pptx()",
            if (step %in% c("02_slide","03_size","04_ph_with","05_print"))
              sprintf("doc <- add_slide(doc, layout = '%s', master = '%s')", ppt_layout, ppt_master)
            else "invisible(NULL)",
            if (step %in% c("03_size","04_ph_with","05_print"))
              "ss <- slide_size(doc); sw <- ss$width; sh <- ss$height"
            else "invisible(NULL)",
            if (step %in% c("04_ph_with","05_print"))
              "doc <- ph_with(doc, value = rvg::dml(ggobj = p), location = ph_location(left=0, top=0, width=sw, height=sh))"
            else "invisible(NULL)",
            if (step %in% c("05_print"))
              "print(doc, target = out)"
            else "invisible(NULL)",
            "cat('OK\\n')"
          )

          writeLines(code, f_scr)

          rscript <- Sys.which("Rscript")
          if (!nzchar(rscript)) stop("No se encontró Rscript en PATH.", call. = FALSE)

          suppressWarnings(system2(rscript, args = c(shQuote(f_scr)), stdout = FALSE, stderr = f_err))

          err <- if (file.exists(f_err)) paste(readLines(f_err, warn = FALSE), collapse = "\n") else ""
          list(stderr = err, out_exists = file.exists(path_out))
        }

        steps <- c("01_read", "02_slide", "03_size", "04_ph_with", "05_print")
        last_ok <- NA_character_

        for (st in steps) {
          out_step <- tempfile(fileext = paste0("_", st, ".pptx"))
          res <- .run_ppt_step(
            step       = st,
            plot_obj   = p_ppt,
            path_out   = out_step,
            ppt_layout = ppt_layout,
            ppt_master = ppt_master
          )

          .log("[TRY] ", st)
          if (nzchar(res$stderr)) {
            .log("[STDERR] ")
            .log(res$stderr)
            .log("[WARN?] ", st, " (ver STDERR arriba)")
          } else {
            .log("[OK] ", st)
            last_ok <- st
          }

          if (st == "05_print") {
            if (isTRUE(res$out_exists)) {
              .log("[OK] 05_print (pptx creado)")
            } else {
              .log("[FAIL] 05_print")
              .log("STOP: aborta en print() o antes (no se creó pptx). Último OK: ", last_ok %||% "ninguno")
            }
          }
        }
        .log("PPT EXPORT DEBUG END")
        message("PPT export debug log -> ", normalizePath(debug_ppt_log, winslash = "/"))
      }

      doc <- if (ppt_append && file.exists(path_salida)) officer::read_pptx(path_salida) else officer::read_pptx()
      doc <- officer::add_slide(doc, layout = ppt_layout, master = ppt_master)

      ss <- officer::slide_size(doc)
      doc <- officer::ph_with(
        doc,
        value    = rvg::dml(ggobj = p_ppt),
        location = officer::ph_location(left = 0, top = 0, width = ss$width, height = ss$height)
      )

      print(doc, target = path_salida)
      return(invisible(p_ppt))
    }

    if (exportar == "word") {
      doc <- if (file.exists(path_salida)) officer::read_docx(path_salida) else officer::read_docx()
      doc <- officer::body_add_par(doc, value = "", style = "Normal")
      doc <- officer::body_add_dml(doc, value = rvg::dml(ggobj = p), width = ancho, height = alto)
      print(doc, target = path_salida)
      return(invisible(p))
    }
  }

  stop("Tipo de exportación no soportado.", call. = FALSE)
}
