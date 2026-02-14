# =============================================================================
# graficar_radar() — plot-ready estilo prosecnur (canvas + export)
# PARCHES:
# 1) eje_label_mult: aleja texto de ejes (para que no tape el polígono)
# 2) coord_equal SIN xlim/ylim fijos + clip="off" + expand: evita cortes “raros”
# 3) colores_series: respeta paleta nombrada por etiqueta final (y maneja factores)
# 4) rellenar_poligono: si FALSE, NO mapea fill (evita colores por defecto)
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

    # ✅ nuevo: no rellenar (solo bordes)
    rellenar_poligono = FALSE,

    etiquetas_series = NULL,  # named: old -> new
    colores_series   = NULL,  # named por etiqueta final

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
    ppt_master = "Office Theme"
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y
  hjust_from_pos <- function(x) switch(x, "izquierda"=0, "centro"=0.5, "derecha"=1, 0.5)

  textos_negrita <- textos_negrita %||% character(0)

  # deps
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

  hjust_titulo  <- hjust_from_pos(pos_titulo)
  hjust_caption <- hjust_from_pos(pos_nota_pie)

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

  # relabel grupos
  if (!is.null(etiquetas_series) && length(etiquetas_series) > 0) {
    if (is.null(names(etiquetas_series))) stop("`etiquetas_series` debe ser nombrado: old -> new.", call. = FALSE)
    mp <- as.character(etiquetas_series)
    names(mp) <- as.character(names(etiquetas_series))
    df0$.grupo <- dplyr::recode(df0$.grupo, !!!mp)
  }

  # niveles de ejes y grupos (fijos)
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

  # wrap ejes
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

  rings <- seq(r_lim[1], r_lim[2], length.out = cortes_grilla)
  rings <- unique(rings)

  grid_df <- NULL
  if (isTRUE(mostrar_tela)) {
    grid_df <- lapply(rings, function(rr) {
      lvl <- angle_tbl |>
        dplyr::mutate(
          .r = rr,
          x  = rr * cos(.data$.ang),
          y  = rr * sin(.data$.ang)
        ) |>
        dplyr::arrange(.data$.idx)
      dplyr::bind_rows(lvl, lvl[1, , drop = FALSE])
    }) |> dplyr::bind_rows()
  }

  axes_df <- NULL
  if (isTRUE(mostrar_radios)) {
    axes_df <- angle_tbl |>
      dplyr::mutate(
        x0 = 0, y0 = 0,
        x1 = r_lim[2] * cos(.data$.ang),
        y1 = r_lim[2] * sin(.data$.ang)
      )
  }

  level_lab <- NULL
  if (isTRUE(mostrar_niveles)) {
    level_lab <- tibble::tibble(.nivel = rings, x = rings, y = 0)
  }

  # ✅ etiquetas ejes alejadas del centro
  label_ring <- r_lim[2] * eje_label_mult
  lab_axes <- angle_tbl |>
    dplyr::mutate(
      eje = lab_ejes[.data$.idx],
      x   = label_ring * cos(.data$.ang),
      y   = label_ring * sin(.data$.ang)
    )

  # ---------------------------------------------------------------------------
  # 4) Paleta — respeta names por etiqueta final
  # ---------------------------------------------------------------------------
  pal <- NULL
  if (!is.null(colores_series)) {
    cs <- as.character(colores_series)

    if (is.null(names(cs))) {
      # sin nombres: asignar por orden de grupos
      cs <- cs[seq_len(min(length(cs), length(grupos)))]
      cs <- stats::setNames(cs, as.character(grupos)[seq_along(cs)])
    } else {
      # con nombres: normalizar nombres y mapear contra niveles reales de grupo
      names(cs) <- trimws(as.character(names(cs)))
    }

    # grupos puede ser factor: convertir a character
    g_chr <- as.character(grupos)

    # map directo por nombre (etiqueta final)
    pal <- cs[g_chr]

    # si quedó NA (por mismatch), intentar también con niveles tal cual
    if (all(is.na(pal)) || length(pal) == 0) pal <- NULL
  } else {
    if (requireNamespace("scales", quietly = TRUE)) {
      pal <- stats::setNames(scales::hue_pal()(length(grupos)), as.character(grupos))
    }
  }

  # ---------------------------------------------------------------------------
  # 5) Plot
  # ---------------------------------------------------------------------------
  leg_pos <- if (!isTRUE(mostrar_leyenda)) "none" else if (leyenda_posicion == "derecha") "right" else "bottom"

  p <- ggplot2::ggplot() +
    ggplot2::theme_minimal(base_size = 9) +
    ggplot2::theme(
      panel.grid       = ggplot2::element_blank(),
      axis.title       = ggplot2::element_blank(),
      axis.text        = ggplot2::element_blank(),
      axis.ticks       = ggplot2::element_blank(),

      # ✅ clave para que NO corte texto “fuera” del panel
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
        color = color_titulo,
        size  = size_titulo,
        face  = if ("titulo" %in% textos_negrita) "bold" else "plain",
        hjust = hjust_titulo
      ),
      plot.subtitle = ggplot2::element_text(
        color = color_subtitulo,
        size  = size_subtitulo,
        face  = if ("subtitulo" %in% textos_negrita) "bold" else "plain",
        hjust = hjust_titulo
      ),
      plot.caption = ggplot2::element_text(
        color = color_nota_pie,
        size  = size_nota_pie,
        face  = if ("nota_pie" %in% textos_negrita) "bold" else "plain",
        hjust = hjust_caption
      ),

      plot.background  = ggplot2::element_rect(fill = color_fondo, color = NA),
      panel.background = ggplot2::element_rect(fill = color_fondo, color = NA)
    ) +
    ggplot2::labs(title = titulo, subtitle = subtitulo, caption = nota_pie)

  if (isTRUE(mostrar_tela) && !is.null(grid_df)) {
    p <- p + ggplot2::geom_polygon(
      data = grid_df,
      ggplot2::aes(x = .data$x, y = .data$y, group = .data$.r),
      fill = NA, color = color_grilla, linewidth = 0.5
    )
  }

  if (isTRUE(mostrar_radios) && !is.null(axes_df)) {
    p <- p + ggplot2::geom_segment(
      data = axes_df,
      ggplot2::aes(x = .data$x0, y = .data$y0, xend = .data$x1, yend = .data$y1),
      color = color_radios, linewidth = 0.5
    )
  }

  # ✅ si no se rellena: NO mapear fill (evita colores por defecto)
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
      ggplot2::aes(
        x = .data$x, y = .data$y,
        label = paste0(round(.data$.nivel * 100), "%")
      ),
      size = 3,
      color = "grey40",
      fontface = if ("niveles" %in% textos_negrita) "bold" else "plain",
      vjust = -0.2
    )
  }

  # ✅ clave: NO fijar xlim/ylim manual (produce “recortes” raros).
  # Expand con "add" da colchón real para etiquetas.
  lim_xy <- r_lim[2] * max(1.28, eje_label_mult * 1.10)
  p <- p +
    ggplot2::coord_equal(clip = "off") +
    ggplot2::scale_x_continuous(limits = c(-lim_xy, lim_xy), expand = ggplot2::expansion(mult = 0, add = 0)) +
    ggplot2::scale_y_continuous(limits = c(-lim_xy, lim_xy), expand = ggplot2::expansion(mult = 0, add = 0))

  # escalas (solo color; fill solo si rellena)
  if (!is.null(pal)) {
    p <- p + ggplot2::scale_color_manual(values = pal, breaks = as.character(grupos), drop = FALSE)
    if (isTRUE(rellenar_poligono)) {
      p <- p + ggplot2::scale_fill_manual(values = pal, breaks = as.character(grupos), drop = FALSE)
    }
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
  # CANVAS (igual filosofía que barras)
  # ---------------------------------------------------------------------------
  if (isTRUE(usar_canvas)) {
    if (!requireNamespace("cowplot", quietly = TRUE)) stop("Para `usar_canvas=TRUE` se requiere cowplot.", call. = FALSE)

    has_header  <- (!is.null(titulo) && nzchar(titulo)) || (!is.null(subtitulo) && nzchar(subtitulo))
    has_caption <- (!is.null(nota_pie) && nzchar(nota_pie))
    has_legend  <- isTRUE(mostrar_leyenda) && leg_pos != "none" && length(grupos) > 0

    p_panel <- p +
      ggplot2::labs(title = NULL, subtitle = NULL, caption = NULL) +
      ggplot2::theme(
        legend.position = "none",
        plot.margin = ggplot2::margin(0,0,0,0)
      )

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

    canvas <- canvas + cowplot::draw_plot(p_panel, x = 0, y = y_panel0, width = 1, height = panel_h)
    if (debug_ph_bordes) canvas <- canvas + .ph_border(0, y_panel0, 1, panel_h)

    if (has_legend && !is.null(leg_grob)) {
      pos_leyenda_x <- 0.5
      if (!is.na(centro_cowplot) && is.finite(centro_cowplot)) pos_leyenda_x <- centro_cowplot

      y_legend_center <- y_legend0 + legend_h * 0.5
      dy_leg <- leyenda_desplazamiento_in / h_total_in

      leg_w_npc <- grid::convertWidth(sum(leg_grob$widths), "npc", valueOnly = TRUE)
      if (!is.finite(leg_w_npc) || leg_w_npc <= 0) leg_w_npc <- 1

      canvas <- canvas + cowplot::draw_grob(
        leg_grob,
        x = pos_leyenda_x,
        y = y_legend_center + dy_leg,
        width  = leg_w_npc,
        height = legend_h,
        hjust = 0.5, vjust = 0.5
      )
      if (debug_ph_bordes) canvas <- canvas + .ph_border(0, y_legend0, 1, legend_h)
    }

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

    if (exportar == "rplot") return(canvas)

    if (is.null(path_salida) || !nzchar(path_salida)) stop("`path_salida` es requerido para exportar.", call. = FALSE)

    if (exportar == "png") {
      ggplot2::ggsave(path_salida, canvas, width = ancho, height = alto, units = "in", dpi = dpi, bg = "transparent")
      return(invisible(canvas))
    }

    if (exportar %in% c("ppt","word")) {
      if (!requireNamespace("officer", quietly = TRUE)) stop("Para exportar a PPT/Word se requiere officer.", call. = FALSE)
      if (!requireNamespace("rvg", quietly = TRUE))     stop("Para exportar a PPT/Word se requiere rvg.", call. = FALSE)

      if (exportar == "ppt") {
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

  # NO CANVAS
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
      doc <- if (ppt_append && file.exists(path_salida)) officer::read_pptx(path_salida) else officer::read_pptx()
      doc <- officer::add_slide(doc, layout = ppt_layout, master = ppt_master)
      doc <- officer::ph_with(doc, value = rvg::dml(ggobj = p), location = officer::ph_location_fullsize())
      print(doc, target = path_salida)
      return(invisible(p))
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
