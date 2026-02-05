graficar_barras_apiladas <- function(
    data,
    var_categoria,
    var_n,
    cols_porcentaje,
    etiquetas_grupos,
    escala_valor          = c("proporcion_1", "proporcion_100"),
    colores_grupos        = NULL,
    mostrar_valores       = TRUE,
    decimales             = 1,
    umbral_etiqueta       = 0.03,
    umbral_etiqueta_peq   = NULL,
    mostrar_barra_extra   = TRUE,
    barra_extra_preset    = c("ninguno", "totales", "top2box", "top3box", "bottom2box"),
    prefijo_barra_extra   = NULL,
    titulo_barra_extra    = NULL,
    barra_extra_vjust     = NULL,   # legacy (ya no se usa para título en canvas)
    titulo                = NULL,
    subtitulo             = NULL,
    nota_pie              = NULL,
    nota_pie_derecha      = NULL,
    pos_titulo            = c("centro", "izquierda", "derecha"),
    pos_nota_pie          = c("derecha", "izquierda", "centro"),
    centro_cowplot        = NA_real_,

    # Estilo de texto y layout
    color_titulo          = "#000000",
    size_titulo           = 11,
    color_subtitulo       = "#000000",
    size_subtitulo        = 9,
    color_nota_pie        = "#000000",
    size_nota_pie         = 8,
    color_leyenda         = "#000000",
    size_leyenda          = 8,
    color_texto_barras    = "white",
    size_texto_barras     = 3,
    size_texto_barras_peq = NULL,
    color_barra_extra     = "#000000",
    size_barra_extra      = 3,
    size_titulo_extra     = 3,
    color_ejes            = "#000000",
    size_ejes             = 9,
    color_fondo           = NA,

    grosor_barras         = 0.7,
    extra_derecha_rel     = 0.10,
    espacio_izquierda_rel = 0,
    ancho_max_eje_y       = NULL,

    mostrar_leyenda       = TRUE,
    usar_leyenda_cowplot  = FALSE, # legacy
    invertir_leyenda      = FALSE,
    invertir_barras       = FALSE,
    invertir_segmentos    = FALSE,
    textos_negrita        = NULL,

    # ==========================
    # CANVAS CONTROLADO
    # ==========================
    usar_canvas           = FALSE,

    canvas_w_etiquetas    = 0.38,
    canvas_w_labels       = NULL,   # legacy alias

    canvas_w_buf_etq_bars   = 0.00,
    canvas_w_buf_bars_extra = 0.00,

    canvas_w_bars         = 0.52,
    canvas_w_extra        = 0.10,

    canvas_h_header_in    = 0.75,
    canvas_h_legend_in    = 0.75,
    canvas_h_caption_in   = 0.40,
    canvas_h_panel_in     = NULL,

    canvas_h_toprow_in    = 0.18,

    # ==========================
    # CONTROL DE GROSOR
    # ==========================
    grosor_modo           = c("manual", "auto"),
    grosor_barras_mult    = 1.00,

    # ==========================
    # LEYENDA (cuadrados perfectos)
    # ==========================
    legend_key_cm         = 0.30,

    # ==========================
    # DEBUG PH
    # ==========================
    debug_ph_bordes       = FALSE,
    debug_ph_col          = "#FF00FF",
    debug_ph_lwd          = 0.6,

    exportar              = c("rplot", "png", "ppt", "word"),
    path_salida           = NULL,
    ancho                 = 10,
    alto                  = 6,
    alto_por_categoria    = NULL,
    dpi                   = 300
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y

  # deps
  if (!requireNamespace("ggplot2", quietly = TRUE)) stop("Requiere ggplot2.", call. = FALSE)
  if (!requireNamespace("dplyr", quietly = TRUE))  stop("Requiere dplyr.", call. = FALSE)
  if (!requireNamespace("tidyr", quietly = TRUE))  stop("Requiere tidyr.", call. = FALSE)

  hjust_from_pos <- function(x) {
    switch(x,
           "izquierda" = 0,
           "centro"    = 0.5,
           "derecha"   = 1,
           0.5
    )
  }

  .pt_to_mm <- function(pt) as.numeric(pt) * 0.3527777778

  escala_valor       <- match.arg(escala_valor)
  exportar           <- match.arg(exportar)
  barra_extra_preset <- match.arg(barra_extra_preset)
  pos_titulo         <- match.arg(pos_titulo)
  pos_nota_pie       <- match.arg(pos_nota_pie)
  grosor_modo        <- match.arg(grosor_modo)

  # legacy alias
  if (!is.null(canvas_w_labels) && is.finite(canvas_w_labels)) {
    canvas_w_etiquetas <- canvas_w_labels
  }

  # normalizaciones
  decimales <- suppressWarnings(as.numeric(decimales))
  if (length(decimales) < 1L || !is.finite(decimales[1])) decimales <- 1 else decimales <- decimales[1]

  size_texto_barras_peq <- size_texto_barras_peq %||% size_texto_barras
  if (is.null(umbral_etiqueta_peq)) umbral_etiqueta_peq <- umbral_etiqueta

  hjust_titulo  <- hjust_from_pos(pos_titulo)
  hjust_caption <- hjust_from_pos(pos_nota_pie)

  textos_negrita <- textos_negrita %||% character(0)

  pulso_azul  <- "#002768"
  pulso_verde <- "#5BAF31"

  # validaciones
  if (!var_categoria %in% names(data)) stop("`var_categoria` no existe en `data`.", call. = FALSE)
  if (!var_n %in% names(data))         stop("`var_n` no existe en `data`.", call. = FALSE)
  if (!all(cols_porcentaje %in% names(data))) {
    faltan <- cols_porcentaje[!cols_porcentaje %in% names(data)]
    stop("Faltan columnas en `data`: ", paste(faltan, collapse = ", "), call. = FALSE)
  }
  if (!all(names(etiquetas_grupos) %in% cols_porcentaje)) {
    stop("Los names de `etiquetas_grupos` deben coincidir con `cols_porcentaje`.", call. = FALSE)
  }

  df <- data

  # ---------------------------------------------------------------------------
  # 1) Ancho -> Largo
  # ---------------------------------------------------------------------------
  df_long <- df |>
    dplyr::select(dplyr::all_of(c(var_categoria, var_n, cols_porcentaje))) |>
    tidyr::pivot_longer(
      cols      = dplyr::all_of(cols_porcentaje),
      names_to  = ".col_pct",
      values_to = ".valor"
    ) |>
    dplyr::mutate(.grupo = dplyr::recode(.data$.col_pct, !!!etiquetas_grupos))

  if (!is.numeric(df_long$.valor)) stop("Las columnas de porcentaje deben ser numéricas.", call. = FALSE)

  df_long$.valor_plot <- if (escala_valor == "proporcion_100") df_long$.valor / 100 else df_long$.valor
  df_long$.valor_plot[!is.finite(df_long$.valor_plot) | is.na(df_long$.valor_plot)] <- 0

  # Normalizar por categoría a suma 1
  df_long <- df_long |>
    dplyr::group_by(.data[[var_categoria]]) |>
    dplyr::mutate(
      .suma_raw   = sum(.valor_plot, na.rm = TRUE),
      .valor_plot = dplyr::if_else(.suma_raw > 0, .valor_plot / .suma_raw, 0)
    ) |>
    dplyr::ungroup()

  # Blindaje (evita fuera de rango por ruido numérico)
  df_long$.valor_plot <- pmax(0, pmin(1, df_long$.valor_plot))

  # Orden de segmentos
  niveles_originales <- unname(etiquetas_grupos)
  niveles_stack   <- if (invertir_segmentos) niveles_originales else rev(niveles_originales)
  niveles_leyenda <- if (invertir_leyenda)  rev(niveles_originales) else niveles_originales
  df_long$.grupo  <- factor(df_long$.grupo, levels = niveles_stack)

  # ---------------------------------------------------------------------------
  # 1.1) ORDEN MASTER de categorías (FIJO y consistente)
  # ---------------------------------------------------------------------------
  cat_chr <- as.character(df_long[[var_categoria]])

  # Niveles en el orden en que aparecen (pero como CHARACTER)
  cat_lvls <- unique(cat_chr)

  # invertir si corresponde
  if (invertir_barras) cat_lvls <- rev(cat_lvls)

  # factor único y consistente para TODOS los plots
  df_long[[var_categoria]] <- factor(cat_chr, levels = cat_lvls)
  n_categorias <- length(cat_lvls)

  cats_df <- dplyr::tibble(
    .cat_chr = cat_lvls,
    .cat     = factor(cat_lvls, levels = cat_lvls)
  )

  # ---------------------------------------------------------------------------
  # 1.5) Grosor de barras  ✅ (define grosor_eff)
  # ---------------------------------------------------------------------------
  if (grosor_modo == "auto") {
    base <- 0.78
    adj  <- if (n_categorias <= 2) 1.00 else if (n_categorias <= 5) 0.92 else if (n_categorias <= 10) 0.85 else 0.78
    grosor_eff <- base * adj * grosor_barras_mult
    grosor_eff <- max(0.20, min(0.95, grosor_eff))
  } else {
    grosor_eff <- grosor_barras
  }

  # ---------------------------------------------------------------------------
  # 2) BARRAS — Y DISCRETA (clave para geom_col y para alinear con PH)
  # ---------------------------------------------------------------------------
  max_suma <- 1
  x_max_bars <- if (usar_canvas) {
    1
  } else {
    if (mostrar_barra_extra) max_suma * (1 + extra_derecha_rel) else max_suma
  }

  expand_x <- if (usar_canvas) {
    ggplot2::expansion(mult = c(0, 0), add = c(0, 0))
  } else {
    ggplot2::expansion(mult = c(espacio_izquierda_rel, 0.05))
  }

  p_bars <- ggplot2::ggplot(
    df_long,
    ggplot2::aes(
      x    = .data$.valor_plot,
      y    = .data[[var_categoria]],
      fill = .data$.grupo
    )
  ) +
    ggplot2::geom_col(width = grosor_eff) +
    ggplot2::scale_x_continuous(limits = c(0, x_max_bars), expand = expand_x) +
    ggplot2::scale_y_discrete(
      limits = cat_lvls, drop = FALSE,
      expand = ggplot2::expansion(mult = c(0, 0), add = c(0, 0))
    ) +
    ggplot2::coord_cartesian(clip = if (usar_canvas) "on" else "off") +
    ggplot2::theme_minimal(base_size = 9) +
    ggplot2::theme(
      panel.grid.major   = ggplot2::element_blank(),
      panel.grid.minor   = ggplot2::element_blank(),
      axis.title         = ggplot2::element_blank(),
      axis.text.x        = ggplot2::element_blank(),
      axis.ticks.x       = ggplot2::element_blank(),
      plot.background    = ggplot2::element_rect(fill = color_fondo, color = NA),
      panel.background   = ggplot2::element_rect(fill = color_fondo, color = NA),
      legend.position    = "none",
      axis.text.y        = ggplot2::element_blank(),
      axis.ticks.y       = ggplot2::element_blank(),
      plot.margin        = ggplot2::margin(0, 0, 0, 0)
    )

  # ---------------------------------------------------------------------------
  # 3) Etiquetas internas (%)
  # ---------------------------------------------------------------------------
  if (mostrar_valores) {

    niveles_fill       <- levels(df_long$.grupo)
    niveles_stack_real <- rev(niveles_fill)

    df_lab <- df_long |>
      dplyr::group_by(.data[[var_categoria]]) |>
      dplyr::arrange(factor(.grupo, levels = niveles_stack_real), .by_group = TRUE) |>
      dplyr::mutate(x_center = cumsum(.valor_plot) - .valor_plot / 2) |>
      dplyr::ungroup()

    .asignar_pct_100 <- function(p) {
      p[is.na(p) | !is.finite(p)] <- 0
      s <- sum(p)
      if (s <= 0) return(rep(0L, length(p)))
      p <- p / s
      x <- p * 100
      base <- floor(x)
      resto <- 100L - sum(base)
      if (resto > 0) {
        frac <- x - base
        idx <- order(frac, decreasing = TRUE)
        base[idx[seq_len(resto)]] <- base[idx[seq_len(resto)]] + 1L
      }
      as.integer(base)
    }

    df_lab <- df_lab |>
      dplyr::group_by(.data[[var_categoria]]) |>
      dplyr::mutate(.pct_int = .asignar_pct_100(.valor_plot),
                    lab = paste0(.pct_int, "%")) |>
      dplyr::ungroup() |>
      dplyr::mutate(
        .tamano_etq = dplyr::case_when(
          .valor_plot >= umbral_etiqueta     ~ "grande",
          .valor_plot >= umbral_etiqueta_peq ~ "peq",
          TRUE                               ~ "ninguna"
        )
      ) |>
      dplyr::filter(.tamano_etq != "ninguna", is.finite(x_center))

    df_lab_grande <- df_lab[df_lab$.tamano_etq == "grande", , drop = FALSE]
    df_lab_peq    <- df_lab[df_lab$.tamano_etq == "peq",    , drop = FALSE]

    if (nrow(df_lab_grande) > 0) {
      p_bars <- p_bars +
        ggplot2::geom_text(
          data    = df_lab_grande,
          mapping = ggplot2::aes(x = x_center, y = .data[[var_categoria]], label = lab),
          color   = color_texto_barras,
          size    = size_texto_barras,
          fontface = if ("porcentajes" %in% textos_negrita) "bold" else "plain",
          inherit.aes = FALSE
        )
    }
    if (nrow(df_lab_peq) > 0) {
      p_bars <- p_bars +
        ggplot2::geom_text(
          data    = df_lab_peq,
          mapping = ggplot2::aes(x = x_center, y = .data[[var_categoria]], label = lab),
          color   = color_texto_barras,
          size    = size_texto_barras_peq,
          fontface = if ("porcentajes" %in% textos_negrita) "bold" else "plain",
          inherit.aes = FALSE
        )
    }
  }

  # ---------------------------------------------------------------------------
  # 4) Colores + leyenda (solo para extraer grob)
  # ---------------------------------------------------------------------------
  wrap_fun <- NULL
  if (requireNamespace("stringr", quietly = TRUE)) wrap_fun <- function(x) stringr::str_wrap(x, width = 40)

  if (!is.null(colores_grupos)) {
    if (is.null(names(colores_grupos))) colores_grupos <- stats::setNames(colores_grupos, niveles_originales)
    valores_leyenda <- colores_grupos[niveles_leyenda]

    p_bars <- p_bars +
      ggplot2::scale_fill_manual(
        breaks = niveles_leyenda,
        values = valores_leyenda,
        labels = if (!is.null(wrap_fun)) wrap_fun else ggplot2::waiver()
      )
  } else {
    p_bars <- p_bars +
      ggplot2::scale_fill_discrete(
        breaks = niveles_leyenda,
        labels = if (!is.null(wrap_fun)) wrap_fun else ggplot2::waiver()
      )
  }

  n_items_leyenda <- length(niveles_leyenda)
  n_por_fila      <- 6L
  n_filas_leyenda <- max(1L, ceiling(n_items_leyenda / n_por_fila))

  p_for_legend <- p_bars +
    ggplot2::theme(
      legend.position = "bottom",
      legend.title    = ggplot2::element_blank(),
      legend.text     = ggplot2::element_text(
        color = color_leyenda,
        size  = size_leyenda,
        face  = if ("leyenda" %in% textos_negrita) "bold" else "plain"
      ),
      legend.key.width  = grid::unit(legend_key_cm, "cm"),
      legend.key.height = grid::unit(legend_key_cm, "cm"),
      plot.margin       = ggplot2::margin(0, 0, 0, 0)
    ) +
    ggplot2::guides(fill = ggplot2::guide_legend(nrow = n_filas_leyenda, byrow = TRUE))

  # ---------------------------------------------------------------------------
  # 5) PH ETIQUETAS — Y DISCRETA
  # ---------------------------------------------------------------------------
  etiquetas_df <- cats_df |>
    dplyr::mutate(.lab = .cat_chr)

  if (!is.null(ancho_max_eje_y)) {
    if (!requireNamespace("stringr", quietly = TRUE)) stop("Para `ancho_max_eje_y` se requiere stringr.", call. = FALSE)
    etiquetas_df$.lab <- stringr::str_wrap(etiquetas_df$.lab, width = ancho_max_eje_y)
  }

  p_etiquetas <- ggplot2::ggplot(etiquetas_df, ggplot2::aes(y = .data$.cat)) +
    ggplot2::geom_text(
      ggplot2::aes(x = 1, label = .data$.lab),
      hjust = 1, vjust = 0.5,
      color = color_ejes,
      size  = .pt_to_mm(size_ejes),
      fontface = if ("eje_y" %in% textos_negrita) "bold" else "plain"
    ) +
    ggplot2::scale_x_continuous(limits = c(0, 1), expand = ggplot2::expansion(mult = c(0, 0), add = c(0, 0))) +
    ggplot2::scale_y_discrete(
      limits = cat_lvls, drop = FALSE,
      expand = ggplot2::expansion(mult = c(0, 0), add = c(0, 0))
    ) +
    ggplot2::coord_cartesian(clip = "on") +
    ggplot2::theme_void() +
    ggplot2::theme(
      plot.background  = ggplot2::element_rect(fill = color_fondo, color = NA),
      panel.background = ggplot2::element_rect(fill = color_fondo, color = NA),
      plot.margin      = ggplot2::margin(0, 0, 0, 0)
    )

  # buffers (vacíos)
  p_buf <- ggplot2::ggplot(etiquetas_df, ggplot2::aes(y = .data$.cat, x = 0)) +
    ggplot2::geom_blank() +
    ggplot2::scale_x_continuous(limits = c(0, 1), expand = ggplot2::expansion(mult = c(0, 0), add = c(0, 0))) +
    ggplot2::scale_y_discrete(
      limits = cat_lvls, drop = FALSE,
      expand = ggplot2::expansion(mult = c(0, 0), add = c(0, 0))
    ) +
    ggplot2::coord_cartesian(clip = "on") +
    ggplot2::theme_void() +
    ggplot2::theme(
      plot.background  = ggplot2::element_rect(fill = color_fondo, color = NA),
      panel.background = ggplot2::element_rect(fill = color_fondo, color = NA),
      plot.margin      = ggplot2::margin(0, 0, 0, 0)
    )

  # ---------------------------------------------------------------------------
  # 6) PH EXTRA — Y DISCRETA + valores
  # ---------------------------------------------------------------------------
  df_wide_extra <- df |>
    dplyr::select(dplyr::all_of(c(var_categoria, var_n, cols_porcentaje))) |>
    dplyr::mutate(valor_extra = .data[[var_n]])

  prefijo_extra_int     <- prefijo_barra_extra %||% ""
  titulo_extra_int      <- titulo_barra_extra
  color_barra_extra_int <- color_barra_extra
  fontface_barra_extra  <- if ("barra_extra" %in% textos_negrita) "bold" else "plain"

  if (barra_extra_preset != "ninguno") {

    if (barra_extra_preset == "totales") {
      if (is.null(titulo_barra_extra) || !nzchar(titulo_barra_extra)) titulo_extra_int <- "Total"
      if (is.null(prefijo_barra_extra)) prefijo_extra_int <- "N = "
      if (is.null(color_barra_extra))   color_barra_extra_int <- pulso_azul
      fontface_barra_extra <- "bold"
    } else {

      base_mat <- df_wide_extra[, cols_porcentaje, drop = FALSE]
      if (escala_valor == "proporcion_100") base_mat <- base_mat / 100
      ordenado <- t(apply(as.matrix(base_mat), 1, sort, decreasing = TRUE))

      if (barra_extra_preset == "top2box") {
        df_wide_extra$valor_extra <- ordenado[, 1] + ordenado[, 2]
        if (is.null(titulo_barra_extra) || !nzchar(titulo_barra_extra)) titulo_extra_int <- "TOP TWO BOX"
      } else if (barra_extra_preset == "top3box") {
        df_wide_extra$valor_extra <- ordenado[, 1] + ordenado[, 2] + ordenado[, 3]
        if (is.null(titulo_barra_extra) || !nzchar(titulo_barra_extra)) titulo_extra_int <- "TOP THREE BOX"
      } else if (barra_extra_preset == "bottom2box") {
        nc <- ncol(ordenado)
        df_wide_extra$valor_extra <- ordenado[, nc] + ordenado[, nc - 1]
        if (is.null(titulo_barra_extra) || !nzchar(titulo_barra_extra)) titulo_extra_int <- "BOTTOM TWO BOX"
      }

      df_wide_extra$valor_extra <- df_wide_extra$valor_extra * 100
      color_barra_extra_int <- pulso_verde
      fontface_barra_extra  <- "bold"
    }
  }

  p_extra <- ggplot2::ggplot(etiquetas_df, ggplot2::aes(y = .data$.cat)) +
    ggplot2::geom_blank() +
    ggplot2::scale_x_continuous(limits = c(0, 1), expand = ggplot2::expansion(mult = c(0, 0), add = c(0, 0))) +
    ggplot2::scale_y_discrete(
      limits = cat_lvls, drop = FALSE,
      expand = ggplot2::expansion(mult = c(0, 0), add = c(0, 0))
    ) +
    ggplot2::coord_cartesian(clip = "on") +
    ggplot2::theme_void() +
    ggplot2::theme(
      plot.background  = ggplot2::element_rect(fill = color_fondo, color = NA),
      panel.background = ggplot2::element_rect(fill = color_fondo, color = NA),
      plot.margin      = ggplot2::margin(0, 0, 0, 0)
    )

  if (mostrar_barra_extra) {

    .format_pct_clean <- function(x) {
      x_round <- round(x, 1)
      txt <- format(x_round, nsmall = 1, trim = TRUE, scientific = FALSE)
      sub("\\.0$", "", txt)
    }

    df_extra_vals <- df_wide_extra |>
      dplyr::select(dplyr::all_of(c(var_categoria, "valor_extra"))) |>
      dplyr::mutate(
        .cat_chr = as.character(.data[[var_categoria]]),
        .cat     = factor(.cat_chr, levels = cat_lvls),
        lab_valor = dplyr::case_when(
          barra_extra_preset %in% c("top2box", "top3box", "bottom2box") ~ paste0(.format_pct_clean(valor_extra), "%"),
          TRUE ~ format(valor_extra, big.mark = ",", scientific = FALSE, trim = TRUE)
        ),
        lab_extra = paste0(prefijo_extra_int, lab_valor)
      ) |>
      dplyr::arrange(.data$.cat)

    p_extra <- p_extra +
      ggplot2::geom_text(
        data = df_extra_vals,
        ggplot2::aes(y = .data$.cat, x = 0.5, label = lab_extra),
        inherit.aes = FALSE,
        hjust = 0.5, vjust = 0.5,
        color = color_barra_extra_int,
        size  = .pt_to_mm(size_barra_extra),
        fontface = fontface_barra_extra
      )
  }

  # ---------------------------------------------------------------------------
  # 6.9) SINCRONIZAR ALTURAS DE PANEL (gtable) -> alineación vertical real
  # ---------------------------------------------------------------------------
  if (!requireNamespace("cowplot", quietly = TRUE)) stop("Requiere cowplot.", call. = FALSE)

  # helper: iguala alturas del PANEL (y, si quieres, de todo el gtable)
  .sync_panel_heights <- function(g_target, g_ref, full = TRUE) {
    # filas del panel
    pr <- unique(g_ref$layout$t[g_ref$layout$name == "panel"])
    pt <- unique(g_target$layout$t[g_target$layout$name == "panel"])

    if (length(pr) == 1 && length(pt) == 1) {
      # copia altura panel
      g_target$heights[pt] <- g_ref$heights[pr]
    }

    # opcional: copia TODAS las alturas -> elimina diferencias sutiles arriba/abajo
    if (isTRUE(full) && length(g_target$heights) == length(g_ref$heights)) {
      g_target$heights <- g_ref$heights
    }

    g_target
  }

  # pasar a grobs
  g_bars <- ggplot2::ggplotGrob(p_bars)
  g_etq  <- ggplot2::ggplotGrob(p_etiquetas)
  g_ext  <- ggplot2::ggplotGrob(p_extra)

  # sincronizar alturas contra barras (referencia)
  g_etq <- .sync_panel_heights(g_etq, g_bars, full = TRUE)
  g_ext <- .sync_panel_heights(g_ext, g_bars, full = TRUE)

  # buffers como grob (spacer)
  g_buf1 <- grid::nullGrob()
  g_buf2 <- grid::nullGrob()

  # ---------------------------------------------------------------------------
  # 7) Caption
  # ---------------------------------------------------------------------------
  caption_text <- NULL
  if (!is.null(nota_pie) && nzchar(nota_pie) && !is.null(nota_pie_derecha) && nzchar(nota_pie_derecha)) {
    caption_text <- paste0(nota_pie, "   ", nota_pie_derecha)
  } else if (!is.null(nota_pie) && nzchar(nota_pie)) {
    caption_text <- nota_pie
  } else if (!is.null(nota_pie_derecha) && nzchar(nota_pie_derecha)) {
    caption_text <- nota_pie_derecha
  }

  # ---------------------------------------------------------------------------
  # 8) No canvas
  # ---------------------------------------------------------------------------
  if (!usar_canvas) {
    out <- p_bars +
      ggplot2::theme(
        legend.position = if (mostrar_leyenda) "bottom" else "none"
      ) +
      ggplot2::labs(title = titulo, subtitle = subtitulo, caption = caption_text)

    if (exportar == "rplot") return(out)
    stop("Exportación fuera de canvas no está activada en este bloque.", call. = FALSE)
  }

  # ---------------------------------------------------------------------------
  # 9) CANVAS
  # ---------------------------------------------------------------------------
  if (!requireNamespace("cowplot", quietly = TRUE)) stop("Para `usar_canvas=TRUE` se requiere cowplot.", call. = FALSE)
  if (!requireNamespace("grid", quietly = TRUE))    stop("Para `usar_canvas=TRUE` se requiere grid.", call. = FALSE)

  .ph_border <- function(x, y, w, h) {
    cowplot::draw_grob(
      grid::rectGrob(
        x = 0, y = 0, width = 1, height = 1,
        just = c("left", "bottom"),
        gp = grid::gpar(col = debug_ph_col, fill = NA, lwd = debug_ph_lwd)
      ),
      x = x, y = y, width = w, height = h,
      hjust = 0, vjust = 0
    )
  }

  alto_por_cat_eff <- alto_por_categoria %||% 0.42
  h_panel_in <- if (!is.null(canvas_h_panel_in) && is.finite(canvas_h_panel_in) && canvas_h_panel_in > 0) {
    canvas_h_panel_in
  } else {
    max(1L, n_categorias) * alto_por_cat_eff
  }

  has_header  <- (!is.null(titulo) && nzchar(titulo)) || (!is.null(subtitulo) && nzchar(subtitulo))
  has_legend  <- isTRUE(mostrar_leyenda) && length(niveles_leyenda) > 0
  has_caption <- !is.null(caption_text) && nzchar(caption_text)

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

  # widths (5 columnas)
  w_etq   <- canvas_w_etiquetas
  w_buf1  <- canvas_w_buf_etq_bars
  w_bars  <- canvas_w_bars
  w_buf2  <- canvas_w_buf_bars_extra
  w_extra <- canvas_w_extra

  w_sum <- w_etq + w_buf1 + w_bars + w_buf2 + w_extra
  if (!is.finite(w_sum) || w_sum <= 0) w_sum <- 1

  w_etq   <- w_etq   / w_sum
  w_buf1  <- w_buf1  / w_sum
  w_bars  <- w_bars  / w_sum
  w_buf2  <- w_buf2  / w_sum
  w_extra <- w_extra / w_sum

  x_etq0   <- 0
  x_buf10  <- x_etq0 + w_etq
  x_bars0  <- x_buf10 + w_buf1
  x_buf20  <- x_bars0 + w_bars
  x_extra0 <- x_buf20 + w_buf2

  # top row (pulgadas -> fracción)
  top_in <- canvas_h_toprow_in %||% 0
  if (!is.finite(top_in) || is.na(top_in) || top_in < 0) top_in <- 0
  top_in <- min(top_in, h_panel_in * 0.45)
  top_h  <- if (top_in > 0) top_in / h_total_in else 0

  main_h  <- panel_h - top_h
  y_top0  <- y_panel0 + main_h
  y_main0 <- y_panel0

  # leyenda grob
  leg_grob <- NULL
  if (has_legend) {
    leg_grob <- cowplot::get_legend(
      p_for_legend + ggplot2::theme(
        legend.position  = "bottom",
        legend.direction = "horizontal",
        legend.box       = "horizontal"
      )
    )
  }

  canvas <- cowplot::ggdraw()

  # HEADER
  if (has_header) {
    header_text <- titulo %||% ""
    sub_text    <- subtitulo %||% ""

    if (nzchar(header_text)) {
      canvas <- canvas + cowplot::draw_text(
        text  = header_text,
        x     = hjust_titulo,
        y     = 1 - (header_h * 0.35),
        hjust = hjust_titulo,
        vjust = 0.5,
        size  = size_titulo,
        colour= color_titulo,
        fontface = if ("titulo" %in% textos_negrita) "bold" else "plain"
      )
    }
    if (nzchar(sub_text)) {
      canvas <- canvas + cowplot::draw_text(
        text  = sub_text,
        x     = hjust_titulo,
        y     = 1 - (header_h * 0.78),
        hjust = hjust_titulo,
        vjust = 0.5,
        size  = size_subtitulo,
        colour= color_subtitulo
      )
    }

    if (debug_ph_bordes) canvas <- canvas + .ph_border(0, y_header0, 1, header_h)
  }

  # TOP ROW (3 PH): etiquetas | barras | extra
  if (top_h > 0) {

    if (debug_ph_bordes) {
      canvas <- canvas +
        .ph_border(x_etq0,   y_top0, w_etq,   top_h) +
        .ph_border(x_bars0,  y_top0, w_bars,  top_h) +
        .ph_border(x_extra0, y_top0, w_extra, top_h)
    }

    if (isTRUE(mostrar_barra_extra) && !is.null(titulo_extra_int) && nzchar(titulo_extra_int)) {
      canvas <- canvas + cowplot::draw_text(
        text     = titulo_extra_int,
        x        = x_extra0 + (w_extra * 0.5),
        y        = y_top0 + (top_h * 0.08),
        hjust    = 0.5,
        vjust    = 0,
        size     = size_titulo_extra,
        colour   = color_barra_extra_int,
        fontface = "bold"
      )
    }
  }

  # MAIN ROW (5 PH): etiquetas | buf | barras | buf | extra
  canvas <- canvas +
    cowplot::draw_grob(g_etq,  x = x_etq0,   y = y_main0, width = w_etq,   height = main_h) +
    cowplot::draw_grob(g_buf1, x = x_buf10,  y = y_main0, width = w_buf1,  height = main_h) +
    cowplot::draw_grob(g_bars, x = x_bars0,  y = y_main0, width = w_bars,  height = main_h) +
    cowplot::draw_grob(g_buf2, x = x_buf20,  y = y_main0, width = w_buf2,  height = main_h) +
    cowplot::draw_grob(g_ext,  x = x_extra0, y = y_main0, width = w_extra, height = main_h)

  if (debug_ph_bordes) {
    canvas <- canvas +
      .ph_border(x_etq0,   y_main0, w_etq,   main_h) +
      .ph_border(x_buf10,  y_main0, w_buf1,  main_h) +
      .ph_border(x_bars0,  y_main0, w_bars,  main_h) +
      .ph_border(x_buf20,  y_main0, w_buf2,  main_h) +
      .ph_border(x_extra0, y_main0, w_extra, main_h)
  }

  # LEYENDA
  if (has_legend && !is.null(leg_grob)) {
    pos_leyenda_x <- 0.5
    if (!is.na(centro_cowplot) && is.finite(centro_cowplot)) pos_leyenda_x <- centro_cowplot

    canvas <- canvas + cowplot::draw_grob(
      leg_grob,
      x = pos_leyenda_x, y = y_legend0,
      width = 1, height = legend_h,
      hjust = 0.5, vjust = 0
    )
    if (debug_ph_bordes) canvas <- canvas + .ph_border(0, y_legend0, 1, legend_h)
  }

  # CAPTION
  if (has_caption) {
    canvas <- canvas + cowplot::draw_text(
      text  = caption_text,
      x     = hjust_caption,
      y     = y_caption0 + (caption_h * 0.35),
      hjust = hjust_caption,
      vjust = 0.5,
      size  = size_nota_pie,
      colour= color_nota_pie
    )
    if (debug_ph_bordes) canvas <- canvas + .ph_border(0, y_caption0, 1, caption_h)
  }

  if (exportar == "rplot") {
    attr(canvas, "alto_word_sugerido") <- h_total_in
    return(canvas)
  }

  stop("Exportación: primero validar en rplot; luego se integra a tu pipeline.", call. = FALSE)
}
