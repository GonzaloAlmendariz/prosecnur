
#' Graficar barras apiladas con canvas controlado y placeholders independientes
#'
#' Genera un gráfico de barras apiladas horizontales (100%) por categoría, con opción
#' de construir un **canvas** (vía {cowplot}) que separa en **placeholders independientes**
#' las **etiquetas del eje Y**, el **panel de barras**, y una **columna extra** (p.ej. `N=`),
#' preservando además buffers entre columnas y permitiendo alineación vertical **milimétrica**
#' mediante coordenadas NPC por fila.
#'
#' La función normaliza los porcentajes por fila a suma 1 y aplica un “cierre exacto”
#' que ajusta solo un segmento por categoría para absorber residuales numéricos y evitar
#' warnings tipo “Removed 1 row … outside the scale range”.
#'
#' @param data data.frame con las columnas requeridas.
#' @param var_categoria Nombre (string) de la variable categórica (eje Y).
#' @param var_n Nombre (string) de la variable numérica para la barra extra (p.ej. N).
#' @param cols_porcentaje Vector de nombres (string) de columnas de porcentajes (formato ancho).
#' @param etiquetas_grupos Vector nombrado que mapea `cols_porcentaje` -> etiqueta del grupo
#'   (p.ej. `c(pct_1="Muy insatisfecho", ...)`). Los `names()` deben coincidir con `cols_porcentaje`.
#' @param escala_valor Escala de entrada: `"proporcion_1"` (0-1) o `"proporcion_100"` (0-100).
#' @param colores_grupos Vector nombrado (opcional) de colores por etiqueta de grupo.
#' @param mostrar_valores Si `TRUE`, dibuja etiquetas % dentro de segmentos (cuando superan umbrales).
#' @param decimales Número de decimales para etiquetas % (se asignan por “largest remainder”
#'   en la escala 100*10^decimales, garantizando suma exacta).
#' @param umbral_etiqueta Umbral (0-1) para etiquetar como “grande”.
#' @param umbral_etiqueta_peq Umbral (0-1) para etiquetar como “peq”. Si `NULL`, usa `umbral_etiqueta`.
#' @param mostrar_barra_extra Si `TRUE`, dibuja columna extra (texto) alineada por fila.
#' @param barra_extra_preset Preset de extra: `"ninguno"`, `"totales"`, `"top2box"`, `"top3box"`, `"bottom2box"`.
#' @param prefijo_barra_extra Prefijo para el texto extra (p.ej. `"N = "`).
#' @param titulo_barra_extra Título arriba del placeholder extra (solo en canvas).
#' @param barra_extra_vjust Legacy (sin uso en canvas); se conserva por compatibilidad.
#' @param titulo Título general.
#' @param subtitulo Subtítulo general.
#' @param nota_pie Nota al pie izquierda.
#' @param nota_pie_derecha Nota al pie derecha.
#' @param pos_titulo Alineación horizontal del título: `"centro"`, `"izquierda"`, `"derecha"`.
#' @param pos_nota_pie Alineación horizontal del caption: `"derecha"`, `"izquierda"`, `"centro"`.
#' @param centro_cowplot Centro horizontal de la leyenda en canvas (0-1). `NA` usa 0.5.
#'
#' @param color_titulo,color_subtitulo,color_nota_pie,color_leyenda,color_texto_barras,color_barra_extra,color_ejes Colores.
#' @param size_titulo,size_subtitulo,size_nota_pie,size_leyenda,size_texto_barras,size_texto_barras_peq,size_barra_extra,size_titulo_extra,size_ejes Tamaños (puntos).
#' @param color_fondo Color de fondo del plot/canvas (NA para transparente).
#'
#' @param grosor_barras Grosor de barras (solo cuando `grosor_modo="manual"`).
#' @param extra_derecha_rel Expansión derecha del eje x cuando NO canvas y se quiere espacio extra.
#' @param espacio_izquierda_rel Expansión izquierda del eje x cuando NO canvas.
#' @param ancho_max_eje_y Si se define, aplica `stringr::str_wrap()` a etiquetas del eje Y.
#'
#' @param mostrar_leyenda Si `TRUE`, incluye leyenda (en canvas, debajo del panel).
#' @param invertir_leyenda Si `TRUE`, invierte el orden de la leyenda.
#' @param invertir_barras Si `TRUE`, invierte el orden de categorías.
#' @param invertir_segmentos Si `TRUE`, invierte el orden de apilamiento.
#' @param textos_negrita Vector de claves que activa negrita: `c("titulo","leyenda","barra_extra","eje_y","porcentajes")`.
#'
#' @param usar_canvas Si `TRUE`, construye un canvas con placeholders independientes.
#' @param canvas_w_etiquetas Ancho relativo de placeholder etiquetas (columna izquierda).
#' @param canvas_w_labels Alias legacy de `canvas_w_etiquetas`.
#' @param canvas_w_buf_etq_bars Ancho relativo buffer entre etiquetas y barras.
#' @param canvas_w_buf_bars_extra Ancho relativo buffer entre barras y extra.
#' @param canvas_w_bars Ancho relativo placeholder barras.
#' @param canvas_w_extra Ancho relativo placeholder extra (columna derecha).
#'
#' @param canvas_h_header_in Altura del header en pulgadas (cuando hay título/subtítulo).
#' @param canvas_h_legend_in Altura de la leyenda en pulgadas (cuando hay leyenda).
#' @param canvas_h_caption_in Altura del caption en pulgadas (cuando hay caption).
#' @param canvas_h_panel_in Altura del panel en pulgadas (si `NULL`, se usa `alto_por_categoria * n_categorias`).
#' @param canvas_h_toprow_in Altura (en pulgadas) de la fila superior del panel (para título del extra).
#'
#' @param grosor_modo `"manual"` o `"auto"`. En `"auto"` el grosor se ajusta por número de categorías.
#' @param grosor_barras_mult Multiplicador del grosor en modo `"auto"`.
#'
#' @param legend_key_cm Tamaño del `key` de la leyenda (en cm).
#' @param legend_espaciado Espaciado horizontal entre ítems de la leyenda.
#' @param legend_n_por_fila Número de ítems por fila en la leyenda.
#'
#' @param encabezado_desplazamiento_in Desplazamiento vertical (pulgadas) del bloque título/subtítulo.
#' @param encabezado_separacion_in Separación vertical total (pulgadas) entre título y subtítulo.
#' @param leyenda_desplazamiento_in Desplazamiento vertical (pulgadas) de la leyenda dentro de su placeholder.
#'
#' @param debug_ph_bordes Si `TRUE`, dibuja bordes de placeholders (debug).
#' @param debug_ph_col Color de borde debug.
#' @param debug_ph_lwd Grosor de borde debug.
#'
#' @param exportar `"rplot"`, `"png"`, `"ppt"` o `"word"`.
#' @param path_salida Ruta de salida para exportación.
#' @param ancho,alto Tamaños (pulgadas) para exportación (png/ppt/word).
#' @param alto_por_categoria Altura por categoría (pulgadas) para definir el panel cuando `canvas_h_panel_in` es `NULL`.
#' @param dpi DPI para exportación PNG.
#' @param ppt_append Si `TRUE` y el archivo existe, agrega una diapositiva al final.
#' @param ppt_layout,ppt_master Layout/master para la diapositiva.
#'
#' @return Un objeto ggplot/cowplot. Si `exportar="rplot"`, retorna el plot/canvas listo para imprimir.
#' @export
graficar_barras_apiladas <- function(
    data,
    var_categoria,
    var_n,
    cols_porcentaje,
    etiquetas_grupos,
    escala_valor          = c("proporcion_1", "proporcion_100"),
    colores_grupos        = NULL,
    mostrar_valores       = TRUE,
    decimales             = 0,
    umbral_etiqueta       = 0.001,
    umbral_etiqueta_peq   = NULL,
    mostrar_barra_extra   = TRUE,
    barra_extra_preset    = c("ninguno", "totales", "top2box", "top3box", "bottom2box"),
    prefijo_barra_extra   = NULL,
    titulo_barra_extra    = NULL,
    barra_extra_vjust     = NULL,   # legacy
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
    invertir_leyenda      = FALSE,
    invertir_barras       = FALSE,
    invertir_segmentos    = FALSE,
    textos_negrita        = NULL,

    # ==========================
    # CANVAS CONTROLADO
    # ==========================
    usar_canvas           = FALSE,

    canvas_w_etiquetas      = 0.38,
    canvas_w_buf_etq_bars   = 0.00,
    canvas_w_buf_bars_extra = 0.00,
    canvas_w_bars           = 0.52,
    canvas_w_extra          = 0.10,

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
    # LEYENDA
    # ==========================
    legend_key_cm         = 0.30,
    legend_espaciado      = 0.20,
    legend_n_por_fila     = 6L,

    # ==========================
    # AJUSTES POSICIONALES
    # ==========================
    encabezado_desplazamiento_in = 0,
    encabezado_separacion_in     = 0.14,
    leyenda_desplazamiento_in    = 0,

    # ==========================
    # DEBUG PH
    # ==========================
    debug_ph_bordes       = FALSE,
    debug_ph_col          = "#FF00FF",
    debug_ph_lwd          = 0.6,

    # ==========================
    # EXPORTAR
    # ==========================
    exportar              = c("rplot", "png", "ppt", "word"),
    path_salida           = NULL,
    ancho                 = 10,
    alto                  = 6,
    alto_por_categoria    = NULL,
    dpi                   = 300,

    ppt_append            = TRUE,
    ppt_layout            = "Blank",
    ppt_master            = "Office Theme"
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y
  hjust_from_pos <- function(x) switch(x, "izquierda" = 0, "centro" = 0.5, "derecha" = 1, 0.5)

  # deps
  if (!requireNamespace("ggplot2", quietly = TRUE)) stop("Requiere ggplot2.", call. = FALSE)
  if (!requireNamespace("dplyr", quietly = TRUE))  stop("Requiere dplyr.", call. = FALSE)
  if (!requireNamespace("tidyr", quietly = TRUE))  stop("Requiere tidyr.", call. = FALSE)
  if (!requireNamespace("grid", quietly = TRUE))   stop("Requiere grid.", call. = FALSE)

  escala_valor       <- match.arg(escala_valor)
  exportar           <- match.arg(exportar)
  barra_extra_preset <- match.arg(barra_extra_preset)
  pos_titulo         <- match.arg(pos_titulo)
  pos_nota_pie       <- match.arg(pos_nota_pie)
  grosor_modo        <- match.arg(grosor_modo)

  # legacy alias
  if (!is.null(canvas_w_labels) && is.finite(canvas_w_labels)) canvas_w_etiquetas <- canvas_w_labels

  # normalizaciones
  decimales <- suppressWarnings(as.integer(decimales))
  if (length(decimales) < 1L || !is.finite(decimales[1]) || decimales[1] < 0L) decimales <- 0L else decimales <- decimales[1]
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

  # Blindaje
  df_long$.valor_plot <- pmax(0, pmin(1, df_long$.valor_plot))

  # Orden de segmentos (DEBE IR ANTES del cierre exacto)
  niveles_originales <- unname(etiquetas_grupos)
  niveles_stack      <- if (invertir_segmentos) niveles_originales else rev(niveles_originales)
  niveles_leyenda    <- if (invertir_leyenda)  rev(niveles_originales) else niveles_originales
  df_long$.grupo     <- factor(df_long$.grupo, levels = niveles_stack)

  # ---------------------------------------------------------------------------
  # 1.05) CIERRE EXACTO A 1
  # Ajusta SOLO el ÚLTIMO del stack (derecha) para absorber residuo numérico.
  # ---------------------------------------------------------------------------
  target_level <- tail(niveles_stack, 1)

  df_long <- df_long |>
    dplyr::group_by(.data[[var_categoria]]) |>
    dplyr::mutate(
      .sum1  = sum(.valor_plot, na.rm = TRUE),
      .delta = 1 - .sum1,
      .valor_plot = dplyr::if_else(
        .data$.grupo == target_level,
        .valor_plot + .delta,
        .valor_plot
      ),
      .valor_plot = pmax(0, .valor_plot)
    ) |>
    dplyr::mutate(
      .sum2 = sum(.valor_plot, na.rm = TRUE),
      .valor_plot = dplyr::if_else(.sum2 > 0, .valor_plot / .sum2, 0)
    ) |>
    dplyr::ungroup() |>
    dplyr::select(-.sum1, -.delta, -.sum2)

  # ---------------------------------------------------------------------------
  # 1.1) ORDEN MASTER de categorías (FIJO)
  # ---------------------------------------------------------------------------
  cat_chr  <- as.character(df_long[[var_categoria]])
  cat_lvls <- unique(cat_chr)
  if (invertir_barras) cat_lvls <- rev(cat_lvls)

  df_long[[var_categoria]] <- factor(cat_chr, levels = cat_lvls)
  n_categorias <- length(cat_lvls)

  # ---------------------------------------------------------------------------
  # 1.5) Grosor de barras
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
  # 2) BARRAS
  # ---------------------------------------------------------------------------
  max_suma <- 1
  x_max_bars <- if (usar_canvas) 1 else if (mostrar_barra_extra) max_suma * (1 + extra_derecha_rel) else max_suma

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
      panel.grid.major = ggplot2::element_blank(),
      panel.grid.minor = ggplot2::element_blank(),
      axis.title       = ggplot2::element_blank(),
      axis.text.x      = ggplot2::element_blank(),
      axis.ticks.x     = ggplot2::element_blank(),
      legend.position  = "none",
      axis.text.y      = ggplot2::element_blank(),
      axis.ticks.y     = ggplot2::element_blank(),
      plot.background  = ggplot2::element_rect(fill = color_fondo, color = NA),
      panel.background = ggplot2::element_rect(fill = color_fondo, color = NA),
      plot.margin      = ggplot2::margin(0,0,0,0)
    )

  # ---------------------------------------------------------------------------
  # 3) Etiquetas internas (%) con asignación exacta (suma 100.00 si decimales=2, etc.)
  # ---------------------------------------------------------------------------
  if (isTRUE(mostrar_valores)) {

    niveles_fill       <- levels(df_long$.grupo)
    niveles_stack_real <- rev(niveles_fill)

    df_lab <- df_long |>
      dplyr::group_by(.data[[var_categoria]]) |>
      dplyr::arrange(factor(.grupo, levels = niveles_stack_real), .by_group = TRUE) |>
      dplyr::mutate(x_center = cumsum(.valor_plot) - .valor_plot / 2) |>
      dplyr::ungroup()

    .asignar_pct_exacto <- function(p, dec) {
      p[is.na(p) | !is.finite(p)] <- 0
      s <- sum(p)
      if (s <= 0) return(rep.int(0L, length(p)))
      p <- p / s

      escala <- 10^dec
      target_units <- as.integer(100L * escala)

      x_units <- p * target_units
      base <- floor(x_units)
      resto <- target_units - sum(base)

      if (resto > 0L) {
        frac <- x_units - base
        idx <- order(frac, decreasing = TRUE)
        base[idx[seq_len(resto)]] <- base[idx[seq_len(resto)]] + 1L
      }
      as.integer(base)
    }

    .fmt_units_pct <- function(units, dec){
      escala <- 10^dec
      val <- units / escala
      out <- format(val, nsmall = dec, trim = TRUE, scientific = FALSE)
      paste0(out, "%")
    }

    df_lab <- df_lab |>
      dplyr::group_by(.data[[var_categoria]]) |>
      dplyr::mutate(
        .pct_units = .asignar_pct_exacto(.valor_plot, decimales),
        lab        = .fmt_units_pct(.pct_units, decimales)
      ) |>
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
  # 4) Colores + leyenda (para extraer grob) — con separación horizontal REAL
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
  n_por_fila <- as.integer(legend_n_por_fila)
  if (!is.finite(n_por_fila) || n_por_fila < 1L) n_por_fila <- 6L

  p_for_legend <- p_bars +
    ggplot2::theme(
      legend.position = "bottom",
      legend.title    = ggplot2::element_blank(),
      legend.text = ggplot2::element_text(
        color = color_leyenda,
        size  = size_leyenda,
        face  = if ("leyenda" %in% textos_negrita) "bold" else "plain",
        # esto crea separación real entre categorías sin deformar el key
        margin = ggplot2::margin(r = legend_espaciado, unit = "pt")
      ),

      legend.key.width  = grid::unit(legend_key_cm, "cm"),
      legend.key.height = grid::unit(legend_key_cm, "cm"),

      legend.key.spacing.x = grid::unit(0.10, "cm"),

      plot.margin = ggplot2::margin(0, 0, 0, 0)
    ) +
    ggplot2::guides(
      fill = ggplot2::guide_legend(
        byrow = TRUE,
        ncol  = n_por_fila,
        keywidth  = grid::unit(legend_key_cm, "cm"),
        keyheight = grid::unit(legend_key_cm, "cm")
      )
    )

  # ---------------------------------------------------------------------------
  # 5) Etiquetas Y y extra como texto (sin ggplot)
  # ---------------------------------------------------------------------------
  etiquetas_vec <- cat_lvls
  if (!is.null(ancho_max_eje_y)) {
    if (!requireNamespace("stringr", quietly = TRUE)) stop("Para `ancho_max_eje_y` se requiere stringr.", call. = FALSE)
    etiquetas_vec <- stringr::str_wrap(etiquetas_vec, width = ancho_max_eje_y)
  }

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

  .format_pct_clean <- function(x, dec){
    x_round <- round(x, dec)
    format(x_round, nsmall = dec, trim = TRUE, scientific = FALSE)
  }

  extra_map <- df_wide_extra |>
    dplyr::mutate(.cat_chr = as.character(.data[[var_categoria]])) |>
    dplyr::select(.cat_chr, valor_extra)

  extra_vals <- vapply(cat_lvls, function(cc) {
    vv <- extra_map$valor_extra[match(cc, extra_map$.cat_chr)]
    if (length(vv) == 0 || is.na(vv)) vv <- NA_real_
    vv
  }, numeric(1))

  extra_labels <- rep("", length(cat_lvls))
  if (isTRUE(mostrar_barra_extra)) {
    extra_labels <- if (barra_extra_preset %in% c("top2box", "top3box", "bottom2box")) {
      paste0(prefijo_extra_int, .format_pct_clean(extra_vals, decimales), "%")
    } else {
      paste0(prefijo_extra_int, format(extra_vals, big.mark = ",", scientific = FALSE, trim = TRUE))
    }
    extra_labels[!is.finite(extra_vals)] <- ""
  }

  # ---------------------------------------------------------------------------
  # 7) Caption (texto)
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
  if (!isTRUE(usar_canvas)) {
    out <- p_bars +
      ggplot2::theme(legend.position = if (mostrar_leyenda) "bottom" else "none") +
      ggplot2::labs(title = titulo, subtitle = subtitulo, caption = caption_text)

    if (exportar == "rplot") return(out)

    # EXPORT PNG / PPT / WORD (sin canvas): se exporta el ggplot directamente
    if (is.null(path_salida) || !nzchar(path_salida)) stop("`path_salida` es requerido para exportar.", call. = FALSE)

    if (exportar == "png") {
      ggplot2::ggsave(filename = path_salida, plot = out, width = ancho, height = alto, units = "in", dpi = dpi, bg = "transparent")
      return(invisible(out))
    }

    if (exportar %in% c("ppt", "word")) {
      if (!requireNamespace("officer", quietly = TRUE)) stop("Para exportar a PPT/Word se requiere officer.", call. = FALSE)
      if (!requireNamespace("rvg", quietly = TRUE))     stop("Para exportar a PPT/Word se requiere rvg (dml).", call. = FALSE)

      if (exportar == "ppt") {
        doc <- if (ppt_append && file.exists(path_salida)) officer::read_pptx(path_salida) else officer::read_pptx()
        doc <- officer::add_slide(doc, layout = ppt_layout, master = ppt_master)
        doc <- officer::ph_with(
          doc,
          value = rvg::dml(ggobj = out),
          location = officer::ph_location_fullsize()
        )
        print(doc, target = path_salida)
        return(invisible(out))
      }

      if (exportar == "word") {
        doc <- if (file.exists(path_salida)) officer::read_docx(path_salida) else officer::read_docx()
        doc <- officer::body_add_par(doc, value = "", style = "Normal")
        doc <- officer::body_add_dml(
          doc,
          value = rvg::dml(ggobj = out),
          width = ancho, height = alto
        )
        print(doc, target = path_salida)
        return(invisible(out))
      }
    }

    stop("Tipo de exportación no soportado.", call. = FALSE)
  }

  # ---------------------------------------------------------------------------
  # 9) CANVAS (cowplot)
  # ---------------------------------------------------------------------------
  if (!requireNamespace("cowplot", quietly = TRUE)) stop("Para `usar_canvas=TRUE` se requiere cowplot.", call. = FALSE)

  # barras “panel puro”
  p_bars_panel <- p_bars +
    ggplot2::theme_void() +
    ggplot2::theme(
      legend.position  = "none",
      plot.background  = ggplot2::element_rect(fill = color_fondo, color = NA),
      panel.background = ggplot2::element_rect(fill = color_fondo, color = NA),
      plot.margin      = ggplot2::margin(0,0,0,0)
    )

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

  # alturas en pulgadas
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

  # widths (5 columnas) — placeholders independientes
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

  # top row (título del extra)
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

  # ============================================================
  # HEADER: centrado + desplazamiento + separación
  # ============================================================
  if (has_header) {
    y_header_center <- y_header0 + (header_h * 0.5)
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
        text  = titulo,
        x     = hjust_titulo,
        y     = y_title,
        hjust = hjust_titulo,
        vjust = 0.5,
        size  = size_titulo,
        colour= color_titulo,
        fontface = if ("titulo" %in% textos_negrita) "bold" else "plain"
      )
    }

    if (has_s) {
      canvas <- canvas + cowplot::draw_text(
        text  = subtitulo,
        x     = hjust_titulo,
        y     = y_sub,
        hjust = hjust_titulo,
        vjust = 0.5,
        size  = size_subtitulo,
        colour= color_subtitulo
      )
    }

    if (debug_ph_bordes) canvas <- canvas + .ph_border(0, y_header0, 1, header_h)
  }

  # TOP ROW (título extra)
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
        y        = y_top0 + (top_h * 0.2),
        hjust    = 0.5,
        vjust    = 0,
        size     = size_titulo_extra,
        colour   = color_barra_extra_int,
        fontface = "bold"
      )
    }
  }

  # MAIN ROW: barras (columna central)
  canvas <- canvas +
    cowplot::draw_plot(p_bars_panel, x = x_bars0, y = y_main0, width = w_bars, height = main_h)

  # Coordenadas Y “milimétricas” por fila
  y_npc <- (seq_len(n_categorias) - 0.5) / n_categorias
  y_abs <- y_main0 + y_npc * main_h

  # Etiquetas (columna izquierda)
  pad_x <- 0.012
  x_lab <- x_etq0 + w_etq * (1 - pad_x)
  fontface_etq <- if ("eje_y" %in% textos_negrita) "bold" else "plain"

  for (i in seq_len(n_categorias)) {
    canvas <- canvas + cowplot::draw_text(
      text     = etiquetas_vec[i],
      x        = x_lab,
      y        = y_abs[i],
      hjust    = 1,
      vjust    = 0.5,
      size     = size_ejes,
      colour   = color_ejes,
      fontface = fontface_etq
    )
  }

  # Extra (columna derecha)
  x_extra_txt <- x_extra0 + (w_extra * 0.5)
  for (i in seq_len(n_categorias)) {
    if (nzchar(extra_labels[i])) {
      canvas <- canvas + cowplot::draw_text(
        text     = extra_labels[i],
        x        = x_extra_txt,
        y        = y_abs[i],
        hjust    = 0.5,
        vjust    = 0.5,
        size     = size_barra_extra,
        colour   = color_barra_extra_int,
        fontface = fontface_barra_extra
      )
    }
  }

  if (debug_ph_bordes) {
    canvas <- canvas +
      .ph_border(x_etq0,   y_main0, w_etq,   main_h) +
      .ph_border(x_buf10,  y_main0, w_buf1,  main_h) +
      .ph_border(x_bars0,  y_main0, w_bars,  main_h) +
      .ph_border(x_buf20,  y_main0, w_buf2,  main_h) +
      .ph_border(x_extra0, y_main0, w_extra, main_h)
  }

  # ============================================================
  # LEYENDA: centrada + desplazamiento
  # ============================================================
  if (has_legend && !is.null(leg_grob)) {

    pos_leyenda_x <- 0.5
    if (!is.na(centro_cowplot) && is.finite(centro_cowplot)) pos_leyenda_x <- centro_cowplot

    y_legend_center <- y_legend0 + (legend_h * 0.5)
    dy_leg <- leyenda_desplazamiento_in / h_total_in

    leg_w_npc <- grid::convertWidth(sum(leg_grob$widths), "npc", valueOnly = TRUE)
    if (!is.finite(leg_w_npc) || leg_w_npc <= 0) leg_w_npc <- 1

    canvas <- canvas + cowplot::draw_grob(
      leg_grob,
      x = pos_leyenda_x,
      y = y_legend_center + dy_leg,
      width  = leg_w_npc,
      height = legend_h,
      hjust  = 0.5,
      vjust  = 0.5
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

  # ---------------------------------------------------------------------------
  # 10) EXPORT
  # ---------------------------------------------------------------------------
  if (exportar == "rplot") {
    attr(canvas, "alto_word_sugerido") <- h_total_in
    return(canvas)
  }

  if (is.null(path_salida) || !nzchar(path_salida)) stop("`path_salida` es requerido para exportar.", call. = FALSE)

  if (exportar == "png") {
    ggplot2::ggsave(filename = path_salida, plot = canvas, width = ancho, height = alto, units = "in", dpi = dpi, bg = "transparent")
    return(invisible(canvas))
  }

  if (exportar %in% c("ppt", "word")) {
    if (!requireNamespace("officer", quietly = TRUE)) stop("Para exportar a PPT/Word se requiere officer.", call. = FALSE)
    if (!requireNamespace("rvg", quietly = TRUE))     stop("Para exportar a PPT/Word se requiere rvg (dml).", call. = FALSE)

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

    if (exportar == "word") {
      doc <- if (file.exists(path_salida)) officer::read_docx(path_salida) else officer::read_docx()
      doc <- officer::body_add_par(doc, value = "", style = "Normal")
      doc <- officer::body_add_dml(
        doc,
        value = rvg::dml(ggobj = canvas),
        width = ancho, height = alto
      )
      print(doc, target = path_salida)
      return(invisible(canvas))
    }
  }

  stop("Tipo de exportación no soportado.", call. = FALSE)
}
