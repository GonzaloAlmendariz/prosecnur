# =============================================================================
# Explorador interactivo: reporte_interactivo()
# - Helpers compartidos
# - Función exportada
# - Wiring UI/Server (delegación a tabs)
# =============================================================================

`%||%` <- function(x, y) if (!is.null(x)) x else y

# -----------------------------------------------------------------------------
# Helpers internos (IGUAL que tu archivo)
# -----------------------------------------------------------------------------

.wrap_y <- function(x, width = 35) {
  x <- as.character(x)
  if (requireNamespace("stringr", quietly = TRUE)) {
    x <- stringr::str_wrap(x, width = width)
  }
  gsub("\n", "<br>", x, fixed = TRUE)  # plotly interpreta <br>
}

.resolver_paleta_var <- function(var,
                                 instrumento,
                                 colores_apiladas_por_listname,
                                 opcion_levels) {

  surv <- instrumento$survey
  pal  <- NULL

  if (!is.null(colores_apiladas_por_listname) &&
      !is.null(surv) &&
      all(c("name", "list_name") %in% names(surv))) {

    ln <- surv$list_name[surv$name == var][1]
    if (!is.na(ln) && ln %in% names(colores_apiladas_por_listname)) {
      pal <- colores_apiladas_por_listname[[ln]]
    }
  }

  if (is.null(pal) || !length(pal)) {
    out <- grDevices::hcl.colors(max(3L, length(opcion_levels)), "Blues")
    out <- out[seq_len(length(opcion_levels))]
    names(out) <- opcion_levels
    return(out)
  }

  if (!is.null(names(pal)) && all(opcion_levels %in% names(pal))) {
    pal2 <- pal[opcion_levels]
    names(pal2) <- opcion_levels
    return(pal2)
  }

  fila <- surv[surv$name == var, , drop = FALSE]
  list_var <- if (nrow(fila)) fila$list_name[1] else NA_character_

  if (!is.null(instrumento$choices) &&
      all(c("list_name", "name", "label") %in% names(instrumento$choices)) &&
      !is.na(list_var) && nzchar(list_var) &&
      !is.null(names(pal))) {

    ch <- instrumento$choices[instrumento$choices$list_name == list_var, , drop = FALSE]
    map_code_to_label <- stats::setNames(as.character(ch$label), as.character(ch$name))

    idx <- names(pal) %in% names(map_code_to_label)
    if (any(idx)) {
      pal_lab <- stats::setNames(pal[idx], map_code_to_label[names(pal)[idx]])

      if (!all(opcion_levels %in% names(pal_lab))) {
        falt <- setdiff(opcion_levels, names(pal_lab))
        extra <- grDevices::hcl.colors(max(3L, length(falt)), "Blues")
        extra <- extra[seq_len(length(falt))]
        pal_lab <- c(pal_lab, stats::setNames(extra, falt))
      }

      pal_lab <- pal_lab[opcion_levels]
      names(pal_lab) <- opcion_levels
      return(pal_lab)
    }
  }

  pal <- rep(pal, length.out = length(opcion_levels))
  names(pal) <- opcion_levels
  pal
}

.obtener_label_var <- function(var, instrumento, data = NULL) {
  survey <- instrumento$survey

  if (!is.null(survey) &&
      all(c("name", "label") %in% names(survey)) &&
      var %in% survey$name) {

    lab <- survey$label[survey$name == var][1]
    if (!is.na(lab) && nzchar(as.character(lab))) return(as.character(lab))
  }

  if (!is.null(data) && var %in% names(data)) {
    vl <- attr(data[[var]], "label", exact = TRUE)
    if (!is.null(vl) && nzchar(as.character(vl))) return(as.character(vl))
  }

  as.character(var)
}

.wrap_titulo_html <- function(txt, width = 120) {
  if (!requireNamespace("stringr", quietly = TRUE)) return(txt)
  txt <- as.character(txt)
  if (!nzchar(txt)) return(txt)
  lineas <- stringr::str_wrap(txt, width = width)
  paste(lineas, collapse = "<br>")
}

.anotar_porcentajes_enteros <- function(df_tab) {
  if (!requireNamespace("dplyr", quietly = TRUE)) {
    stop("Se requiere 'dplyr' para .anotar_porcentajes_enteros().", call. = FALSE)
  }

  df_tab$pct[is.na(df_tab$pct)] <- 0
  df_tab$pct[df_tab$pct < 0]    <- 0

  df_split <- split(df_tab, df_tab$estrato_label, drop = FALSE)

  df_list <- lapply(df_split, function(df_g) {
    total <- sum(df_g$pct, na.rm = TRUE)

    if (is.na(total) || total <= 0) {
      df_g$porc_raw <- 0
      df_g$porc_int <- 0L
      return(df_g)
    }

    pct_norm <- df_g$pct / total

    raw  <- pct_norm * 100
    base <- floor(raw + 1e-9)
    frac <- raw - base

    suma_base <- sum(base)
    rem       <- as.integer(round(100 - suma_base))

    if (rem > 0) {
      ord <- order(frac, decreasing = TRUE, na.last = NA)
      k   <- min(rem, length(ord))
      if (k > 0) base[ord[seq_len(k)]] <- base[ord[seq_len(k)]] + 1L
    } else if (rem < 0) {
      ord <- order(frac, decreasing = FALSE, na.last = NA)
      k   <- min(-rem, length(ord))
      if (k > 0) base[ord[seq_len(k)]] <- pmax(0L, base[ord[seq_len(k)]] - 1L)
    }

    df_g$porc_raw <- pct_norm
    df_g$porc_int <- base
    df_g
  })

  dplyr::bind_rows(df_list)
}

.preparar_tabla_proporciones <- function(data,
                                         instrumento,
                                         var,
                                         var_cruce = NULL,
                                         codigos_perdidos = NULL) {

  if (!requireNamespace("dplyr", quietly = TRUE)) {
    stop("Se requiere 'dplyr' para `reporte_interactivo()`.", call. = FALSE)
  }

  survey  <- instrumento$survey
  choices <- instrumento$choices %||% NULL

  if (is.null(survey) || !"name" %in% names(survey)) {
    stop("El `instrumento` debe contener `survey` válido.", call. = FALSE)
  }

  fila_var <- survey[survey$name == var, , drop = FALSE]
  if (nrow(fila_var) == 0L) {
    stop("La variable '", var, "' no está en `instrumento$survey`.", call. = FALSE)
  }
  list_main <- fila_var$list_name[1]

  if (!is.null(choices) &&
      all(c("list_name", "name", "label") %in% names(choices)) &&
      !is.na(list_main) && nzchar(list_main)) {

    ch_main      <- choices[choices$list_name == list_main, , drop = FALSE]
    codigos_main <- as.character(ch_main$name)
    labels_main  <- as.character(ch_main$label)
  } else {
    codigos_main <- sort(unique(as.character(data[[var]])))
    labels_main  <- codigos_main
  }

  map_main <- stats::setNames(labels_main, codigos_main)

  # ✅ FIX 1: definir SIEMPRE el orden (para cruce y no-cruce)
  orden_lvls_main <- map_main[codigos_main]

  df <- data
  if (!var %in% names(df)) stop("La variable '", var, "' no existe en `data`.", call. = FALSE)

  df[[var]] <- as.character(df[[var]])
  df <- df[!is.na(df[[var]]), , drop = FALSE]

  if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0) {
    df <- df[!(df[[var]] %in% as.character(codigos_perdidos)), , drop = FALSE]
  }

  if (nrow(df) == 0L) stop("No hay datos válidos para '", var, "'.", call. = FALSE)

  # ---------------------- SIN CRUCE ----------------------
  if (is.null(var_cruce) || !nzchar(var_cruce)) {

    df_tab <- df |>
      dplyr::count(.data[[var]], name = "n") |>
      dplyr::mutate(
        pct           = n / sum(n),
        opcion_code   = as.character(.data[[var]]),
        opcion_label  = map_main[opcion_code] %||% opcion_code,
        estrato_label = ""
      ) |>
      dplyr::select(estrato_label, opcion_label, pct, n)

    # ✅ FIX 2: usar orden_lvls_main (y no el typo orden_lvls_main inexistente)
    df_tab$opcion_label <- factor(
      df_tab$opcion_label,
      levels = unique(orden_lvls_main[!is.na(orden_lvls_main)])
    )

    df_tab <- df_tab[order(df_tab$opcion_label), , drop = FALSE]
    return(df_tab)
  }

  # ---------------------- CON CRUCE ----------------------
  if (!var_cruce %in% names(df)) stop("Cruce '", var_cruce, "' no existe en `data`.", call. = FALSE)

  df[[var_cruce]] <- as.character(df[[var_cruce]])

  fila_cruce <- survey[survey$name == var_cruce, , drop = FALSE]
  list_cruce <- if (nrow(fila_cruce)) fila_cruce$list_name[1] else NA_character_

  if (!is.null(choices) &&
      all(c("list_name", "name", "label") %in% names(choices)) &&
      !is.na(list_cruce) && nzchar(list_cruce)) {

    ch_cruce  <- choices[choices$list_name == list_cruce, , drop = FALSE]
    map_cruce <- stats::setNames(as.character(ch_cruce$label), as.character(ch_cruce$name))
  } else {
    niveles_cruce <- sort(unique(df[[var_cruce]]))
    map_cruce     <- stats::setNames(niveles_cruce, niveles_cruce)
  }

  df_tab <- df |>
    dplyr::count(.data[[var_cruce]], .data[[var]], name = "n") |>
    dplyr::group_by(.data[[var_cruce]]) |>
    dplyr::mutate(pct = n / sum(n)) |>
    dplyr::ungroup() |>
    dplyr::mutate(
      opcion_code   = as.character(.data[[var]]),
      opcion_label  = map_main[opcion_code] %||% opcion_code,
      estrato_code  = as.character(.data[[var_cruce]]),
      estrato_label = map_cruce[estrato_code] %||% estrato_code
    ) |>
    dplyr::select(estrato_label, opcion_label, pct, n)

  df_tab$opcion_label  <- factor(df_tab$opcion_label,
                                 levels = unique(orden_lvls_main[!is.na(orden_lvls_main)]))
  df_tab$estrato_label <- factor(df_tab$estrato_label,
                                 levels = sort(unique(df_tab$estrato_label)))

  if (length(unique(df_tab$estrato_label)) == 1 &&
      unique(as.character(df_tab$estrato_label)) %in% c("Total", "TOTAL", "total")) {
    df_tab$estrato_label <- factor(rep("", nrow(df_tab)))
  }

  df_tab[order(df_tab$estrato_label, df_tab$opcion_label), , drop = FALSE]
}

.construir_tabla_resumen <- function(df_tab) {
  if (!requireNamespace("dplyr", quietly = TRUE)) {
    stop("Se requiere 'dplyr' para la tabla resumen.", call. = FALSE)
  }

  df_tab <- .anotar_porcentajes_enteros(df_tab)

  if (all(as.character(df_tab$estrato_label) %in% c("", NA))) {
    df_tab |>
      dplyr::arrange(opcion_label) |>
      dplyr::transmute(
        Respuesta  = as.character(.data$opcion_label),
        N          = .data$n,
        Porcentaje = paste0(.data$porc_int, "%")
      )
  } else {
    df_tab |>
      dplyr::arrange(estrato_label, opcion_label) |>
      dplyr::transmute(
        Estrato    = as.character(.data$estrato_label),
        Respuesta  = as.character(.data$opcion_label),
        N          = .data$n,
        Porcentaje = paste0(.data$porc_int, "%")
      )
  }
}

.construir_plotly_barras <- function(df_tab,
                                     titulo,
                                     var_paleta = NULL,
                                     instrumento = NULL,
                                     colores_apiladas_por_listname = NULL,
                                     paleta_colores = NULL,
                                     height = NULL,
                                     mostrar_leyenda = TRUE) {

  if (!requireNamespace("plotly", quietly = TRUE)) {
    stop("Se requiere 'plotly' para `reporte_interactivo()`.", call. = FALSE)
  }

  df_tab$pct[is.na(df_tab$pct)] <- 0
  df_tab$pct[df_tab$pct < 0]    <- 0
  df_tab$n[is.na(df_tab$n)]     <- 0

  df_tab <- .anotar_porcentajes_enteros(df_tab)

  df_tab$texto_pct      <- paste0(df_tab$porc_int, "%")
  df_tab$texto_pct_html <- paste0("<b>", df_tab$porc_int, "%</b>")

  opcion_levels  <- levels(df_tab$opcion_label) %||% unique(df_tab$opcion_label)
  estrato_levels <- levels(df_tab$estrato_label) %||% unique(df_tab$estrato_label)

  df_tab$opcion_label  <- factor(df_tab$opcion_label,  levels = opcion_levels)
  df_tab$estrato_label <- factor(df_tab$estrato_label, levels = estrato_levels)

  solo_total <- all(as.character(df_tab$estrato_label) %in% c("", NA))

  if (is.null(paleta_colores) || !length(paleta_colores)) {
    if (!is.null(var_paleta) && !is.null(instrumento)) {
      paleta_colores <- .resolver_paleta_var(
        var = var_paleta,
        instrumento = instrumento,
        colores_apiladas_por_listname = colores_apiladas_por_listname,
        opcion_levels = as.character(opcion_levels)
      )
    } else {
      paleta_colores <- grDevices::hcl.colors(max(3L, length(opcion_levels)), "Blues")
      paleta_colores <- paleta_colores[seq_len(length(opcion_levels))]
      names(paleta_colores) <- as.character(opcion_levels)
    }
  } else {
    if (is.null(names(paleta_colores))) {
      paleta_colores <- rep(paleta_colores, length.out = length(opcion_levels))
      names(paleta_colores) <- as.character(opcion_levels)
    } else if (!all(as.character(opcion_levels) %in% names(paleta_colores))) {
      falt <- setdiff(as.character(opcion_levels), names(paleta_colores))
      extra <- grDevices::hcl.colors(max(3L, length(falt)), "Blues")
      extra <- extra[seq_len(length(falt))]
      paleta_colores <- c(paleta_colores, stats::setNames(extra, falt))
    }
    paleta_colores <- paleta_colores[as.character(opcion_levels)]
    names(paleta_colores) <- as.character(opcion_levels)
  }

  n_estratos <- length(unique(df_tab$estrato_label))
  if (is.null(height)) height <- max(220, min(650, 160 + 60 * n_estratos))

  if (mostrar_leyenda) {
    titulo_margin_top <- 60
    margin_left       <- if (solo_total) 20 else 170
    margin_right      <- 25
    margin_bottom     <- 45
  } else {
    titulo_margin_top <- 35
    margin_left       <- if (solo_total) 20 else 120
    margin_right      <- 10
    margin_bottom     <- 25
  }

  p <- plotly::plot_ly(height = height)

  for (opt in as.character(opcion_levels)) {
    df_opt <- df_tab[df_tab$opcion_label == opt, , drop = FALSE]
    if (!nrow(df_opt)) next

    if (solo_total) {
      df_opt$hover_text <- sprintf("%s: %s<br>N: %s", opt, df_opt$texto_pct, df_opt$n)
    } else {
      df_opt$hover_text <- sprintf(
        "%s<br>%s: %s<br>N: %s",
        as.character(df_opt$estrato_label),
        opt,
        df_opt$texto_pct,
        df_opt$n
      )
    }

    df_opt$texto_in  <- paste0("<b>", df_opt$texto_pct, "</b>")

    p <- p |>
      plotly::add_bars(
        data             = df_opt,
        x                = ~pct,
        y                = ~estrato_label,
        name             = opt,
        orientation      = "h",
        text             = ~texto_in,
        textposition     = "inside",
        insidetextanchor = "middle",
        textfont         = list(color = "white", size = 11),
        customdata       = ~hover_text,
        hovertemplate    = "%{customdata}<extra></extra>",
        marker           = list(
          color = unname(paleta_colores[opt]),
          line  = list(width = 0)
        )
      )
  }

  p <- p |>
    plotly::layout(
      barmode = "stack",
      bargap  = 0.25,
      xaxis   = list(
        title          = "",
        range          = c(0, 1),
        showgrid       = FALSE,
        zeroline       = FALSE,
        showticklabels = FALSE,
        ticks          = ""
      ),
      yaxis   = list(
        title          = "",
        automargin     = !solo_total,
        showticklabels = !solo_total,
        showgrid       = FALSE,
        zeroline       = FALSE,
        ticks          = ""
      ),
      legend = list(
        orientation = "h",
        x = 0.5, xanchor = "center",
        y = -0.12
      ),
      margin = list(l = margin_left, r = margin_right, t = titulo_margin_top, b = margin_bottom),
      uniformtext = list(minsize = 10, mode = "hide"),
      hovermode  = "closest",
      showlegend = mostrar_leyenda,
      transition = list(duration = 450, easing = "cubic-in-out")
    ) |>
    plotly::config(displayModeBar = FALSE, responsive = TRUE)

  plotly::animation_opts(
    p,
    frame      = 600,
    transition = 450,
    easing     = "cubic-in-out",
    redraw     = TRUE
  )
}

.construir_kpi_halfdonut <- function(df,
                                     var_kpi,
                                     instrumento,
                                     colores_apiladas_por_listname,
                                     codigos_perdidos = NULL) {

  if (!requireNamespace("plotly", quietly = TRUE)) return(NULL)
  if (!var_kpi %in% names(df)) return(NULL)

  df_kpi <- df[!is.na(df[[var_kpi]]), , drop = FALSE]
  if (!nrow(df_kpi)) return(NULL)

  df_tab <- .preparar_tabla_proporciones(
    data             = df_kpi,
    instrumento      = instrumento,
    var              = var_kpi,
    var_cruce        = NULL,
    codigos_perdidos = codigos_perdidos
  )
  df_tab <- .anotar_porcentajes_enteros(df_tab)

  df_tab$opcion_label <- as.character(df_tab$opcion_label)
  df_tab <- df_tab[order(df_tab$opcion_label), , drop = FALSE]

  titulo_kpi <- .wrap_titulo_html(
    .obtener_label_var(var_kpi, instrumento, df_kpi),
    width = 45
  )

  opcion_levels <- as.character(df_tab$opcion_label)
  paleta <- .resolver_paleta_var(
    var = var_kpi,
    instrumento = instrumento,
    colores_apiladas_por_listname = colores_apiladas_por_listname,
    opcion_levels = opcion_levels
  )

  legend_df <- data.frame(
    label = opcion_levels,
    color = unname(paleta[opcion_levels]),
    stringsAsFactors = FALSE
  )

  p <- plotly::plot_ly(
    data   = df_tab,
    labels = ~opcion_label,
    values = ~porc_int,
    type   = "pie",
    hole   = 0.68,
    direction = "clockwise",
    rotation  = 180,
    sort      = FALSE,
    textinfo  = "none",
    marker    = list(colors = unname(paleta[as.character(df_tab$opcion_label)])),
    hovertemplate = "%{label}: %{value}%<extra></extra>"
  ) |>
    plotly::layout(
      title = NULL,
      showlegend = FALSE,
      margin = list(l = 10, r = 10, t = 10, b = 5),
      annotations = list(),
      transition = list(duration = 450, easing = "cubic-in-out")
    ) |>
    plotly::animation_opts(
      frame      = 600,
      transition = 450,
      easing     = "cubic-in-out",
      redraw     = TRUE
    ) |>
    plotly::config(displayModeBar = FALSE, responsive = TRUE)

  list(plot = p, legend = legend_df, title_html = titulo_kpi)
}

.resolver_var_spec <- function(var_sel, ctx){

  # ctx debe tener: data, instrumento, sm_cols_map, label_var
  data <- ctx$data
  inst <- ctx$instrumento

  # 1) Si existe como columna -> SO (o una recod ya materializada)
  if (var_sel %in% names(data)) {
    return(list(
      tipo      = "so",
      var_madre = var_sel,
      cols      = var_sel
    ))
  }

  # 2) Si NO existe pero es madre SM -> usar mapa de dummies
  if (!is.null(ctx$sm_cols_map) && var_sel %in% names(ctx$sm_cols_map)) {

    cols <- ctx$sm_cols_map[[var_sel]]
    cols <- cols[cols %in% names(data)]
    if (!length(cols)) stop("SM madre sin dummies encontradas: ", var_sel, call. = FALSE)

    # construir opciones (labels) desde instrumento$choices usando list_name de survey
    surv <- inst$survey
    choices <- inst$choices

    fila <- surv[surv$name == var_sel, , drop = FALSE]
    ln   <- if (nrow(fila)) as.character(fila$list_name[1]) else NA_character_

    if (!is.na(ln) && nzchar(ln) &&
        !is.null(choices) && all(c("list_name","name","label") %in% names(choices))) {

      ch <- choices[choices$list_name == ln, , drop = FALSE]
      # ojo: tus dummies suelen terminar en ".<code>"
      # entonces se usa `name` como code (p.ej. "70") y label como opción
      map_code_to_label <- stats::setNames(as.character(ch$label), as.character(ch$name))

    } else {
      map_code_to_label <- NULL
    }

    return(list(
      tipo      = "sm",
      var_madre = var_sel,
      cols      = cols,
      map_code_to_label = map_code_to_label
    ))
  }

  stop("No se pudo resolver variable seleccionada: ", var_sel, call. = FALSE)
}

.plot_dummy_yesno <- function(x, label_opcion) {

  x <- x[!is.na(x)]
  if (!length(x)) {
    return(plotly::plot_ly(height = BAR_HEIGHT) |>
             plotly::layout(
               annotations = list(list(text="Sin datos", showarrow=FALSE)),
               margin = list(l=10, r=10, t=0, b=0)
             ) |>
             plotly::config(displayModeBar = FALSE)
    )
  }

  n_si <- sum(x == 1)
  n_no <- sum(x == 0)
  tot  <- n_si + n_no

  tab <- data.frame(
    resp = c("Sí","No"),
    pct  = c(n_si, n_no) / tot
  )

  p <- plotly::plot_ly(height = BAR_HEIGHT) |>
    plotly::add_bars(
      data = tab,
      x = ~pct,
      y = I("Total"),
      orientation = "h",
      marker = list(
        color = c("#1B679D", "#E5ECF6"),
        line = list(width = 0)
      ),
      text = paste0("<b>", round(100 * tab$pct, 0), "%</b>"),
      textposition = "inside",
      insidetextanchor = "middle",
      textfont = list(color="white", size=PCT_FSIZE),
      hoverinfo = "skip"
    ) |>
    plotly::layout(
      barmode = "stack",
      xaxis = list(range=c(0,1), visible=FALSE),
      yaxis = list(visible=FALSE),
      margin = list(l=10, r=10, t=0, b=0),
      showlegend = FALSE
    ) |>
    plotly::config(displayModeBar = FALSE)

  p
}


# =============================================================================
# resolver_var_spec()
# -----------------------------------------------------------------------------
# Helper para variables select_multiple "madre" que en la data viven como dummies
# (p.ej. p106_recod.1, p106_recod.2, ...) y/o como compacta (_ORIG o madre).
#
# Retorna:
# - cols: vector de dummies disponibles (ordenadas)
# - map_code_to_label: named list (code -> label) para subtítulos UI
# - var_madre: nombre de la madre (tal cual se pidió)
# - list_name: list_name del XLSForm si existe
# - col_compact: columna compacta detectada si existe (p.ej. p106_recod_ORIG)
# =============================================================================
resolver_var_spec <- function(var_madre, ctx, df = NULL) {

  `%||%` <- get0("%||%", ifnotfound = function(x, y) if (!is.null(x)) x else y)

  # data: priorizar df si viene (p.ej. data_filtrada()), si no usar ctx$data
  data <- df %||% ctx$data
  inst <- ctx$instrumento

  if (is.null(data) || !is.data.frame(data)) {
    return(list(
      var_madre = var_madre,
      cols = character(0),
      map_code_to_label = list(),
      list_name = NA_character_,
      col_compact = NA_character_
    ))
  }

  # ------------------------------------------------------------
  # 1) Detectar dummies disponibles
  # ------------------------------------------------------------
  var_esc <- gsub("([\\W])", "\\\\\\1", var_madre)
  pat_dum <- paste0("^", var_esc, "\\.")
  cols <- grep(pat_dum, names(data), value = TRUE)

  # ------------------------------------------------------------
  # 2) Detectar columna compacta (madre o _ORIG) si existe
  # ------------------------------------------------------------
  col_compact <- NA_character_
  cand1 <- paste0(var_madre, "_ORIG")
  if (cand1 %in% names(data)) {
    col_compact <- cand1
  } else if (var_madre %in% names(data)) {
    col_compact <- var_madre
  }

  # ------------------------------------------------------------
  # 3) Obtener list_name y diccionario code->label desde inst
  # ------------------------------------------------------------
  surv <- inst$survey %||% NULL
  ch   <- inst$choices %||% NULL

  list_name <- NA_character_
  if (!is.null(surv) && all(c("name","list_name") %in% names(surv)) && var_madre %in% surv$name) {
    list_name <- as.character(surv$list_name[surv$name == var_madre][1])
    if (is.na(list_name) || !nzchar(list_name)) list_name <- NA_character_
  }

  map_code_to_label <- NULL

  # 3a) preferir choices del instrumento
  if (!is.null(ch) && all(c("list_name","name") %in% names(ch))) {
    # label puede ser "label" o "label::Spanish (ES)" etc.
    label_col <- NULL
    if ("label" %in% names(ch)) {
      label_col <- "label"
    } else {
      lab_candidates <- grep("^label(::|$)", names(ch), value = TRUE)
      if (length(lab_candidates)) label_col <- lab_candidates[1]
    }

    if (!is.na(list_name) && nzchar(list_name) && !is.null(label_col) && label_col %in% names(ch)) {
      ch_v <- ch[ch$list_name == list_name, , drop = FALSE]
      if (nrow(ch_v)) {
        map_code_to_label <- stats::setNames(as.character(ch_v[[label_col]]), as.character(ch_v$name))
      }
    }
  }

  # 3b) fallback: labels del atributo si existe alguna dummy con labels
  if (is.null(map_code_to_label)) {
    # buscar en madre si existe, si no en primera dummy
    cand_attr <- NULL
    if (!is.na(col_compact) && col_compact %in% names(data)) cand_attr <- col_compact
    if (is.null(cand_attr) && length(cols)) cand_attr <- cols[1]

    if (!is.null(cand_attr) && cand_attr %in% names(data)) {
      labs <- attr(data[[cand_attr]], "labels", exact = TRUE)
      if (!is.null(labs) && length(labs) > 0) {
        map_code_to_label <- stats::setNames(as.character(unname(labs)), as.character(names(labs)))
      }
    }
  }

  if (is.null(map_code_to_label)) map_code_to_label <- character(0)

  # ------------------------------------------------------------
  # 4) Orden de opciones (codes) y reordenamiento de dummies
  # ------------------------------------------------------------
  # helper: extraer code desde dummy "p106_recod.70" -> "70"
  dummy_code <- function(x) sub(paste0("^", var_madre, "\\."), "", x)

  dummy_codes <- if (length(cols)) dummy_code(cols) else character(0)

  # si hay compacta, intentar ordenar por aparición/choices
  codes_order <- character(0)

  # 4a) si existe diccionario, usar su orden
  if (length(map_code_to_label) > 0) {
    codes_order <- as.character(names(map_code_to_label))
  }

  # 4b) si no, pero hay compacta, derivar codes existentes en data (split ;)
  if (!length(codes_order) && !is.na(col_compact) && col_compact %in% names(data)) {
    x <- as.character(data[[col_compact]])
    x <- x[!is.na(x) & nzchar(x) & x != "NA"]
    if (length(x)) {
      vals <- unlist(strsplit(x, "\\s*;\\s*"), use.names = FALSE)
      vals <- trimws(vals)
      vals <- vals[!is.na(vals) & nzchar(vals) & vals != "NA"]
      codes_order <- unique(vals)
    }
  }

  # 4c) si no, usar codes de dummies
  if (!length(codes_order) && length(dummy_codes)) {
    codes_order <- unique(dummy_codes)
  }

  # ordenar con heurística numérica si aplica
  if (length(codes_order)) {
    suppressWarnings({
      num <- as.numeric(codes_order)
    })
    if (!all(is.na(num))) {
      # mezcla numérica: ordenar numéricos primero, luego alfanuméricos
      ord <- order(is.na(num), num, codes_order)
      codes_order <- codes_order[ord]
    } else {
      codes_order <- sort(codes_order)
    }
  }

  # Reordenar cols según codes_order
  if (length(cols) && length(codes_order)) {
    ord_idx <- match(dummy_codes, codes_order)
    # los no encontrados al final
    ord_idx[is.na(ord_idx)] <- max(ord_idx, na.rm = TRUE) + seq_len(sum(is.na(ord_idx)))
    cols <- cols[order(ord_idx)]
  }

  # ------------------------------------------------------------
  # 5) Asegurar que map_code_to_label cubra todos los codes visibles
  # ------------------------------------------------------------
  if (length(dummy_codes)) {
    falt <- setdiff(dummy_codes, names(map_code_to_label))
    if (length(falt)) {
      extra <- stats::setNames(falt, falt)
      map_code_to_label <- c(map_code_to_label, extra)
    }
  }

  # devolver como LIST para acceso [[code]] fácil
  map_list <- as.list(map_code_to_label)

  list(
    var_madre = var_madre,
    cols = cols,
    map_code_to_label = map_list,
    list_name = list_name,
    col_compact = col_compact
  )
}




# -----------------------------------------------------------------------------
# Registry de pestañas
# -----------------------------------------------------------------------------

.make_tabs_registry <- function(ctx, tabs = c("resumen", "relacion", "base_datos")) {

  registry <- list(

    resumen = list(
      ui = function(ctx) shiny::tabPanel(title = "Resumen", .ui_tab_resumen(ctx)),
      server = function(ctx, input, output, session) .server_tab_resumen(ctx, input, output, session)
    ),

    relacion = list(
      ui = function(ctx) shiny::tabPanel(title = "Relación", relacion_tab_ui("relacion")),
      server = function(ctx, input, output, session) {
        relacion_tab_server(
          id          = "relacion",
          data        = ctx$data,
          instrumento = ctx$instrumento,
          secciones   = ctx$secciones_limpias,

          vars_so       = ctx$so_vars %||% character(0),
          vars_sm_madres = ctx$sm_madres %||% character(0),

          colores_apiladas_por_listname = ctx$colores_apiladas_por_listname,
          codigos_perdidos = ctx$codigos_perdidos,
          weight_col = "peso",
          orders_list = ctx$instrumento$orders_list %||% NULL,
          labels_override = NULL
        )
      }
    ),

    base_datos = list(
      ui = function(ctx) shiny::tabPanel(title = "Base de datos", .ui_tab_base_datos(ctx)),
      server = function(ctx, input, output, session) .server_tab_base_datos(ctx, input, output, session)
    )
  )

  tabs <- (tabs %||% c("resumen", "base_datos"))
  tabs <- tabs[tabs %in% names(registry)]
  if (!length(tabs)) stop("`tabs` no contiene pestañas válidas.", call. = FALSE)

  registry[tabs]
}

# -----------------------------------------------------------------------------
# Función exportada
# -----------------------------------------------------------------------------

#' Explorador interactivo de resultados (pestañas parametrizables)
#' @export
#' @importFrom stats setNames
#' @importFrom dplyr n_distinct
reporte_interactivo <- function(
    data,
    instrumento,
    secciones,
    fuente      = NULL,
    titulo      = "Explorador interactivo",
    colores_apiladas_por_listname = NULL,
    codigos_perdidos = NULL,
    facet_vars = NULL,
    id_unidad  = NULL,
    kpi_vars   = NULL,
    logo_png   = NULL,
    logo_alt   = "Logo",
    logo_height_px = 52,
    tabs = c("resumen", "relacion", "base_datos")
) {

  if (!requireNamespace("shiny", quietly = TRUE) ||
      !requireNamespace("plotly", quietly = TRUE) ||
      !requireNamespace("dplyr",  quietly = TRUE) ||
      !requireNamespace("DT",     quietly = TRUE)) {
    stop("Se requieren 'shiny', 'plotly', 'dplyr' y 'DT' para `reporte_interactivo()`.", call. = FALSE)
  }

  # --- mantener EXACTO lo que haces hoy ---
  tiene_labels <- any(vapply(names(data), function(v) {
    !is.null(attr(data[[v]], "label",  exact = TRUE)) ||
      !is.null(attr(data[[v]], "labels", exact = TRUE)) ||
      !is.null(attr(data[[v]], "measure", exact = TRUE))
  }, logical(1)))

  if (!inherits(data, "prosecnur_reporte_tbl") || !tiene_labels) {
    data <- reporte_data(
      data        = data,
      instrumento = instrumento
    )
  }

  survey <- instrumento$survey
  if (is.null(survey) || !"name" %in% names(survey)) {
    stop("El `instrumento` debe contener un `survey` válido.", call. = FALSE)
  }

  if (is.null(secciones) || !length(secciones)) {
    stop("`secciones` debe ser una lista nombrada con vectores de variables.", call. = FALSE)
  }

  .is_tecnica <- function(v, instrumento) {
    if (!nzchar(v)) return(TRUE)
    if (startsWith(v, "_")) return(TRUE)

    vf <- as.character(attr(data, "vars_fecha", exact = TRUE) %||% instrumento$vars_fecha %||% character(0))
    vh <- as.character(attr(data, "vars_hora", exact = TRUE) %||% instrumento$vars_hora %||% character(0))
    vd <- as.character(attr(data, "vars_datetime", exact = TRUE) %||% instrumento$vars_datetime %||% character(0))
    if (v %in% c(vf, vh, vd)) return(TRUE)

    if (!is.null(instrumento$survey) && all(c("name","type") %in% names(instrumento$survey))) {
      fila <- instrumento$survey[instrumento$survey$name == v, , drop = FALSE]
      if (nrow(fila)) {
        tp <- tolower(as.character(fila$type[1]))
        if (tp %in% c("start","end","deviceid","subscriberid","simserial","phonenumber","today","username","audit")) {
          return(TRUE)
        }
      }
    }

    FALSE
  }

  label_var <- function(v) .obtener_label_var(v, instrumento, data)

  # visibles
  vars_data_visibles <- setdiff(
    names(data),
    names(data)[vapply(names(data), .is_tecnica, logical(1), instrumento = instrumento)]
  )

  # diccionario conceptual
  so_inst <- survey$name[grepl("^select_one\\b",  tolower(survey$type))]
  sm_inst <- survey$name[grepl("^select_multiple\\b", tolower(survey$type))]

  so_vars <- intersect(so_inst, vars_data_visibles)

  sm_disponibles <- sm_inst[vapply(sm_inst, function(v) {
    patt <- paste0("^", v, "(\\.|_recod\\.|_otro$)")
    any(grepl(patt, vars_data_visibles))
  }, logical(1))]

  vars_diccionario_all <- sort(unique(c(so_vars, sm_disponibles)))

  sm_cols_map <- stats::setNames(vector("list", length(sm_disponibles)), sm_disponibles)
  for (v in sm_disponibles) {
    patt <- paste0("^", v, "(\\.|_recod\\.|_otro$)")
    sm_cols_map[[v]] <- grep(patt, vars_data_visibles, value = TRUE)
  }

  .to_labels_df <- function(df) {
    out <- df
    for (v in names(out)) {
      labs <- attr(out[[v]], "labels", exact = TRUE)
      if (!is.null(labs) && length(labs) > 0) {
        codes <- names(labs)
        lbls  <- unname(labs)

        x <- out[[v]]
        x_chr <- as.character(x)

        map_code_to_label <- stats::setNames(as.character(lbls), as.character(codes))
        x_lbl <- unname(map_code_to_label[x_chr])
        x_lbl[is.na(x_lbl) & !is.na(x_chr)] <- x_chr[is.na(x_lbl) & !is.na(x_chr)]
        out[[v]] <- x_lbl
      } else {
        out[[v]] <- out[[v]]
      }
    }
    out
  }

  # KPIs
  kpi_vars <- (kpi_vars %||% character(0))
  kpi_vars <- unique(kpi_vars[kpi_vars %in% names(data)])
  if (length(kpi_vars) > 2L) kpi_vars <- kpi_vars[1:2]

  # secciones: mantener vars que existan como columna
  # O mantener SM madres si tienen hijas en sm_cols_map
  secciones_limpias <- lapply(secciones, function(vs) {

    # las que sí existen como columnas (SO típicamente)
    keep <- vs[vs %in% names(data)]

    # las que NO existen como columna, pero sí son SM madres con hijas
    falt <- setdiff(vs, keep)
    if (length(falt)) {
      falt_sm <- falt[falt %in% names(sm_cols_map)]
      falt_sm <- falt_sm[vapply(falt_sm, function(v) length(sm_cols_map[[v]]) > 0, logical(1))]
      keep <- c(keep, falt_sm)
    }

    unique(keep)
  })

  secciones_limpias <- secciones_limpias[vapply(secciones_limpias, length, integer(1)) > 0]
  if (!length(secciones_limpias)) {
    stop("Ninguna sección de `secciones` tiene variables presentes en `data`.", call. = FALSE)
  }
  secciones_nombres <- names(secciones_limpias)

  facet_vars <- (facet_vars %||% character(0))
  facet_vars <- facet_vars[facet_vars %in% names(data)]
  facet_choices <- stats::setNames(facet_vars, vapply(facet_vars, label_var, character(1)))

  logo_src <- NULL
  if (!is.null(logo_png) && nzchar(logo_png)) {
    # Si viene "www/xxx.png", normalizar a "xxx.png"
    logo_src <- sub("^www/", "", logo_png)
  }

  # ---- Contexto compartido ----
  ctx <- list(
    data = data,
    instrumento = instrumento,
    secciones_limpias = secciones_limpias,
    secciones_nombres = secciones_nombres,
    facet_choices = facet_choices,
    vars_data_visibles = vars_data_visibles,
    vars_diccionario_all = vars_diccionario_all,
    sm_cols_map = sm_cols_map,
    .to_labels_df = .to_labels_df,
    label_var = label_var,
    codigos_perdidos = codigos_perdidos,
    colores_apiladas_por_listname = colores_apiladas_por_listname,
    id_unidad = id_unidad,
    kpi_vars = kpi_vars,

    # para Relación
    so_vars   = so_vars,
    sm_madres = sm_disponibles
  )

  # registry (según vector `tabs`)
  tabs_registry <- .make_tabs_registry(ctx, tabs = tabs)

  # -------------------------------- UI ---------------------------------------
  ui <- shiny::fluidPage(

    shiny::tags$head(
      shiny::tags$style(shiny::HTML("
    /* ====== Base ====== */
    body { background: #f5f6fa; color: #1f2933; }
    .container-fluid { max-width: 1400px; }

    /* ====== Tipografía ====== */
    h2, h3, h4 { font-weight: 800; color: #002457; }
    .title { font-weight: 900; color: #002457; }

    /* ====== Sidebar ====== */
    .well, .sidebarPanel {
      background: #ffffff !important;
      border: 1px solid #e6e9f2 !important;
      border-radius: 16px !important;
      box-shadow: 0 12px 28px rgba(0, 36, 87, 0.06);
    }
    .sidebar h3 { margin-top: 0; color: #002457; }
    .sidebar p  { color: #5f6b7a; font-size: 13px; }
    .sidebar hr { border-top: 1px solid #edf0f7; }

    /* ====== Inputs ====== */
    .selectize-input, .form-control {
      border-radius: 12px !important;
      border: 1px solid #e6e9f2 !important;
      box-shadow: none !important;
      font-size: 13px;
    }
    .selectize-input.focus, .form-control:focus {
      border-color: #002457 !important;
      box-shadow: 0 0 0 3px rgba(0, 36, 87, 0.15) !important;
    }

    /* ====== Botones ====== */
    .btn {
      border-radius: 12px !important;
      border: 1px solid #e6e9f2 !important;
      background: #ffffff !important;
      font-weight: 700;
      color: #002457 !important;
    }
    .btn:hover {
      background: rgba(0, 36, 87, 0.05) !important;
      border-color: #002457 !important;
    }

    /* ====== Cards ====== */
    .cardbox {
      background: #ffffff;
      border: 1px solid #e6e9f2;
      border-radius: 18px;
      box-shadow: 0 14px 34px rgba(0, 36, 87, 0.07);
      padding: 12px;
    }

    /* ====== Layout spacing ====== */
    .row { margin-left: -10px; margin-right: -10px; }
    .col-sm-6, .col-sm-12, .col-sm-9, .col-sm-3 { padding-left: 10px; padding-right: 10px; }

    /* ====== Header con logo ====== */
    .topbar{
      background:#ffffff;
      border:1px solid #e6e9f2;
      border-radius:18px;
      box-shadow:0 14px 34px rgba(0, 36, 87, 0.07);
      padding:14px 16px;
      margin-bottom:14px;
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:14px;
    }
    .topbar-title{
      font-size:26px;
      font-weight:900;
      color:#002457;
      line-height:1.1;
      flex: 1 1 auto;
    }
    .topbar-logo{
      height:52px;
      max-width:240px;
      object-fit:contain;
      display:block;
      flex: 0 0 auto;
    }

    /* ====== Card header (editorial) ====== */
    .cardbox-header{
      padding:10px 12px 6px 12px;
      border-bottom:1px solid #edf0f7;
      margin:-12px -12px 10px -12px;
    }
    .cardbox-title{
      font-size:18px;
      font-weight:900;
      color:#002457;
      line-height:1.15;
      margin:0;
    }
    .cardbox-subtitle{
      margin-top:4px;
      font-size:12px;
      color:#5f6b7a;
    }

    /* ====== Plotly ====== */
    .plot-container, .svg-container { width: 100% !important; }
    .plotly .main-svg { overflow: visible !important; }

    /* ====== DataTable: look más “ejecutivo” ====== */
    table.dataTable { border-collapse: collapse !important; }
    table.dataTable thead th{
      background:#f1f3f9;
      color:#002457;
      font-weight:800;
      border-bottom: 1px solid #dfe5f2 !important;
      border-right: 1px solid #dfe5f2 !important;
    }
    table.dataTable tbody td{
      font-size:12px;
      color:#1f2933;
      border-bottom: 1px solid #edf0f7 !important;
      border-right: 1px solid #edf0f7 !important;
    }
    table.dataTable tbody tr:hover td{
      background: #fafbff !important;
    }

    /* ====== Toggle (switch) elegante: Códigos <-> Etiquetas ====== */
    .toggle-row{
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:10px;
      margin-top: 10px;
      margin-bottom: 10px;
    }
    .toggle-label{
      font-size: 12px;
      color: #5f6b7a;
      font-weight: 700;
      white-space: nowrap;
    }
    .switch {
      position: relative;
      display: inline-block;
      width: 52px;
      height: 28px;
      flex: 0 0 auto;
    }
    .switch input { display:none; }
    .slider {
      position: absolute;
      cursor: pointer;
      top: 0; left: 0; right: 0; bottom: 0;
      background-color: #e6e9f2;
      transition: .25s;
      border-radius: 999px;
      border: 1px solid #dfe5f2;
    }
    .slider:before {
      position: absolute;
      content: \"\";
      height: 22px;
      width: 22px;
      left: 3px;
      bottom: 2.5px;
      background-color: white;
      transition: .25s;
      border-radius: 50%;
      box-shadow: 0 6px 14px rgba(0,0,0,0.12);
    }
    input:checked + .slider {
      background-color: rgba(0, 36, 87, 0.20);
      border-color: rgba(0, 36, 87, 0.35);
    }
    input:checked + .slider:before {
      transform: translateX(23px);
    }

    /* ====== Diccionario ====== */
    .dicc-kv{
      display:grid;
      grid-template-columns: 92px 1fr;
      gap: 6px 10px;
      font-size: 12px;
      color: #1f2933;
    }
    .dicc-k{
      color: #5f6b7a;
      font-weight: 800;
    }
    .dicc-v{
      color: #1f2933;
      font-weight: 600;
      word-break: break-word;
    }

    /* DataTable fijo + wrap */
    table.dataTable { table-layout: fixed !important; width: 100% !important; }
    table.dataTable thead th, table.dataTable tbody td{
      white-space: normal !important;
      word-wrap: break-word !important;
      overflow-wrap: anywhere !important;
    }

    /* ====== KPI BLOCK (Perfil) ====== */
    .kpi-block{
      display:flex;
      flex-direction:column;
      gap:10px;
      padding-bottom: 6px;
    }

    .kpi-block-title{
      font-size:14px;
      font-weight:900;
      color:#002457;
      line-height:1.15;
      margin:0;
    }

    .kpi-block-subtitle{
      margin-top:4px;
      font-size:12px;
      color:#5f6b7a;
    }

    .kpi-n-chip{
      width:100%;
      padding:10px 12px;
      border:1px solid #edf0f7;
      border-radius:14px;
      background:#fafbff;
      display:flex;
      align-items:center;
      justify-content:center;
    }

    .kpi-n-text{
      font-size:18px;
      font-weight:900;
      color:#002457;
      letter-spacing:0.01em;
    }

    /* Donuts como “pareja” */
    .kpi-grid{
      display:flex;
      gap:12px;
      width:100%;
      align-items:stretch;
    }

    .kpi-cell{
      flex:1 1 0;
      border:1px solid #edf0f7;
      border-radius:16px;
      padding:8px 8px 10px 8px;
      background:#ffffff;
    }

    /* Leyenda más secundaria */
    .kpi-legend{
      margin-top:6px;
      display:flex;
      flex-wrap:wrap;
      gap:4px 10px;
      justify-content:center;
      font-size:10px;
      color:#5f6b7a;
      line-height:1.15;
    }

    .kpi-legend-item{
      display:inline-flex;
      align-items:center;
      gap:6px;
    }

    .kpi-legend-swatch{
      display:inline-block;
      width:10px;
      height:10px;
      border-radius:3px;
    }

    /* ====== KPI cell: evitar desbordes del título ====== */
    .kpi-cell{
      overflow: hidden;
    }

    /* plotly title dentro del KPI: wrap fuerte */
    .kpi-cell .plotly .gtitle,
    .kpi-cell .plotly .g-gtitle,
    .kpi-cell .plotly text{
      white-space: normal !important;
    }

    .kpi-cell .plotly{
      overflow: hidden !important;
    }

    /* Título encima del donut (wrap real, centrado) */
    .kpi-donut-title{
      font-size: 14px;
      font-weight: 900;
      color: #002457;
      text-align: center;
      line-height: 1.15;
      margin: 4px 6px 2px 6px;
      white-space: normal;
      overflow-wrap: anywhere;
      word-break: break-word;
    }

    /* KPI cell: layout vertical controlado */
    .kpi-cell{
      display: flex;
      flex-direction: column;
      align-items: stretch;
      justify-content: flex-start;
    }

    /* Centrar encabezados y celdas en DataTable */
    table.dataTable thead th { text-align: center !important; vertical-align: middle !important; }
    table.dataTable tbody td { text-align: center !important; vertical-align: middle !important; }

    /* ====== PERFIL (nuevo layout horizontal) ====== */
    .kpi-profile-row{
      display:flex;
      gap:12px;
      align-items:stretch;
    }

    /* Columna izquierda: N (cuasi-cuadrado) */
    .kpi-n-card{
      flex: 0 0 42%;
      min-width: 320px;
      border:1px solid #edf0f7;
      border-radius:16px;
      background:#ffffff;
      padding:12px;
      display:flex;
      flex-direction:column;
      justify-content:center;
    }

    /* Título “Perfil de la muestra” arriba del N */
    .kpi-n-card .kpi-block-title{
      margin:0 0 8px 0;
    }

    /* Chip de N más protagonista */
    .kpi-n-chip{
      padding:18px 14px;
      border-radius:16px;
    }
    .kpi-n-text{
      font-size:22px;
    }

    /* Columna derecha: dos donuts */
    .kpi-donuts{
      flex: 1 1 auto;
      display:flex;
      gap:12px;
      align-items:stretch;
    }

    /* Cada donut mantiene estética actual */
    .kpi-donuts .kpi-cell{
      flex:1 1 0;
      min-width: 260px;
    }

    /* ====== RESUMEN SECCIÓN: lista de filas ====== */
    .section-summary{
      display:flex;
      flex-direction:column;
      gap:10px;
    }

    /* Cada fila editorial */
    .summary-row{
      border:1px solid #edf0f7;
      border-radius:16px;
      background:#ffffff;
      padding:10px 12px;
      box-shadow: 0 10px 22px rgba(0, 36, 87, 0.04);
    }

    /* Título de la fila (wrap fuerte) */
    .summary-row-title{
      font-size:13px;
      font-weight:900;
      color:#002457;
      line-height:1.2;
      margin:0 0 6px 0;
      overflow-wrap:anywhere;
    }

    /* Subtítulo (SO vs SM) */
    .summary-row-subtitle{
      font-size:11px;
      color:#5f6b7a;
      font-weight:700;
      margin:0 0 8px 0;
    }

    /* Contenedor del plot: alto fijo para consistencia */
    .summary-row-plot{
      height:84px;
      overflow:hidden;
    }

    /* ====== Plotly: texto dentro de barras ====== */
    .plotly text{
      font-weight:800 !important;
    }

    /* Evita recortes raros de svg */
    .plotly .main-svg{
      overflow: visible !important;
    }

    /* ====== Plotly hover: look limpio ====== */
    .plotly .hoverlayer .hovertext{
      font-family: Arial, sans-serif !important;
      border-radius: 10px !important;
    }

    /* ====== DT centrado total ====== */
    table.dataTable thead th,
    table.dataTable tbody td{
      text-align:center !important;
      vertical-align:middle !important;
    }

    /* Sidebar KPI stack: todo estira al mismo ancho */
    .kpi-sidebar-stack{
      display: flex;
      flex-direction: column;
      gap: 12px;
      align-items: stretch;   /* clave: mismo ancho */
      width: 100%;
      box-sizing: border-box;
    }

    /* Los 3 bloques deben tener el mismo ancho */
    .kpi-n-card,
    .kpi-cell{
      width: 100%;
      box-sizing: border-box;
    }

    /* N centrado y sin “sobresalirse” */
    .kpi-n-card{
      display: flex;
      align-items: center;
      justify-content: center;
      text-align: center;
      padding: 12px 10px;
      border-radius: 12px;
      overflow: hidden;       /* por si el texto es largo */
    }

    /* Texto N */
    .kpi-n-text{
      font-weight: 700;
      font-size: 16px;
      line-height: 1.2;
      max-width: 100%;
      white-space: normal;
      word-break: break-word;
    }

    /* Donut title centrado (si no lo estaba) */
    .kpi-donut-title{
      text-align: center;
    }


    /* ============================================================
       ====== PATCH FINAL KPI SIDEBAR (NO ROMPER NADA) ======
       - Corrige el overflow horrible del N (min-width / flex-basis).
       - Asegura mismo ancho real en sidebar.
       - Evita que Plotly empuje el layout.
       ============================================================ */

    /* El contenedor card no deja salir nada (solo en sidebar ayuda muchísimo) */
    .sidebarPanel .cardbox{ overflow:hidden; }

    /* Cuando el perfil está en modo sidebar vertical,
       se anulan las decisiones “horizontales” que lo revientan. */
    .kpi-sidebar-stack .kpi-profile-row{ display:block !important; }
    .kpi-sidebar-stack .kpi-donuts{ display:block !important; }

    /* Mata los min-width/42% que causan el desborde en pantallas chicas */
    .kpi-sidebar-stack .kpi-n-card{
      flex: 0 0 auto !important;
      min-width: 0 !important;
      width: 100% !important;
      max-width: 100% !important;
      box-sizing: border-box !important;
      align-items: center !important;
      justify-content: center !important;
      padding: 12px 12px !important;
      border-radius: 16px !important;
    }

    /* El chip N debe medir lo mismo que las cards (y nunca “salirse”) */
    .kpi-sidebar-stack .kpi-n-chip{
      width: 100% !important;
      max-width: 100% !important;
      box-sizing: border-box !important;
      margin: 0 !important;
      justify-content: center !important;
    }

    .kpi-sidebar-stack .kpi-n-text{
      width: 100% !important;
      text-align: center !important;
      max-width: 100% !important;
      white-space: normal !important;
      word-break: break-word !important;
      font-weight: 900 !important;
      font-size: 18px !important;
    }

    /* KPI cards: mismo ancho y sin empujar */
    .kpi-sidebar-stack .kpi-cell{
      width: 100% !important;
      max-width: 100% !important;
      min-width: 0 !important;
      box-sizing: border-box !important;
      margin: 0 !important;
    }

    /* Plotly dentro del sidebar: nunca exceder el contenedor */
    .kpi-sidebar-stack .plotly.html-widget,
    .kpi-sidebar-stack .plot-container,
    .kpi-sidebar-stack .svg-container{
      width: 100% !important;
      max-width: 100% !important;
    }

    /* Si algún SVG/capa intenta desbordar, se corta en el KPI */
    .kpi-sidebar-stack .kpi-cell{ overflow:hidden !important; }

    /* ============================================================
   PATCH EXTRA: evitar KPIs “comprimidos” en sidebar
   (no rompe nada: solo estabiliza altura/auto-size de Plotly)
   ============================================================ */

/* Asegura que cada KPI tenga aire suficiente */
.kpi-sidebar-stack .kpi-cell{
  min-height: 310px;        /* title + donut + leyenda */
  padding: 10px 10px 12px 10px;
  gap: 6px;                 /* si el browser soporta gap en flex */
}

/* Fuerza la altura REAL del contenedor Shiny del widget */
#kpi_plot_1, #kpi_plot_2{
  height: 210px !important;
  min-height: 210px !important;
}

/* Fuerza la altura del contenedor interno de Plotly */
#kpi_plot_1 .plot-container,
#kpi_plot_2 .plot-container,
#kpi_plot_1 .svg-container,
#kpi_plot_2 .svg-container{
  height: 210px !important;
  min-height: 210px !important;
}

/* Evita que el widget “se estire raro” o se encoja en flex */
.kpi-sidebar-stack .plotly.html-widget{
  width: 100% !important;
  max-width: 100% !important;
  flex: 0 0 auto !important;
}

/* IMPORTANTE: para donuts, permitir overflow visible del SVG,
   porque recortar aquí los hace verse “aplastados” */
.kpi-sidebar-stack .kpi-cell{
  overflow: visible !important;
}

.kpi-sidebar-stack .kpi-legend{
  max-height: 36px;
  overflow: hidden;
}

/* ============================================================
   FIX: leyenda KPI no se corta (sidebar)
   - La card permite scroll suave SOLO si la leyenda crece mucho.
   - El donut mantiene su alto fijo.
   ============================================================ */

/* La card ya no “corta” el contenido inferior */
.kpi-sidebar-stack .kpi-cell{
  overflow: visible !important;
  padding-bottom: 14px !important;  /* aire para la leyenda */
}

/* La leyenda puede ocupar más líneas sin recortarse */
.kpi-sidebar-stack .kpi-legend{
  margin-top: 8px !important;
  padding-bottom: 6px !important;
  line-height: 1.25 !important;
  white-space: normal !important;
}

/* Si la leyenda se vuelve muy larga: scroll interno en vez de corte */
.kpi-sidebar-stack .kpi-cell .kpi-legend{
  max-height: 90px;       /* ajusta 70–120 según te guste */
  overflow-y: auto;
}

/* Mantener el donut estable */
#kpi_plot_1, #kpi_plot_2{
  height: 210px !important;
  min-height: 210px !important;
}
#kpi_plot_1 .plot-container,
#kpi_plot_2 .plot-container,
#kpi_plot_1 .svg-container,
#kpi_plot_2 .svg-container{
  height: 210px !important;
  min-height: 210px !important;
}

.sidebarPanel .cardbox{
  overflow: visible !important;
}

/* ===== KPI sidebar: más alto real + nada de recorte ===== */

/* 1) No recortar dentro de la card del KPI */
.kpi-sidebar-stack .kpi-cell{
  overflow: visible !important;
  padding-bottom: 14px !important;
}

/* 2) Asegurar que el plotlyOutput respete el alto nuevo */
#kpi_plot_1, #kpi_plot_2{
  height: 260px !important;
  min-height: 260px !important;
}

#kpi_plot_1 .plot-container,
#kpi_plot_2 .plot-container,
#kpi_plot_1 .svg-container,
#kpi_plot_2 .svg-container{
  height: 260px !important;
  min-height: 260px !important;
}

/* 3) La leyenda puede ocupar 2–3 líneas sin cortarse */
.kpi-sidebar-stack .kpi-legend{
  margin-top: 8px !important;
  padding-bottom: 8px !important;
  line-height: 1.25 !important;
  white-space: normal !important;
}

/* 4) Si el contenedor general del sidebar está cortando (muy común) */
.sidebarPanel .cardbox{
  overflow: visible !important;
}

/* ============================================================
   ====== PATCH KPI SIDEBAR v2 (leyenda DENTRO y completa) ======
   - La leyenda NO debe salir del bloque.
   - El bloque debe tener altura suficiente para donut + leyenda.
   - El plot se hace un poco más bajo para “dejar aire” abajo.
   ============================================================ */

/* 1) El KPI card SÍ contiene (no deja que la leyenda se salga) */
.kpi-sidebar-stack .kpi-cell{
  overflow: hidden !important;     /* clave: todo queda dentro del borde */
  height: auto !important;
  min-height: 340px !important;    /* ajusta si tu leyenda es muy larga */
  padding-bottom: 14px !important; /* aire para leyenda */
}

/* 2) El plot dentro del KPI se hace un poco más bajo */
#kpi_plot_1, #kpi_plot_2{
  height: 220px !important;
  min-height: 220px !important;
}

/* Asegura que plotly respete el alto */
#kpi_plot_1 .plot-container,
#kpi_plot_2 .plot-container,
#kpi_plot_1 .svg-container,
#kpi_plot_2 .svg-container{
  height: 220px !important;
  min-height: 220px !important;
}

/* 3) Leyenda: dentro, con padding y wrap */
.kpi-sidebar-stack .kpi-legend{
  margin-top: 8px !important;
  padding: 0 8px 10px 8px !important;
  line-height: 1.25 !important;
  white-space: normal !important;
  justify-content: center !important;
}

/* 4) IMPORTANTÍSIMO: si antes dejaste esto en “visible”, lo anulamos aquí */
.sidebarPanel .cardbox{
  overflow: hidden !important;     /* el contenedor general no debe dejar “flotar” cosas */
}

/* ============================================================
   FIX RESUMEN: Select_multiple con múltiples barras
   - No rompe SO (solo aplica cuando existe .sm-card-inner)
   ============================================================ */
.summary-row-plot:has(.sm-card-inner){
  height: auto !important;
  overflow: visible !important;
}

.sm-card-inner{
  display: flex;
  flex-direction: column;
  gap: 12px;
  height: auto !important;
  overflow: visible !important;
}

.sm-option-block{
  height: auto !important;
  overflow: visible !important;
}


/* Respiración general arriba */
body{
  padding-top: 14px;
}

/* Topbar: más aire y menos sensación de “pegado” */
.topbar{
  padding: 18px 18px;         /* antes 14px 16px */
  margin-bottom: 16px;         /* antes 14px */
}

/* Título un poco más relajado (sin aplastarlo) */
.topbar-title{
  padding-top: 2px;
}

/* Contenedor nav: darle aire */
.navbar{
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
  margin-bottom: 18px !important;
}

/* UL de tabs: que parezca barra moderna */
.navbar .navbar-nav{
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
  margin: 0;
  padding: 10px 12px;
  background: #ffffff;
  border: 1px solid #e6e9f2;
  border-radius: 18px;
  box-shadow: 0 14px 34px rgba(0, 36, 87, 0.06);
}

/* Cada tab como “pill” */
.navbar .navbar-nav > li > a{
  border-radius: 999px !important;
  padding: 10px 14px !important;
  font-weight: 800;
  color: #002457 !important;
  background: transparent !important;
  border: 1px solid transparent !important;
}

/* Hover elegante */
.navbar .navbar-nav > li > a:hover{
  background: rgba(0, 36, 87, 0.06) !important;
  border-color: rgba(0, 36, 87, 0.12) !important;
}

/* Tab activa: pill sólida suave */
.navbar .navbar-nav > .active > a,
.navbar .navbar-nav > .active > a:focus,
.navbar .navbar-nav > .active > a:hover{
  background: rgba(0, 36, 87, 0.10) !important;
  border-color: rgba(0, 36, 87, 0.22) !important;
  color: #002457 !important;
}

/* El navbar NO debe tener padding propio */
.navbar{
  padding-left: 0 !important;
  padding-right: 0 !important;
}

/* Alinear pestañas con sidebar */
.navbar .navbar-nav{
  margin-left: 0 !important;
  padding-left: 0 !important;
}

/* Match exacto con sidebar */
.navbar .navbar-nav{
  padding-left: 10px;   /* prueba 10–12 si quieres microajuste */
}

.col-sm-3, .col-sm-9{
  padding-left: 10px;
  padding-right: 10px;
}

/* ====== Alineación navbar con sidebar ====== */

/* El navbar no debe empujar el contenido */
.navbar{
  padding-left: 0 !important;
  padding-right: 0 !important;
}

/* El UL de pestañas empieza donde empieza el sidebar */
.navbar .navbar-nav{
  margin-left: 0 !important;
  padding-left: 10px;   /* mismo padding que columnas */
  padding-right: 10px;
}


  "))
    ),

    shiny::div(
      class = "topbar",
      shiny::div(class = "topbar-title", titulo),
      if (!is.null(logo_src)) shiny::tags$img(
        src   = logo_src,
        alt   = logo_alt,
        class = "topbar-logo",
        style = paste0("height:", as.integer(logo_height_px), "px;")
      )
    ),

    # ✅ navbar dinámica SIN args nombrados para tabPanels (evita buildTabset error)
    do.call(
      shiny::navbarPage,
      c(
        list(title = NULL, id = "tabs_main"),
        unname(lapply(tabs_registry, function(def) def$ui(ctx)))
      )
    )
  )

  # ------------------------------- SERVER ------------------------------------
  server <- function(input, output, session) {
    for (nm in names(tabs_registry)) {
      tabs_registry[[nm]]$server(ctx, input, output, session)
    }
  }

  shiny::shinyApp(ui = ui, server = server)
}
