# =============================================================================
# Explorador interactivo: reporte_interactivo()
# - Helpers compartidos
# - Función exportada
# - Wiring UI/Server (delegación a tabs)
# =============================================================================

`%||%` <- function(x, y) if (!is.null(x)) x else y

# -----------------------------------------------------------------------------
# Helpers internos
# -----------------------------------------------------------------------------

.get_label_col_safe <- function(df) {
  if (is.null(df)) return(NULL)
  if ("label" %in% names(df)) return("label")
  lab_candidates <- grep("^label(::|$)", names(df), value = TRUE)
  if (length(lab_candidates)) return(lab_candidates[1])
  NULL
}

.get_list_name_safe <- function(survey, var) {
  if (is.null(survey) || !all(c("name", "list_name") %in% names(survey))) {
    return(NA_character_)
  }
  i <- which(!is.na(survey$name) & survey$name == var)[1]
  if (is.na(i)) return(NA_character_)

  ln <- as.character(survey$list_name[i])
  if (is.na(ln) || !nzchar(ln)) return(NA_character_)
  ln
}

.wrap_y <- function(x, width = 35) {
  x <- as.character(x)
  if (requireNamespace("stringr", quietly = TRUE)) {
    x <- stringr::str_wrap(x, width = width)
  }
  gsub("\n", "<br>", x, fixed = TRUE)
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

    ln <- .get_list_name_safe(surv, var)
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

  label_col <- .get_label_col_safe(instrumento$choices)

  if (!is.null(instrumento$choices) &&
      all(c("list_name", "name") %in% names(instrumento$choices)) &&
      !is.null(label_col) && label_col %in% names(instrumento$choices) &&
      !is.na(list_var) && nzchar(list_var) &&
      !is.null(names(pal))) {

    ch <- instrumento$choices[instrumento$choices$list_name == list_var, , drop = FALSE]
    map_code_to_label <- stats::setNames(
      as.character(ch[[label_col]]),
      as.character(ch$name)
    )

    idx <- names(pal) %in% names(map_code_to_label)
    if (any(idx)) {
      pal_lab <- stats::setNames(
        pal[idx],
        map_code_to_label[names(pal)[idx]]
      )

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

  var <- trimws(as.character(var)[1])
  surv <- instrumento$survey

  if (!is.null(surv) && "name" %in% names(surv)) {

    label_col <- .get_label_col_safe(surv)

    if (!is.null(label_col) && label_col %in% names(surv)) {
      nm <- trimws(as.character(surv$name))
      i  <- which(!is.na(nm) & nm == var)[1]

      if (!is.na(i)) {
        lab <- surv[[label_col]][i]
        if (!is.na(lab) && nzchar(trimws(as.character(lab)))) {
          return(as.character(lab))
        }
      }
    }
  }

  if (!is.null(data) && var %in% names(data)) {
    vl <- attr(data[[var]], "label", exact = TRUE)
    if (!is.null(vl) && nzchar(trimws(as.character(vl)))) {
      return(as.character(vl))
    }
  }

  var
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
  label_col <- .get_label_col_safe(choices)

  if (is.null(survey) || !"name" %in% names(survey)) {
    stop("El `instrumento` debe contener `survey` válido.", call. = FALSE)
  }

  idx_var <- which(!is.na(survey$name) & as.character(survey$name) == var)[1]
  if (is.na(idx_var)) {
    stop("La variable '", var, "' no está en `instrumento$survey`.", call. = FALSE)
  }
  list_main <- as.character(survey$list_name[idx_var])

  if (!is.null(choices) &&
      all(c("list_name", "name") %in% names(choices)) &&
      !is.null(label_col) && label_col %in% names(choices) &&
      !is.na(list_main) && nzchar(list_main)) {

    ch_main      <- choices[choices$list_name == list_main, , drop = FALSE]
    codigos_main <- as.character(ch_main$name)
    labels_main  <- as.character(ch_main[[label_col]])
  } else {
    codigos_main <- sort(unique(as.character(data[[var]])))
    labels_main  <- codigos_main
  }

  map_main <- stats::setNames(labels_main, codigos_main)
  orden_lvls_main <- map_main[codigos_main]

  df <- data
  if (!var %in% names(df)) {
    stop("La variable '", var, "' no existe en `data`.", call. = FALSE)
  }

  df[[var]] <- as.character(df[[var]])
  df <- df[!is.na(df[[var]]), , drop = FALSE]

  if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0) {
    df <- df[!(df[[var]] %in% as.character(codigos_perdidos)), , drop = FALSE]
  }

  if (nrow(df) == 0L) {
    stop("No hay datos válidos para '", var, "'.", call. = FALSE)
  }

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

    df_tab$opcion_label <- factor(
      df_tab$opcion_label,
      levels = unique(orden_lvls_main[!is.na(orden_lvls_main)])
    )

    df_tab <- df_tab[order(df_tab$opcion_label), , drop = FALSE]
    return(df_tab)
  }

  if (!var_cruce %in% names(df)) {
    stop("Cruce '", var_cruce, "' no existe en `data`.", call. = FALSE)
  }

  df[[var_cruce]] <- as.character(df[[var_cruce]])

  fila_cruce <- survey[survey$name == var_cruce, , drop = FALSE]
  list_cruce <- if (nrow(fila_cruce)) fila_cruce$list_name[1] else NA_character_

  if (!is.null(choices) &&
      all(c("list_name", "name") %in% names(choices)) &&
      !is.null(label_col) && label_col %in% names(choices) &&
      !is.na(list_cruce) && nzchar(list_cruce)) {

    ch_cruce  <- choices[choices$list_name == list_cruce, , drop = FALSE]
    map_cruce <- stats::setNames(as.character(ch_cruce[[label_col]]), as.character(ch_cruce$name))
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

  df_tab$opcion_label  <- factor(
    df_tab$opcion_label,
    levels = unique(orden_lvls_main[!is.na(orden_lvls_main)])
  )
  df_tab$estrato_label <- factor(
    df_tab$estrato_label,
    levels = sort(unique(df_tab$estrato_label))
  )

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

# =============================================================================
# Helper para variables select_multiple "madre" que en la data viven como dummies
# =============================================================================
resolver_var_spec <- function(var_madre, ctx, df = NULL) {

  `%||%` <- get0("%||%", ifnotfound = function(x, y) if (!is.null(x)) x else y)

  data <- df %||% ctx$data
  inst <- ctx$instrumento

  if (is.null(data) || !is.data.frame(data) || is.null(inst)) {
    return(list(
      var_madre = var_madre,
      cols = character(0),
      map_code_to_label = list(),
      list_name = NA_character_,
      col_compact = NA_character_
    ))
  }

  var_esc <- gsub("([\\W])", "\\\\\\1", var_madre)
  pat_dum <- paste0("^", var_esc, "(\\.|_recod\\.)")
  cols <- grep(pat_dum, names(data), value = TRUE)

  col_compact <- NA_character_
  cand1 <- paste0(var_madre, "_ORIG")
  if (cand1 %in% names(data)) {
    col_compact <- cand1
  } else if (var_madre %in% names(data)) {
    col_compact <- var_madre
  }

  surv <- inst$survey %||% NULL
  ch   <- inst$choices %||% NULL

  list_name <- NA_character_
  if (!is.null(surv) && all(c("name", "list_name") %in% names(surv))) {
    i <- which(!is.na(surv$name) & surv$name == var_madre)[1]
    if (!is.na(i)) {
      list_name <- as.character(surv$list_name[i])
      if (is.na(list_name) || !nzchar(list_name)) list_name <- NA_character_
    }
  }

  map_code_to_label <- NULL
  label_col <- .get_label_col_safe(ch)

  if (!is.null(ch) &&
      all(c("list_name", "name") %in% names(ch)) &&
      !is.null(label_col) && label_col %in% names(ch)) {

    if (!is.na(list_name) && nzchar(list_name)) {
      ch_v <- ch[ch$list_name == list_name, , drop = FALSE]
      if (nrow(ch_v)) {
        map_code_to_label <- stats::setNames(
          as.character(ch_v[[label_col]]),
          as.character(ch_v$name)
        )
      }
    }
  }

  if (is.null(map_code_to_label)) {
    cand_attr <- NULL
    if (!is.na(col_compact) && col_compact %in% names(data)) cand_attr <- col_compact
    if (is.null(cand_attr) && length(cols)) cand_attr <- cols[1]

    if (!is.null(cand_attr) && cand_attr %in% names(data)) {
      labs <- attr(data[[cand_attr]], "labels", exact = TRUE)
      if (!is.null(labs) && length(labs) > 0) {
        map_code_to_label <- stats::setNames(
          as.character(unname(labs)),
          as.character(names(labs))
        )
      }
    }
  }

  if (is.null(map_code_to_label)) map_code_to_label <- character(0)

  dummy_code <- function(x) {
    sub(paste0("^", var_esc, "(\\.|_recod\\.)"), "", x)
  }

  dummy_codes <- if (length(cols)) dummy_code(cols) else character(0)

  codes_order <- character(0)
  if (length(map_code_to_label) > 0) {
    codes_order <- as.character(names(map_code_to_label))
  }

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

  if (!length(codes_order) && length(dummy_codes)) {
    codes_order <- unique(dummy_codes)
  }

  if (length(codes_order)) {
    suppressWarnings(num <- as.numeric(codes_order))
    if (!all(is.na(num))) {
      ord <- order(is.na(num), num, codes_order)
      codes_order <- codes_order[ord]
    } else {
      codes_order <- sort(codes_order)
    }
  }

  if (length(cols) && length(codes_order)) {
    ord_idx <- match(dummy_codes, codes_order)
    if (all(is.na(ord_idx))) {
      ord_idx <- seq_along(cols)
    } else {
      nf <- is.na(ord_idx)
      if (any(nf)) {
        base <- max(ord_idx, na.rm = TRUE)
        ord_idx[nf] <- base + seq_len(sum(nf))
      }
    }
    cols <- cols[order(ord_idx)]
  }

  if (length(dummy_codes)) {
    falt <- setdiff(dummy_codes, names(map_code_to_label))
    if (length(falt)) {
      extra <- stats::setNames(falt, falt)
      map_code_to_label <- c(map_code_to_label, extra)
    }
  }

  map_list <- as.list(map_code_to_label)

  list(
    var_madre = var_madre,
    cols = cols,
    map_code_to_label = map_list,
    list_name = list_name,
    col_compact = col_compact
  )
}

.plot_dummy_yesno <- function(x, label_opcion = NULL) {

  BAR_HEIGHT <- 64
  PCT_FSIZE  <- 13

  x <- x[!is.na(x)]
  if (!length(x)) {
    return(
      plotly::plot_ly(height = BAR_HEIGHT) |>
        plotly::layout(
          annotations = list(list(text = "Sin datos", showarrow = FALSE)),
          margin = list(l = 10, r = 10, t = 0, b = 0)
        ) |>
        plotly::config(displayModeBar = FALSE, responsive = TRUE)
    )
  }

  n_si <- sum(x == 1)
  n_no <- sum(x == 0)
  tot  <- n_si + n_no

  tab <- data.frame(
    resp = c("Sí", "No"),
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
      textfont = list(color = "white", size = PCT_FSIZE),
      hoverinfo = "skip"
    ) |>
    plotly::layout(
      barmode = "stack",
      xaxis = list(range = c(0, 1), visible = FALSE),
      yaxis = list(visible = FALSE),
      margin = list(l = 10, r = 10, t = 0, b = 0),
      showlegend = FALSE
    ) |>
    plotly::config(displayModeBar = FALSE, responsive = TRUE)

  p
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
          vars_so     = ctx$so_vars %||% character(0),
          vars_sm_madres = ctx$sm_madres %||% character(0),
          colores_apiladas_por_listname = ctx$colores_apiladas_por_listname,
          codigos_perdidos = ctx$codigos_perdidos,
          weight_col = "peso",
          orders_list = ctx$instrumento$orders_list %||% NULL,
          labels_override = NULL,
          theme_app = ctx$theme_app
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
    tabs = c("resumen", "relacion", "base_datos"),
    theme_app  = NULL
) {

  if (!requireNamespace("shiny", quietly = TRUE) ||
      !requireNamespace("plotly", quietly = TRUE) ||
      !requireNamespace("dplyr",  quietly = TRUE) ||
      !requireNamespace("DT",     quietly = TRUE)) {
    stop("Se requieren 'shiny', 'plotly', 'dplyr' y 'DT' para `reporte_interactivo()`.", call. = FALSE)
  }

  if (!exists("reporte_interactivo_theme_css", mode = "function") ||
      !exists("reporte_interactivo_theme_js",  mode = "function")) {
    stop(
      "No se encontraron las funciones de tema visual. ",
      "Asegúrate de cargar también el archivo `reporte_interactivo_theme.R`.",
      call. = FALSE
    )
  }

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

    if (!is.null(instrumento$survey) && all(c("name", "type") %in% names(instrumento$survey))) {
      fila <- instrumento$survey[instrumento$survey$name == v, , drop = FALSE]
      if (nrow(fila)) {
        tp <- tolower(as.character(fila$type[1]))
        if (tp %in% c("start", "end", "deviceid", "subscriberid", "simserial",
                      "phonenumber", "today", "username", "audit")) {
          return(TRUE)
        }
      }
    }

    FALSE
  }

  label_var <- function(v) .obtener_label_var(v, instrumento, data)

  vars_data_visibles <- setdiff(
    names(data),
    names(data)[vapply(names(data), .is_tecnica, logical(1), instrumento = instrumento)]
  )

  so_inst <- survey$name[grepl("^select_one\\b", tolower(survey$type))]
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

  kpi_vars <- (kpi_vars %||% character(0))
  kpi_vars <- unique(kpi_vars[kpi_vars %in% names(data)])
  if (length(kpi_vars) > 2L) kpi_vars <- kpi_vars[1:2]

  secciones_limpias <- lapply(secciones, function(vs) {

    keep <- vs[vs %in% names(data)]

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
    logo_src <- sub("^www/", "", logo_png)
  }

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
    so_vars   = so_vars,
    sm_madres = sm_disponibles,
    theme_app = theme_app
  )

  tabs_registry <- .make_tabs_registry(ctx, tabs = tabs)

  ui <- shiny::fluidPage(

    shiny::tags$head(
      reporte_interactivo_theme_css(theme_app = theme_app),
      reporte_interactivo_theme_js()
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

    do.call(
      shiny::navbarPage,
      c(
        list(title = NULL, id = "tabs_main"),
        unname(lapply(tabs_registry, function(def) def$ui(ctx)))
      )
    )
  )

  server <- function(input, output, session) {
    for (nm in names(tabs_registry)) {
      tabs_registry[[nm]]$server(ctx, input, output, session)
    }
  }

  shiny::shinyApp(ui = ui, server = server)
}
