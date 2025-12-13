# =============================================================================
# Explorador interactivo: reporte_interactivo()
# Pestaña 1: Resumen (gráfico + tabla + perfil N + 2 KPIs)
# Pestaña 2: Base de datos (diccionario + tabla de datos + descarga + toggle etiquetas/códigos)
# =============================================================================

`%||%` <- function(x, y) if (!is.null(x)) x else y

# -----------------------------------------------------------------------------
# Helpers internos
# -----------------------------------------------------------------------------

#' Resolver paleta para una variable (alineada a labels visibles)
#' - Soporta paletas nombradas por label o por code.
#' - Si faltan categorías, completa de forma estable.
#' @keywords internal
#' @noRd
.resolver_paleta_var <- function(var,
                                 instrumento,
                                 colores_apiladas_por_listname,
                                 opcion_levels) {

  surv <- instrumento$survey
  pal  <- NULL

  # 1) buscar list_name
  if (!is.null(colores_apiladas_por_listname) &&
      !is.null(surv) &&
      all(c("name", "list_name") %in% names(surv))) {

    ln <- surv$list_name[surv$name == var][1]
    if (!is.na(ln) && ln %in% names(colores_apiladas_por_listname)) {
      pal <- colores_apiladas_por_listname[[ln]]
    }
  }

  # 2) fallback
  if (is.null(pal) || !length(pal)) {
    out <- grDevices::hcl.colors(max(3L, length(opcion_levels)), "Blues")
    out <- out[seq_len(length(opcion_levels))]
    names(out) <- opcion_levels
    return(out)
  }

  # 3) si paleta ya está nombrada por labels: alinear
  if (!is.null(names(pal)) && all(opcion_levels %in% names(pal))) {
    pal2 <- pal[opcion_levels]
    names(pal2) <- opcion_levels
    return(pal2)
  }

  # 4) si paleta está nombrada por codes: mapear code -> label con choices
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

      # completar faltantes si hiciera falta
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

  # 5) último recurso: reciclar y nombrar
  pal <- rep(pal, length.out = length(opcion_levels))
  names(pal) <- opcion_levels
  pal
}

#' Obtener etiqueta legible de una variable
#' @keywords internal
#' @noRd
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

#' Wrap del título en HTML (usa <br>)
#' @keywords internal
#' @noRd
.wrap_titulo_html <- function(txt, width = 120) {
  if (!requireNamespace("stringr", quietly = TRUE)) return(txt)
  txt <- as.character(txt)
  if (!nzchar(txt)) return(txt)
  lineas <- stringr::str_wrap(txt, width = width)
  paste(lineas, collapse = "<br>")
}

#' Recalcular porcentajes ENTEROS por estrato (suma exacta = 100)
#' @keywords internal
#' @noRd
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

#' Tabla de proporciones (simple o cruzada)
#' @keywords internal
#' @noRd
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

  df <- data
  if (!var %in% names(df)) stop("La variable '", var, "' no existe en `data`.", call. = FALSE)

  df[[var]] <- as.character(df[[var]])
  df <- df[!is.na(df[[var]]), , drop = FALSE]

  if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0) {
    df <- df[!(df[[var]] %in% as.character(codigos_perdidos)), , drop = FALSE]
  }

  if (nrow(df) == 0L) stop("No hay datos válidos para '", var, "'.", call. = FALSE)

  # --- sin cruce
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

    orden_lvls <- map_main[codigos_main]
    df_tab$opcion_label <- factor(df_tab$opcion_label, levels = unique(orden_lvls[!is.na(orden_lvls)]))
    df_tab <- df_tab[order(df_tab$opcion_label), , drop = FALSE]
    return(df_tab)
  }

  # --- con cruce
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

  orden_lvls_main <- map_main[codigos_main]
  df_tab$opcion_label <- factor(df_tab$opcion_label, levels = unique(orden_lvls_main[!is.na(orden_lvls_main)]))
  df_tab$estrato_label <- factor(df_tab$estrato_label, levels = sort(unique(df_tab$estrato_label)))

  if (length(unique(df_tab$estrato_label)) == 1 &&
      unique(as.character(df_tab$estrato_label)) %in% c("Total", "TOTAL", "total")) {
    df_tab$estrato_label <- factor(rep("", nrow(df_tab)))
  }

  df_tab[order(df_tab$estrato_label, df_tab$opcion_label), , drop = FALSE]
}

#' Tabla resumen
#' @keywords internal
#' @noRd
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

#' Plot principal (barras apiladas horizontales) → plotly
#' @keywords internal
#' @noRd
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

  # paleta robusta SIEMPRE
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

    # asegurar columnas con longitudes idénticas al data=
    df_opt$texto_in  <- paste0("<b>", df_opt$texto_pct, "</b>")
    # df_opt$hover_text ya se construye arriba como vector del mismo largo

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

# -----------------------------------------------------------------------------
# KPI half-donut (SIEMPRE) + leyenda abajo (sin recortes)
# Devuelve list(plot = <plotly>, legend = data.frame(label, color))
# -----------------------------------------------------------------------------
#' @keywords internal
#' @noRd
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


# -----------------------------------------------------------------------------
# App principal
# -----------------------------------------------------------------------------

#' Explorador interactivo de resultados (pestaña Resumen + pestaña Base de datos)
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
    logo_height_px = 52
) {

  if (!requireNamespace("shiny", quietly = TRUE) ||
      !requireNamespace("plotly", quietly = TRUE) ||
      !requireNamespace("dplyr",  quietly = TRUE) ||
      !requireNamespace("DT",     quietly = TRUE)) {
    stop("Se requieren 'shiny', 'plotly', 'dplyr' y 'DT' para `reporte_interactivo()`.", call. = FALSE)
  }


  # -------------------------------------------------------------------------
  # Asegurar metadatos SPSS-style (label / labels / measure) dentro de la app
  # -------------------------------------------------------------------------
  tiene_labels <- any(vapply(names(data), function(v) {
    !is.null(attr(data[[v]], "label",  exact = TRUE)) ||
      !is.null(attr(data[[v]], "labels", exact = TRUE)) ||
      !is.null(attr(data[[v]], "measure", exact = TRUE))
  }, logical(1)))

  if (!inherits(data, "prosecnur_reporte_tbl") || !tiene_labels) {
    # Re-derivar los metadatos desde instrumento (sin mostrar nada “riesgoso”)
    data <- reporte_data(
      data        = data,
      instrumento = instrumento
      # var_peso = NULL  # si lo usas, pásalo aquí
    )
  }

  survey <- instrumento$survey
  if (is.null(survey) || !"name" %in% names(survey)) {
    stop("El `instrumento` debe contener un `survey` válido.", call. = FALSE)
  }

  if (is.null(secciones) || !length(secciones)) {
    stop("`secciones` debe ser una lista nombrada con vectores de variables.", call. = FALSE)
  }

  # -------------------- Helpers de pestaña Base de datos ---------------------

  .is_tecnica <- function(v, instrumento) {
    if (!nzchar(v)) return(TRUE)
    if (startsWith(v, "_")) return(TRUE)
    # variables temporales declaradas por instrumento
    vf <- as.character(attr(data, "vars_fecha", exact = TRUE) %||% instrumento$vars_fecha %||% character(0))
    vh <- as.character(attr(data, "vars_hora", exact = TRUE) %||% instrumento$vars_hora %||% character(0))
    vd <- as.character(attr(data, "vars_datetime", exact = TRUE) %||% instrumento$vars_datetime %||% character(0))
    if (v %in% c(vf, vh, vd)) return(TRUE)

    # “start/end/deviceid/etc” si existieran
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

  # variables candidatas (para mostrar en data): quitar técnicas
  vars_data_visibles <- setdiff(names(data), names(data)[vapply(names(data), .is_tecnica, logical(1), instrumento = instrumento)])

  # ==========================================================
  # Diccionario (CONCEPTUAL): desde instrumento$survey
  # - select_one: variable única
  # - select_multiple: variable madre (NO dummies)
  # ==========================================================

  survey <- instrumento$survey
  choices <- instrumento$choices %||% NULL

  # tipos desde instrumento
  so_inst <- survey$name[grepl("^select_one\\b",  tolower(survey$type))]
  sm_inst <- survey$name[grepl("^select_multiple\\b", tolower(survey$type))]

  # visibles en data (sin técnicas)
  vars_data_visibles <- setdiff(
    names(data),
    names(data)[vapply(names(data), .is_tecnica, logical(1), instrumento = instrumento)]
  )

  # select_one disponible si existe como columna
  so_vars <- intersect(so_inst, vars_data_visibles)

  # select_multiple disponible si existen dummies en data
  # (p.ej. var.xxx o var_recod.xxx o var_otro)
  sm_disponibles <- sm_inst[vapply(sm_inst, function(v) {
    patt <- paste0("^", v, "(\\.|_recod\\.|_otro$)")
    any(grepl(patt, vars_data_visibles))
  }, logical(1))]

  # Variables que el diccionario permitirá seleccionar (solo madres)
  vars_diccionario_all <- sort(unique(c(so_vars, sm_disponibles)))

  # índice: var madre -> columnas dummies presentes en data
  sm_cols_map <- stats::setNames(vector("list", length(sm_disponibles)), sm_disponibles)
  for (v in sm_disponibles) {
    patt <- paste0("^", v, "(\\.|_recod\\.|_otro$)")
    sm_cols_map[[v]] <- grep(patt, vars_data_visibles, value = TRUE)
  }

  # vista etiquetada: transformar columnas usando attr(labels), sin “inventar” nada
  .to_labels_df <- function(df) {
    out <- df
    for (v in names(out)) {
      labs <- attr(out[[v]], "labels", exact = TRUE)
      if (!is.null(labs) && length(labs) > 0) {
        # labs típicamente: names = codes, values = labels (como lo has definido en reporte_data)
        codes <- names(labs)
        lbls  <- unname(labs)

        x <- out[[v]]
        x_chr <- as.character(x)

        map_code_to_label <- stats::setNames(as.character(lbls), as.character(codes))
        x_lbl <- unname(map_code_to_label[x_chr])
        # fallback: si algo no mapea, mantener el código tal cual (sin resaltar NAs)
        x_lbl[is.na(x_lbl) & !is.na(x_chr)] <- x_chr[is.na(x_lbl) & !is.na(x_chr)]
        out[[v]] <- x_lbl
      } else {
        # si no hay labels, dejar tal cual
        out[[v]] <- out[[v]]
      }
    }
    out
  }

  # KPIs (máx 2)
  kpi_vars <- (kpi_vars %||% character(0))
  kpi_vars <- unique(kpi_vars[kpi_vars %in% names(data)])
  if (length(kpi_vars) > 2L) kpi_vars <- kpi_vars[1:2]

  # Secciones: sólo variables presentes en data
  secciones_limpias <- lapply(secciones, function(v) v[v %in% names(data)])
  secciones_limpias <- secciones_limpias[vapply(secciones_limpias, length, integer(1)) > 0]
  if (!length(secciones_limpias)) {
    stop("Ninguna sección de `secciones` tiene variables presentes en `data`.", call. = FALSE)
  }
  secciones_nombres <- names(secciones_limpias)

  # Filtros / cruces (solo pestaña 1)
  facet_vars <- (facet_vars %||% character(0))
  facet_vars <- facet_vars[facet_vars %in% names(data)]
  facet_choices <- stats::setNames(facet_vars, vapply(facet_vars, label_var, character(1)))

  logo_src <- NULL
  if (!is.null(logo_png) && nzchar(logo_png) && file.exists(logo_png)) {
    shiny::addResourcePath("reporte_logo", normalizePath(dirname(logo_png), winslash = "/"))
    logo_src <- paste0("reporte_logo/", basename(logo_png))
  }

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

    shiny::navbarPage(
      title = NULL,
      id    = "tabs_main",

      # =========================
      # TAB 1 — RESUMEN (igual lógica)
      # =========================
      shiny::tabPanel(
        title = "Resumen",

        shiny::sidebarLayout(
          shiny::sidebarPanel(
            width = 3,

            shiny::h3("Variables"),
            shiny::p("Seleccione la sección y la variable principal a visualizar."),

            shiny::selectInput(
              inputId  = "seccion",
              label    = "Sección",
              choices  = stats::setNames(secciones_nombres, secciones_nombres),
              selected = secciones_nombres[1]
            ),

            shiny::selectInput(
              inputId  = "var_principal",
              label    = "Variable principal",
              choices  = NULL
            ),

            shiny::hr(),

            shiny::h3("Filtros"),

            shiny::selectInput(
              inputId  = "filtro_var",
              label    = "Variable de filtro",
              choices  = c("Ninguno" = "", facet_choices),
              selected = ""
            ),

            shiny::uiOutput("filtro_categorias_ui"),

            shiny::actionButton(
              inputId = "limpiar_filtros",
              label   = "Limpiar filtros"
            ),

            shiny::hr(),

            shiny::h3("Cruce"),

            shiny::selectInput(
              inputId  = "var_cruce",
              label    = "Cruce",
              choices  = c("Ninguno" = "", facet_choices),
              selected = ""
            ),

            shiny::actionButton(
              inputId = "limpiar_cruce",
              label   = "Limpiar cruce"
            )
          ),

          shiny::mainPanel(
            width = 9,

            shiny::fluidRow(
              shiny::column(
                width = 12,
                shiny::div(
                  class = "cardbox",
                  shiny::div(class = "cardbox-header", shiny::uiOutput("plot_header")),
                  plotly::plotlyOutput("plot_principal", height = "420px")
                )
              )
            ),

            shiny::br(),

            shiny::fluidRow(
              shiny::column(
                width = 6,
                shiny::div(
                  class = "cardbox",
                  style = "height: 460px; overflow-y: auto;",
                  shiny::div(
                    class = "cardbox-header",
                    shiny::div(class = "cardbox-title", shiny::textOutput("titulo_tabla"))
                  ),
                  DT::dataTableOutput("tabla_principal")
                )
              ),
              shiny::column(
                width = 6,
                shiny::div(
                  class = "cardbox",
                  style = paste(
                    "height: 460px;",
                    "display: flex; flex-direction: column; align-items: stretch;",
                    "overflow: hidden;"
                  ),
                  shiny::uiOutput("kpi_panel")
                )
              )
            ),

            shiny::div(style = "height: 48px;")
          )
        )
      ),

      # =========================
      # TAB 2 — BASE DE DATOS
      # =========================
      shiny::tabPanel(
        title = "Base de datos",

        shiny::sidebarLayout(
          shiny::sidebarPanel(
            width = 3,
            shiny::h3("Diccionario"),
            shiny::p("Información de variables con categorías codificadas."),

            shiny::selectInput(
              inputId  = "data_seccion",
              label    = "Sección",
              choices  = stats::setNames(secciones_nombres, secciones_nombres),
              selected = secciones_nombres[1]
            ),

            shiny::selectInput(
              inputId  = "dicc_var",
              label    = "Variable",
              choices  = c(),
              selected = NULL
            ),

            shiny::div(
              class = "cardbox",
              style = "padding: 10px; margin-top: 10px;",
              shiny::uiOutput("diccionario_detalle")
            ),

            shiny::hr(),

            shiny::h3("Vista"),
            shiny::div(
              class = "toggle-row",
              shiny::span(class = "toggle-label", "Códigos"),
              shiny::tags$label(
                class = "switch",
                shiny::tags$input(id = "vista_etiquetas", type = "checkbox", checked = "checked"),
                shiny::tags$span(class = "slider")
              ),
              shiny::span(class = "toggle-label", "Etiquetas")
            ),
          ),

          shiny::mainPanel(
            width = 9,

            shiny::fluidRow(
              shiny::column(
                width = 12,
                shiny::div(
                  class = "cardbox",
                  shiny::div(class = "cardbox-header",
                             shiny::div(class = "cardbox-title", "Base de datos"),
                             shiny::div(class = "cardbox-subtitle", "Búsqueda, ordenamiento y paginación disponibles.")
                  ),
                  DT::dataTableOutput("tabla_data")
                )
              )
            ),

            shiny::div(style = "height: 48px;")
          )
        )
      )
    )
  )

  # ------------------------------- SERVER ------------------------------------
  server <- function(input, output, session) {

    # ======================
    # TAB 1 — RESUMEN
    # ======================

    shiny::observe({
      sec      <- input$seccion
      vars_sec <- secciones_limpias[[sec]]

      if (is.null(vars_sec) || !length(vars_sec)) {
        shiny::updateSelectInput(session, "var_principal", choices = c(), selected = "")
      } else {
        choices_sec <- stats::setNames(vars_sec, vapply(vars_sec, label_var, character(1)))
        shiny::updateSelectInput(session, "var_principal", choices = choices_sec, selected = vars_sec[1])
      }
    })

    output$filtro_categorias_ui <- shiny::renderUI({
      v <- input$filtro_var
      if (is.null(v) || !nzchar(v) || !v %in% names(data)) return(NULL)

      vals <- sort(unique(as.character(data[[v]])))
      vals <- vals[!is.na(vals)]
      if (!length(vals)) return(NULL)

      surv <- instrumento$survey
      ch   <- instrumento$choices %||% NULL

      labels_vals <- vals

      if (!is.null(surv) && all(c("name", "list_name") %in% names(surv)) &&
          !is.null(ch)   && all(c("list_name","name","label") %in% names(ch))) {

        ln <- surv$list_name[surv$name == v][1]
        if (!is.na(ln) && nzchar(ln)) {
          ch_v <- ch[ch$list_name == ln, , drop = FALSE]
          if (nrow(ch_v)) {
            map_code_to_label <- stats::setNames(as.character(ch_v$label), as.character(ch_v$name))
            labels_vals <- unname(map_code_to_label[vals])
            labels_vals[is.na(labels_vals) | labels_vals == ""] <- vals[is.na(labels_vals) | labels_vals == ""]
          }
        }
      }

      shiny::checkboxGroupInput(
        inputId  = "filtro_categorias",
        label    = label_var(v),
        choices  = stats::setNames(vals, labels_vals),
        selected = vals
      )
    })

    shiny::observeEvent(input$limpiar_filtros, {
      shiny::updateSelectInput(session, inputId = "filtro_var", selected = "")
      if (!is.null(input$filtro_categorias)) {
        shiny::updateCheckboxGroupInput(session, inputId = "filtro_categorias", selected = character(0))
      }
    })

    shiny::observeEvent(input$limpiar_cruce, {
      shiny::updateSelectInput(session, "var_cruce", selected = "")
    })

    data_filtrada <- shiny::reactive({
      df <- data
      v_filtro <- input$filtro_var

      if (!is.null(v_filtro) && nzchar(v_filtro) && v_filtro %in% names(df) &&
          !is.null(input$filtro_categorias)) {

        vals_sel <- input$filtro_categorias
        if (length(vals_sel) > 0L) df <- df[df[[v_filtro]] %in% vals_sel, , drop = FALSE]
      }
      df
    })

    output$plot_header <- shiny::renderUI({
      shiny::req(input$var_principal)

      var_main <- input$var_principal
      titulo_h <- .obtener_label_var(var_main, instrumento, data)
      titulo_h <- .wrap_titulo_html(titulo_h, width = 110)

      cruce <- input$var_cruce
      cruce_txt <- if (!is.null(cruce) && nzchar(cruce)) {
        paste0("Cruce: ", .obtener_label_var(cruce, instrumento, data))
      } else {
        NULL
      }

      shiny::tagList(
        shiny::div(class = "cardbox-title", shiny::HTML(titulo_h)),
        if (!is.null(cruce_txt)) shiny::div(class = "cardbox-subtitle", cruce_txt)
      )
    })

    output$titulo_tabla <- shiny::renderText({
      if (!is.null(input$var_cruce) && nzchar(input$var_cruce)) {
        "Distribución de respuestas por estrato"
      } else {
        "Distribución de respuestas"
      }
    })

    output$plot_principal <- plotly::renderPlotly({
      shiny::req(input$var_principal)

      var_main <- input$var_principal
      df_all   <- data_filtrada()
      df <- if (var_main %in% names(df_all)) df_all[!is.na(df_all[[var_main]]), , drop = FALSE] else df_all

      var_cruce <- input$var_cruce
      if (!nzchar(var_cruce)) var_cruce <- NULL

      if (nrow(df) == 0L) shiny::validate(shiny::need(FALSE, "No hay datos válidos (después de filtros)."))

      titulo_plot <- .wrap_titulo_html(.obtener_label_var(var_main, instrumento, data), width = 120)

      df_tab <- .preparar_tabla_proporciones(
        data             = df,
        instrumento      = instrumento,
        var              = var_main,
        var_cruce        = var_cruce,
        codigos_perdidos = codigos_perdidos
      )

      opcion_levels <- levels(df_tab$opcion_label) %||% unique(df_tab$opcion_label)
      paleta_main <- .resolver_paleta_var(
        var = var_main,
        instrumento = instrumento,
        colores_apiladas_por_listname = colores_apiladas_por_listname,
        opcion_levels = as.character(opcion_levels)
      )

      .construir_plotly_barras(
        df_tab          = df_tab,
        titulo          = titulo_plot,
        var_paleta      = var_main,
        instrumento     = instrumento,
        colores_apiladas_por_listname = colores_apiladas_por_listname,
        paleta_colores  = paleta_main,
        height          = 420,
        mostrar_leyenda = TRUE
      )
    })

    output$tabla_principal <- DT::renderDataTable({
      shiny::req(input$var_principal)

      var_main <- input$var_principal
      df_all   <- data_filtrada()
      df <- if (var_main %in% names(df_all)) df_all[!is.na(df_all[[var_main]]), , drop = FALSE] else df_all

      var_cruce <- input$var_cruce
      if (!nzchar(var_cruce)) var_cruce <- NULL
      if (!nrow(df)) return(NULL)

      df_tab <- .preparar_tabla_proporciones(
        data             = df,
        instrumento      = instrumento,
        var              = var_main,
        var_cruce        = var_cruce,
        codigos_perdidos = codigos_perdidos
      )

      tabla <- .construir_tabla_resumen(df_tab)

      DT::datatable(
        tabla,
        rownames = FALSE,
        options  = list(
          paging = FALSE,
          searching = FALSE,
          info = FALSE,
          language = list(url = "//cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json")
        )
      )
    })

    output$kpi_panel <- shiny::renderUI({

      df_all   <- data_filtrada()
      var_main <- input$var_principal

      df <- if (!is.null(var_main) && nzchar(var_main) && var_main %in% names(df_all)) {
        df_all[!is.na(df_all[[var_main]]), , drop = FALSE]
      } else {
        df_all
      }

      if (!nrow(df)) return(shiny::div("Sin datos para la pregunta seleccionada."))

      n_unidades <- if (!is.null(id_unidad) && id_unidad %in% names(df)) {
        dplyr::n_distinct(df[[id_unidad]])
      } else {
        nrow(df)
      }

      n_sufijo <- if (!is.null(id_unidad) && nzchar(id_unidad)) id_unidad else ""
      texto_N  <- paste0(
        "N: ",
        format(n_unidades, big.mark = ",", scientific = FALSE),
        if (nzchar(n_sufijo)) paste0(" ", n_sufijo) else ""
      )

      # Leyenda (más “secundaria”)
      legend_html <- function(legend_df) {
        shiny::div(
          class = "kpi-legend",
          lapply(seq_len(nrow(legend_df)), function(i) {
            shiny::div(
              class = "kpi-legend-item",
              shiny::span(
                class = "kpi-legend-swatch",
                style = paste0("background:", legend_df$color[i], ";")
              ),
              shiny::span(legend_df$label[i])
            )
          })
        )
      }

      # construir KPIs (máx 2)
      kpi_obj_1 <- NULL
      kpi_obj_2 <- NULL

      if (length(kpi_vars) >= 1) {
        kpi_obj_1 <- .construir_kpi_halfdonut(
          df = df,
          var_kpi = kpi_vars[1],
          instrumento = instrumento,
          colores_apiladas_por_listname = colores_apiladas_por_listname,
          codigos_perdidos = codigos_perdidos
        )
        if (!is.null(kpi_obj_1)) {
          output$kpi_plot_1 <- plotly::renderPlotly(kpi_obj_1$plot)
        }
      }

      if (length(kpi_vars) >= 2) {
        kpi_obj_2 <- .construir_kpi_halfdonut(
          df = df,
          var_kpi = kpi_vars[2],
          instrumento = instrumento,
          colores_apiladas_por_listname = colores_apiladas_por_listname,
          codigos_perdidos = codigos_perdidos
        )
        if (!is.null(kpi_obj_2)) {
          output$kpi_plot_2 <- plotly::renderPlotly(kpi_obj_2$plot)
        }
      }

      # UI final (bloque editorial)
      shiny::div(
        class = "kpi-block",

        # Header igual al bloque de la tabla (mismo look)
        shiny::div(
          class = "cardbox-header",
          shiny::div(class = "cardbox-title", "Perfil de la muestra")
        ),

        # Chip N
        shiny::div(
          class = "kpi-n-chip",
          shiny::div(class = "kpi-n-text", texto_N)
        ),

        # Donuts: 1 o 2
        if (length(kpi_vars) == 1 && !is.null(kpi_obj_1)) {

          shiny::div(
            class = "kpi-cell",
            shiny::div(class = "kpi-donut-title", shiny::HTML(kpi_obj_1$title_html)),
            plotly::plotlyOutput("kpi_plot_1", height = "230px"),
            legend_html(kpi_obj_1$legend)
          )

        } else if (length(kpi_vars) >= 2 && !is.null(kpi_obj_1) && !is.null(kpi_obj_2)) {

          shiny::div(
            class = "kpi-grid",

            shiny::div(
              class = "kpi-cell",
              shiny::div(class = "kpi-donut-title", shiny::HTML(kpi_obj_1$title_html)),
              plotly::plotlyOutput("kpi_plot_1", height = "230px"),
              legend_html(kpi_obj_1$legend)
            ),

            shiny::div(
              class = "kpi-cell",
              shiny::div(class = "kpi-donut-title", shiny::HTML(kpi_obj_2$title_html)),
              plotly::plotlyOutput("kpi_plot_2", height = "230px"),
              legend_html(kpi_obj_2$legend)
            )
          )

        } else {
          shiny::div(
            style="font-size:12px;color:#5f6b7a;",
            "No se pudieron construir KPIs para la selección actual."
          )
        }
      )
    })

    # ======================
    # TAB 2 — BASE DE DATOS
    # ======================

    # ----------------------
    # TAB 2: Diccionario por sección (variables CONCEPTUALES)
    # - select_one: variable existe como columna
    # - select_multiple: variable madre (aunque en data existan solo dummies)
    # ----------------------

    dicc_vars_por_seccion <- lapply(secciones_limpias, function(vs) {
      intersect(vs, vars_diccionario_all)
    })

    shiny::observe({
      sec <- input$data_seccion
      vars_sec <- dicc_vars_por_seccion[[sec]] %||% character(0)

      if (!length(vars_sec)) {
        shiny::updateSelectInput(session, "dicc_var", choices = c(), selected = NULL)
      } else {
        ch <- stats::setNames(vars_sec, vapply(vars_sec, label_var, character(1)))
        shiny::updateSelectInput(session, "dicc_var", choices = ch, selected = vars_sec[1])
      }
    })

    # ----------------------
    # Diccionario: detalle (NO controla la tabla)
    # - select_multiple: muestra opciones del instrumento (NO dummies)
    # ----------------------
    output$diccionario_detalle <- shiny::renderUI({
      v <- input$dicc_var

      if (is.null(v) || !nzchar(v) || !v %in% vars_diccionario_all) {
        return(shiny::div(style="font-size:12px;color:#5f6b7a;", "Sin variables codificadas disponibles."))
      }

      fila <- instrumento$survey[instrumento$survey$name == v, , drop = FALSE]
      tipo_survey <- if (nrow(fila)) tolower(as.character(fila$type[1])) else ""

      es_so <- grepl("^select_one\\b", tipo_survey)
      es_sm <- grepl("^select_multiple\\b", tipo_survey)

      # etiqueta: SO desde data; SM desde instrumento
      etq <- if (es_so && v %in% names(data)) {
        attr(data[[v]], "label", exact = TRUE) %||% label_var(v)
      } else {
        .obtener_label_var(v, instrumento, data = NULL)
      }

      meas <- if (es_so && v %in% names(data)) attr(data[[v]], "measure", exact = TRUE) else NULL
      meas <- if (!is.null(meas) && nzchar(as.character(meas))) toupper(as.character(meas)) else "—"

      tipo <- if (es_so) "Selección única" else if (es_sm) "Selección múltiple" else "Variable codificada"

      shiny::tagList(
        shiny::div(class="dicc-kv",
                   shiny::div(class="dicc-k","Variable"), shiny::div(class="dicc-v", v),
                   shiny::div(class="dicc-k","Etiqueta"), shiny::div(class="dicc-v", as.character(etq)),
                   shiny::div(class="dicc-k","Tipo"),     shiny::div(class="dicc-v", tipo),
                   shiny::div(class="dicc-k","Medición"), shiny::div(class="dicc-v", meas)
        ),
        shiny::hr(),
        shiny::div(style="font-size:12px;font-weight:800;color:#002457;margin-bottom:6px;", "Categorías"),
        DT::DTOutput("dicc_opciones")
      )
    })

    # ----------------------
    # Diccionario: opciones (siempre desde instrumento$choices)
    # ----------------------
    output$dicc_opciones <- DT::renderDT({
      v <- input$dicc_var
      if (is.null(v) || !nzchar(v) || !v %in% vars_diccionario_all) return(NULL)

      fila <- instrumento$survey[instrumento$survey$name == v, , drop = FALSE]
      tipo_survey <- if (nrow(fila)) tolower(as.character(fila$type[1])) else ""
      es_so <- grepl("^select_one\\b", tipo_survey)
      es_sm <- grepl("^select_multiple\\b", tipo_survey)

      ln <- if (nrow(fila) && "list_name" %in% names(fila)) as.character(fila$list_name[1]) else NA_character_
      ch <- instrumento$choices %||% NULL

      # ---- 1) opciones base (siempre desde instrumento$choices)
      opts_df <- NULL
      if (!is.null(ch) && all(c("list_name","name","label") %in% names(ch)) &&
          !is.na(ln) && nzchar(ln)) {

        chv <- ch[ch$list_name == ln, c("name","label"), drop = FALSE]
        if (nrow(chv)) {
          opts_df <- data.frame(
            Código   = as.character(chv$name),
            Etiqueta = as.character(chv$label),
            stringsAsFactors = FALSE
          )
        }
      }

      if (is.null(opts_df) || !nrow(opts_df)) {
        opts_df <- data.frame(Código = character(0), Etiqueta = character(0), stringsAsFactors = FALSE)
      }

      # ---- 2) regla perdidos: solo mostrar 96/97/98/99 si aparecen en la data
      cod_perd <- as.character(codigos_perdidos %||% character(0))
      if (length(cod_perd) > 0 && nrow(opts_df) > 0) {

        # valores observados en data para la variable (códigos)
        vals_obs <- character(0)

        if (es_so && v %in% names(data)) {
          x <- as.character(data[[v]])
          vals_obs <- unique(x[!is.na(x)])

        } else if (es_sm) {
          # para select_multiple: revisar si hay registros con algún dummy=1
          # y usar eso para detectar si se usa "No sabe/No responde/etc" (96..99) cuando existan como opción
          cols <- sm_cols_map[[v]] %||% character(0)
          cols <- cols[cols %in% names(data)]
          if (length(cols)) {
            # casos con al menos un 1
            m <- data[, cols, drop = FALSE]
            m <- as.data.frame(lapply(m, function(z) suppressWarnings(as.numeric(as.character(z)))))
            any_one <- apply(m, 1, function(r) any(r == 1, na.rm = TRUE))
            if (any(any_one, na.rm = TRUE)) {
              # cuáles dummies fueron seleccionadas (colnames con algún 1)
              cols_on <- cols[colSums(m == 1, na.rm = TRUE) > 0]
              # extraer "choice code" después del primer punto
              # (si es var_recod.algo también cae aquí porque el patrón ya filtró por ^var(\.|_recod\.)
              choice_codes <- sub(paste0("^", v, "(_recod)?\\."), "", cols_on)
              vals_obs <- unique(choice_codes)
            }
          }
        }

        # si NO aparecen, se eliminan del layout
        if (length(vals_obs)) {
          keep_perd <- intersect(cod_perd, vals_obs)
        } else {
          keep_perd <- character(0)
        }

        # mantener siempre los no-perdidos + los perdidos observados
        es_perd <- opts_df$Código %in% cod_perd
        opts_df <- opts_df[!es_perd | (opts_df$Código %in% keep_perd), , drop = FALSE]
      }

      DT::datatable(
        opts_df,
        rownames = FALSE,
        options = list(
          paging    = FALSE,
          searching = FALSE,
          info      = FALSE,
          language  = list(
            search      = "Buscar:",
            zeroRecords = "Sin resultados"
          )
        )
      )
    })

    # ======================
    # TAB 2 — BASE DE DATOS
    # ======================

    # columnas por sección (solo secciones, sin técnicas)
    vars_data_por_seccion <- lapply(secciones_limpias, function(v) {
      intersect(v, vars_data_visibles)
    })
    vars_data_por_seccion <- vars_data_por_seccion[vapply(vars_data_por_seccion, length, integer(1)) > 0]

    # tabla base: muestra TODAS las columnas de la sección
    data_base_filtrada <- shiny::reactive({
      sec  <- input$data_seccion
      cols <- vars_data_por_seccion[[sec]] %||% character(0)

      if (!length(cols)) cols <- head(vars_data_visibles, 10)

      data[, cols, drop = FALSE]
    })

    data_base_vista <- shiny::reactive({
      df <- data_base_filtrada()

      use_labels <- isTRUE(input$vista_etiquetas)

      if (use_labels) {
        df <- .to_labels_df(df)

        # Cambiar headers a etiquetas
        cn <- vapply(names(df), function(v) {
          lab <- attr(data[[v]], "label", exact = TRUE)
          if (!is.null(lab) && nzchar(as.character(lab))) as.character(lab) else v
        }, character(1))
        names(df) <- cn
      }

      df
    })

    output$tabla_data <- DT::renderDataTable({

      df <- data_base_vista()
      use_labels <- isTRUE(input$vista_etiquetas)

      # ancho fijo por modo
      col_w <- if (use_labels) 220 else 120  # ajusta a gusto

      cb_txt <- paste0(
        "function(settings) {
  var api = this.api();
  var thead = $(api.table().header());

  function escapeRegex(s) {
    return s.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&');
  }

  // recrear fila de filtros solo si no existe
  if ($(thead).find('tr').length < 2) {
    var filterRow = $('<tr class=\"dt-filter-row\">').appendTo(thead);

    api.columns().every(function() {
      var col = this;
      var th  = $('<th>').appendTo(filterRow);

      var uniq = col.data().unique().toArray()
        .filter(function(x){ return x !== null && x !== undefined && x !== ''; });

      uniq.sort();

      if (uniq.length <= 20) {

        var sel = $('<select multiple></select>')
          .css({
            'width':'100%',
            'font-size':'11px',
            'box-sizing':'border-box'
          })
          .appendTo(th);

        $('<option></option>').attr('value','__ALL__').text('(Todos)').appendTo(sel);

        uniq.forEach(function(v){
          $('<option></option>').attr('value', v).text(v).appendTo(sel);
        });

        var $sel = $(sel).selectize({
          plugins: ['remove_button'],
          maxItems: null,
          closeAfterSelect: false,
          hideSelected: false,
          placeholder: 'Filtrar...',
          dropdownParent: 'body',

          render: {
            option: function(item, escape) {
              var label = item.text || item.value;
              var isAll = (item.value === '__ALL__');
              return '<div style=\"display:flex;align-items:center;gap:8px;\">'
                + '<input type=\"checkbox\" style=\"pointer-events:none;\"/>'
                + '<span>' + escape(label) + '</span>'
                + (isAll ? '<span style=\"margin-left:auto;color:#5f6b7a;font-weight:700;\">*</span>' : '')
                + '</div>';
            },
            item: function(item, escape) {
              return '<div>' + escape(item.text || item.value) + '</div>';
            }
          },

          onChange: function(vals) {
            vals = vals || [];

            // ALL o vacío => limpiar filtro
            if (vals.length === 0 || vals.indexOf('__ALL__') >= 0) {
              col.search('').draw();
              return;
            }

            // OR exacto
            var rx = '^(' + vals.map(escapeRegex).join('|') + ')$';
            col.search(rx, true, false).draw();
          }
        });

        var inst = $sel[0].selectize;
        var $ctrl = $(inst.$control);
        $ctrl.css({
          'border':'1px solid #e6e9f2',
          'border-radius':'10px',
          'min-height':'30px',
          'padding':'2px 4px',
          'box-shadow':'none'
        });

      } else {

        var inp = $('<input type=\"text\" placeholder=\"Filtrar\"/>')
          .css({
            'width':'100%',
            'border':'1px solid #e6e9f2',
            'border-radius':'10px',
            'padding':'6px 8px',
            'font-size':'11px',
            'box-sizing':'border-box'
          })
          .appendTo(th);

        inp.on('keyup change clear', function() {
          if (col.search() !== this.value) {
            col.search(this.value).draw();
          }
        });
      }
    });
  }
}"
      )

      cb <- DT::JS(cb_txt)

      DT::datatable(
        df,
        rownames   = FALSE,
        extensions = c("Scroller"),
        options = list(
          destroy    = TRUE,
          serverSide = FALSE,
          autoWidth  = FALSE,

          # ancho fijo igual para todas las columnas
          columnDefs = list(list(width = paste0(col_w, "px"), targets = "_all")),

          deferRender = TRUE,
          scrollX     = TRUE,
          scrollY     = 560,
          scroller    = TRUE,
          pageLength  = 15,
          lengthMenu  = c(10, 15, 25, 50),
          initComplete = cb,
          language = list(
            lengthMenu   = "Mostrando _MENU_ registros",
            search       = "Buscar:",
            info         = "Mostrando _START_ a _END_ de _TOTAL_ registros",
            infoEmpty    = "Mostrando 0 a 0 de 0 registros",
            infoFiltered = "(filtrado de _MAX_ registros)",
            zeroRecords  = "Sin resultados",
            paginate     = list(previous = "Anterior", `next` = "Siguiente")
          )
        )
      )
    })


  }

  shiny::shinyApp(ui = ui, server = server)
}
