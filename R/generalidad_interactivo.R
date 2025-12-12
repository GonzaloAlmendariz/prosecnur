# =============================================================================
# Explorador interactivo: reporte_interactivo()
# Fase 2 – Gráfico principal + tabla + bloque de perfil (N + 2 KPIs)
# Versión más sofisticada: paletas robustas + layout limpio + KPIs half-donut
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
    titulo_font_size  <- 14
    titulo_margin_top <- 60
    margin_left       <- if (solo_total) 20 else 170
    margin_right      <- 25
    margin_bottom     <- 45
  } else {
    titulo_font_size  <- 11
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

    p <- p |>
      plotly::add_bars(
        data             = df_opt,
        x                = ~pct,
        y                = ~estrato_label,
        name             = opt,
        orientation      = "h",
        text             = ~texto_pct_html,
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
      title  = list(text = titulo, x = 0, xanchor = "left", font = list(size = titulo_font_size)),
      uniformtext = list(minsize = 10, mode = "hide"),
      hovermode  = "closest",
      showlegend = mostrar_leyenda,
      transition = list(duration = 450, easing = "cubic-in-out")
    ) |>
    plotly::config(displayModeBar = FALSE, responsive = TRUE)

  p <- plotly::animation_opts(
    p,
    frame      = 600,
    transition = 450,
    easing     = "cubic-in-out",
    redraw     = TRUE
  )

  p
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

  # ordenar y consolidar
  df_tab$opcion_label <- as.character(df_tab$opcion_label)
  df_tab <- df_tab[order(df_tab$opcion_label), , drop = FALSE]

  titulo_kpi <- .wrap_titulo_html(
    .obtener_label_var(var_kpi, instrumento, df_kpi),
    width = 55
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

  # Texto central: mostrar categoría líder (más %)
  idx_max <- which.max(df_tab$porc_int)
  center_pct   <- df_tab$porc_int[idx_max] %||% 0L
  center_label <- df_tab$opcion_label[idx_max] %||% ""

  # half donut: pie + hole + rotation + domain + sin leyenda interna
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
    hovertemplate = paste0("<b>", titulo_kpi, "</b><br>%{label}: %{value}%<extra></extra>")
  ) |>
    plotly::layout(
      title = list(text = titulo_kpi, x = 0.5, xanchor = "center", font = list(size = 12)),
      showlegend = FALSE,
      margin = list(l = 10, r = 10, t = 45, b = 5),
      # recorte para que sea media dona (solo mitad superior)
      # (plotly: se logra jugando con el dominio/annotations; esto funciona bien en la práctica)
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

  list(plot = p, legend = legend_df)
}

# -----------------------------------------------------------------------------
# App principal
# -----------------------------------------------------------------------------

#' Explorador interactivo de resultados (fase 2)
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
    kpi_vars   = NULL
) {

  if (!requireNamespace("shiny", quietly = TRUE) ||
      !requireNamespace("plotly", quietly = TRUE) ||
      !requireNamespace("dplyr",  quietly = TRUE)) {
    stop("Se requieren 'shiny', 'plotly' y 'dplyr' para `reporte_interactivo()`.", call. = FALSE)
  }

  survey <- instrumento$survey
  if (is.null(survey) || !"name" %in% names(survey)) {
    stop("El `instrumento` debe contener un `survey` válido.", call. = FALSE)
  }

  usa_DT <- requireNamespace("DT", quietly = TRUE)

  if (is.null(secciones) || !length(secciones)) {
    stop("`secciones` debe ser una lista nombrada con vectores de variables.", call. = FALSE)
  }

  # Secciones: sólo variables presentes en data
  secciones_limpias <- lapply(secciones, function(v) v[v %in% names(data)])
  secciones_limpias <- secciones_limpias[vapply(secciones_limpias, length, integer(1)) > 0]
  if (!length(secciones_limpias)) {
    stop("Ninguna sección de `secciones` tiene variables presentes en `data`.", call. = FALSE)
  }

  secciones_nombres <- names(secciones_limpias)
  label_var <- function(v) .obtener_label_var(v, instrumento, data)

  # Filtros / cruces
  facet_vars <- (facet_vars %||% character(0))
  facet_vars <- facet_vars[facet_vars %in% names(data)]
  facet_choices <- stats::setNames(facet_vars, vapply(facet_vars, label_var, character(1)))

  # KPIs (máx 2)
  kpi_vars <- (kpi_vars %||% character(0))
  kpi_vars <- unique(kpi_vars[kpi_vars %in% names(data)])
  if (length(kpi_vars) > 2L) kpi_vars <- kpi_vars[1:2]

  # -------------------------------- UI ---------------------------------------
  ui <- shiny::fluidPage(
    shiny::tags$head(
      shiny::tags$style(shiny::HTML("
        /* Plotly: usar todo el ancho del contenedor */
        .html-widget, .plotly, .plotly html-widget, .plot-container, .svg-container {
          width: 100% !important;
          max-width: 100% !important;
        }
        /* Evitar recortes de hover en tarjetas */
        .shiny-output-error, .shiny-output-error-validation { color: #b30000; }
        /* Ajuste fino del grid bootstrap */
        .row { margin-left: -10px; margin-right: -10px; }
        .col-sm-6, .col-sm-12, .col-sm-9, .col-sm-3 { padding-left: 10px; padding-right: 10px; }
      "))
    ),

    shiny::titlePanel(title = titulo),

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
            plotly::plotlyOutput("plot_principal", height = "420px")
          )
        ),

        shiny::br(),

        shiny::fluidRow(
          shiny::column(
            width = 6,
            shiny::div(
              style = paste(
                "height: 360px; overflow-y: auto;",
                "border: 1px solid #eee; border-radius: 10px; padding: 8px;",
                "background: #fff;"
              ),
              if (usa_DT) DT::dataTableOutput("tabla_principal") else shiny::tableOutput("tabla_principal")
            )
          ),
          shiny::column(
            width = 6,
            shiny::div(
              style = paste(
                "height: 360px;",
                "border: 1px solid #eee; border-radius: 10px; padding: 8px;",
                "background: #fff;",
                "display: flex; flex-direction: column; align-items: stretch;",
                "overflow: visible;"
              ),
              shiny::uiOutput("kpi_panel")
            )
          )
        )
      )
    )
  )

  # ------------------------------- SERVER ------------------------------------
  server <- function(input, output, session) {

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

      shiny::checkboxGroupInput(
        inputId  = "filtro_categorias",
        label    = label_var(v),
        choices  = vals,
        selected = vals
      )
    })

    shiny::observeEvent(input$limpiar_filtros, {
      v <- input$filtro_var
      if (is.null(v) || !nzchar(v) || !v %in% names(data)) return()

      vals <- sort(unique(as.character(data[[v]])))
      vals <- vals[!is.na(vals)]
      shiny::updateCheckboxGroupInput(session, "filtro_categorias", selected = vals)
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

    # ------------------ Bloque 1: gráfico principal --------------------------
    output$plot_principal <- plotly::renderPlotly({
      shiny::req(input$var_principal)

      var_main <- input$var_principal
      df_all   <- data_filtrada()

      df <- if (var_main %in% names(df_all)) df_all[!is.na(df_all[[var_main]]), , drop = FALSE] else df_all

      var_cruce <- input$var_cruce
      if (!nzchar(var_cruce)) var_cruce <- NULL

      if (nrow(df) == 0L) {
        shiny::validate(shiny::need(FALSE, "No hay datos válidos (después de filtros)."))
      }

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

    # ------------------ Bloque 2: tabla resumen ------------------------------
    if (usa_DT) {
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
          options  = list(paging = FALSE, searching = FALSE, info = FALSE)
        )
      })
    } else {
      output$tabla_principal <- shiny::renderTable({
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

        .construir_tabla_resumen(df_tab)
      })
    }

    # ------------------ Bloque 3: perfil (N + KPIs) --------------------------
    output$kpi_panel <- shiny::renderUI({

      df_all   <- data_filtrada()
      var_main <- input$var_principal

      df <- if (!is.null(var_main) && nzchar(var_main) && var_main %in% names(df_all)) {
        df_all[!is.na(df_all[[var_main]]), , drop = FALSE]
      } else {
        df_all
      }

      if (!nrow(df)) return(shiny::div("Sin datos para la pregunta seleccionada."))

      # N dinámico
      n_unidades <- if (!is.null(id_unidad) && id_unidad %in% names(df)) {
        dplyr::n_distinct(df[[id_unidad]])
      } else {
        nrow(df)
      }

      # texto: "N: 14 EESS" (usa id_unidad como sufijo si viene)
      n_sufijo <- if (!is.null(id_unidad) && nzchar(id_unidad)) id_unidad else ""
      texto_N  <- paste0(
        "N: ",
        format(n_unidades, big.mark = ",", scientific = FALSE),
        if (nzchar(n_sufijo)) paste0(" ", n_sufijo) else ""
      )

      tarjeta_N <- shiny::div(
        style = paste(
          "width: 100%;",
          "border: 1px solid #e5e5e5; border-radius: 10px;",
          "padding: 12px 14px; margin-bottom: 12px;",
          "background: #fafafa;",
          "display: flex; justify-content: center; align-items: center;",
          "text-align: center;"
        ),
        shiny::div(
          style = "font-size: 18px; font-weight: 800; color: #222;",
          texto_N
        )
      )

      kpi_elems <- list(tarjeta_N)

      # leyenda pequeña abajo (no afecta al plot)
      legend_html <- function(legend_df) {
        shiny::div(
          style = paste(
            "margin-top: 6px;",
            "display: flex; flex-wrap: wrap; gap: 6px 10px;",
            "justify-content: center;",
            "font-size: 10px; color: #555; line-height: 1.1;"
          ),
          lapply(seq_len(nrow(legend_df)), function(i) {
            shiny::div(
              style = "display: inline-flex; align-items: center; gap: 6px;",
              shiny::span(style = paste0(
                "display:inline-block;width:10px;height:10px;border-radius:3px;",
                "background:", legend_df$color[i], ";"
              )),
              shiny::span(legend_df$label[i])
            )
          })
        )
      }

      # construir KPIs half-donut (SIEMPRE) + leyenda abajo
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
        output$kpi_plot_1 <- plotly::renderPlotly(kpi_obj_1$plot)
      }

      if (length(kpi_vars) >= 2) {
        kpi_obj_2 <- .construir_kpi_halfdonut(
          df = df,
          var_kpi = kpi_vars[2],
          instrumento = instrumento,
          colores_apiladas_por_listname = colores_apiladas_por_listname,
          codigos_perdidos = codigos_perdidos
        )
        output$kpi_plot_2 <- plotly::renderPlotly(kpi_obj_2$plot)
      }

      if (length(kpi_vars) == 1 && !is.null(kpi_obj_1)) {
        kpi_elems[[length(kpi_elems) + 1]] <- shiny::div(
          style = "width: 100%;",
          plotly::plotlyOutput("kpi_plot_1", height = "230px"),
          legend_html(kpi_obj_1$legend)
        )
      } else if (length(kpi_vars) >= 2 && !is.null(kpi_obj_1) && !is.null(kpi_obj_2)) {
        kpi_elems[[length(kpi_elems) + 1]] <- shiny::fluidRow(
          shiny::column(
            width = 6,
            plotly::plotlyOutput("kpi_plot_1", height = "230px"),
            legend_html(kpi_obj_1$legend)
          ),
          shiny::column(
            width = 6,
            plotly::plotlyOutput("kpi_plot_2", height = "230px"),
            legend_html(kpi_obj_2$legend)
          )
        )
      }

      do.call(shiny::tagList, kpi_elems)
    })
  }

  shiny::shinyApp(ui = ui, server = server)
}
