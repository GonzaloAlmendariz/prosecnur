# =============================================================================
# Tab 1: Resumen (UI + server) — v3.9 debug KPI + color fix
# -----------------------------------------------------------------------------
# - Títulos/subtítulos y % externo SM usan color_primario
# - Texto del nombre de cada opción SM usa color_texto
# - SM 0% no muestra etiqueta
# - Logs ampliados para rastrear por qué no se construyen los KPI
# =============================================================================
#' @keywords internal
#' @noRd

.ui_tab_resumen <- function(ctx) {

  shiny::sidebarLayout(
    shiny::sidebarPanel(
      width = 3,

      shiny::h3("Resumen"),

      shiny::selectInput(
        inputId  = "seccion",
        label    = "Sección",
        choices  = stats::setNames(ctx$secciones_nombres, ctx$secciones_nombres),
        selected = ctx$secciones_nombres[1]
      ),

      shiny::hr(),

      shiny::h3("Filtros"),

      shiny::selectInput(
        inputId  = "filtro_var",
        label    = "Variable de filtro",
        choices  = c("Ninguno" = "", ctx$facet_choices),
        selected = ""
      ),

      shiny::uiOutput("filtro_categorias_ui"),

      shiny::actionButton(
        inputId = "limpiar_filtros",
        label   = "Limpiar filtros"
      ),

      shiny::hr(),

      shiny::div(
        class = "cardbox",
        shiny::div(
          class = "cardbox-header",
          shiny::div(class = "cardbox-title", "Perfil de la muestra")
        ),
        shiny::uiOutput("kpi_panel")
      ),

      shiny::div(style = "height: 24px;")
    ),

    shiny::mainPanel(
      width = 9,

      shiny::fluidRow(
        shiny::column(
          width = 12,
          shiny::div(
            class = "cardbox",
            shiny::div(
              class = "cardbox-header",
              shiny::div(class = "cardbox-title", shiny::uiOutput("section_title_ui"))
            ),
            shiny::uiOutput("section_summary_ui")
          )
        )
      ),

      shiny::div(style = "height: 48px;")
    )
  )
}

#' @keywords internal
#' @noRd
.server_tab_resumen <- function(ctx, input, output, session) {

  data        <- ctx$data
  instrumento <- ctx$instrumento

  MAX_SO_ROWS <- 16L
  BAR_HEIGHT  <- 64
  PCT_FSIZE   <- 13

  `%||%` <- get0("%||%", ifnotfound = function(x, y) if (!is.null(x)) x else y)

  # ---------------------------------------------------------------------------
  # LOG helper
  # ---------------------------------------------------------------------------
  .log_resumen <- function(...) {
    msg <- paste(..., collapse = "")
    message("[tab_resumen] ", msg)
  }

  .safe_chr <- function(x, max_n = 80) {
    if (is.null(x)) return("NULL")
    x <- as.character(x)
    if (!length(x)) return("")
    if (length(x) > max_n) {
      x <- c(x[seq_len(max_n)], paste0("...(+", length(x) - max_n, " más)"))
    }
    paste(x, collapse = ", ")
  }

  # ---------------------------------------------------------------------------
  # Tema visual
  # ---------------------------------------------------------------------------
  theme_default <- if (exists("reporte_interactivo_theme_default", mode = "function")) {
    reporte_interactivo_theme_default()
  } else {
    list(
      color_primario      = "#002457",
      color_fondo_app     = "#f5f6fa",
      color_borde         = "#e6e9f2",
      color_texto         = "#1f2933",
      color_texto_suave   = "#5f6b7a",
      color_superficie    = "#ffffff",
      color_superficie_2  = "#fafbff",
      color_header_tabla  = "#f1f3f9"
    )
  }

  theme_app <- theme_default
  if (!is.null(ctx$theme_app) && is.list(ctx$theme_app)) {
    nm <- intersect(names(ctx$theme_app), names(theme_app))
    if (length(nm)) theme_app[nm] <- ctx$theme_app[nm]
  }

  # Reglas visuales pedidas:
  # - barras SM: color_primario
  # - % externo SM: color_primario
  # - títulos/subtítulos: color_primario
  # - nombre de cada opción SM: color_texto
  SM_COLOR_YES        <- theme_app$color_primario
  SM_COLOR_BG         <- theme_app$color_superficie_2
  SM_TEXT_OUT         <- theme_app$color_primario
  SM_SUBTITLE         <- theme_app$color_primario
  SM_OPTION_TEXT      <- theme_app$color_texto
  MSG_COLOR           <- theme_app$color_texto_suave

  .log_resumen(
    "Theme -> primario=", SM_COLOR_YES,
    " | bg_sm=", SM_COLOR_BG,
    " | texto_out=", SM_TEXT_OUT,
    " | sm_option_text=", SM_OPTION_TEXT
  )

  # ---------------------------------------------------------------------------
  # Helpers locales
  # ---------------------------------------------------------------------------
  .wrap_titulo_html <- get0(
    ".wrap_titulo_html",
    ifnotfound = function(txt, width = 110) {
      if (!requireNamespace("stringr", quietly = TRUE)) return(as.character(txt))
      if (is.null(txt)) return("")
      lineas <- stringr::str_wrap(as.character(txt), width = width)
      paste(lineas, collapse = "<br>")
    }
  )

  .get_label_col_safe_local <- function(df) {
    if (is.null(df)) return(NULL)
    if ("label" %in% names(df)) return("label")
    lab_candidates <- grep("^label(::|$)", names(df), value = TRUE)
    if (length(lab_candidates)) return(lab_candidates[1])
    NULL
  }

  .obtener_label_var <- get0(
    ".obtener_label_var",
    ifnotfound = function(var, instrumento, data = NULL) {
      var <- trimws(as.character(var)[1])
      surv <- instrumento$survey

      if (!is.null(surv) && "name" %in% names(surv)) {
        label_col <- .get_label_col_safe_local(surv)
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
        if (!is.null(vl) && nzchar(as.character(vl))) return(as.character(vl))
      }

      as.character(var)
    }
  )

  .get_choice_label_col <- function(ch) {
    if (is.null(ch)) return(NULL)
    if ("label" %in% names(ch)) return("label")
    lab_candidates <- grep("^label(::|$)", names(ch), value = TRUE)
    if (length(lab_candidates)) return(lab_candidates[1])
    NULL
  }

  .resolver_paleta_var_safe <- function(var, opcion_levels) {
    f <- get0(".resolver_paleta_var", mode = "function", ifnotfound = NULL)
    if (is.null(f)) {
      f <- get0("resolver_paleta_var", mode = "function", ifnotfound = NULL)
    }

    .log_resumen(
      "Paleta -> var=", var,
      " | helper=", if (is.null(f)) "NO" else "SI",
      " | niveles=", .safe_chr(opcion_levels)
    )

    if (!is.null(f)) {
      pal <- tryCatch(
        f(
          var = var,
          instrumento = instrumento,
          colores_apiladas_por_listname = ctx$colores_apiladas_por_listname,
          opcion_levels = opcion_levels
        ),
        error = function(e) {
          .log_resumen("Paleta ERROR en var=", var, " -> ", conditionMessage(e))
          NULL
        }
      )
      if (!is.null(pal) && length(pal)) {
        .log_resumen("Paleta OK var=", var, " -> ", .safe_chr(paste(names(pal), pal, sep = "=")))
        return(pal)
      }
    }

    out <- grDevices::hcl.colors(max(3L, length(opcion_levels)), "Blues")
    out <- out[seq_len(length(opcion_levels))]
    names(out) <- opcion_levels
    .log_resumen("Paleta fallback var=", var, " -> ", .safe_chr(paste(names(out), out, sep = "=")))
    out
  }

  .preparar_tabla_kpi_safe <- function(df, var, codigos_perdidos = NULL) {

    .log_resumen("KPI prep -> var=", var)

    if (!var %in% names(df)) {
      .log_resumen("KPI prep FAIL -> var no existe en df: ", var)
      return(NULL)
    }

    surv <- instrumento$survey %||% NULL
    ch   <- instrumento$choices %||% NULL
    label_col <- .get_choice_label_col(ch)

    x <- as.character(df[[var]])
    x <- x[!is.na(x) & nzchar(x) & x != "NA"]

    .log_resumen("KPI prep -> casos iniciales válidos=", length(x))

    if (!is.null(codigos_perdidos) && length(codigos_perdidos)) {
      x <- x[!(x %in% as.character(codigos_perdidos))]
      .log_resumen("KPI prep -> tras excluir perdidos=", length(x), " | perdidos=", .safe_chr(codigos_perdidos))
    }

    if (!length(x)) {
      .log_resumen("KPI prep FAIL -> no quedan casos para ", var)
      return(NULL)
    }

    map_code_to_label <- NULL

    if (!is.null(surv) &&
        all(c("name", "list_name") %in% names(surv)) &&
        !is.null(ch) &&
        all(c("list_name", "name") %in% names(ch)) &&
        !is.null(label_col) && label_col %in% names(ch)) {

      i <- which(!is.na(surv$name) & surv$name == var)[1]
      if (!is.na(i)) {
        ln <- as.character(surv$list_name[i])
        .log_resumen("KPI prep -> list_name=", ln, " | label_col=", label_col)
        if (!is.na(ln) && nzchar(ln)) {
          ch_v <- ch[ch$list_name == ln, , drop = FALSE]
          .log_resumen("KPI prep -> nrow choices=", nrow(ch_v))
          if (nrow(ch_v)) {
            map_code_to_label <- stats::setNames(
              as.character(ch_v[[label_col]]),
              as.character(ch_v$name)
            )
          }
        }
      }
    }

    if (is.null(map_code_to_label)) {
      labs <- attr(df[[var]], "labels", exact = TRUE)
      if (!is.null(labs) && length(labs) > 0) {
        .log_resumen("KPI prep -> usando labels de atributo para ", var)
        map_code_to_label <- stats::setNames(
          as.character(unname(labs)),
          as.character(names(labs))
        )
      }
    }

    if (is.null(map_code_to_label)) {
      vals <- sort(unique(x))
      .log_resumen("KPI prep -> sin diccionario; usando valores crudos")
      map_code_to_label <- stats::setNames(vals, vals)
    }

    tab <- as.data.frame(table(x), stringsAsFactors = FALSE)
    names(tab) <- c("code", "n")
    tab$n <- as.numeric(tab$n)

    tab$label <- unname(map_code_to_label[tab$code])
    tab$label[is.na(tab$label) | tab$label == ""] <- tab$code[is.na(tab$label) | tab$label == ""]

    orden <- unique(unname(map_code_to_label))
    orden <- orden[!is.na(orden) & nzchar(orden)]
    if (length(orden)) {
      tab$label <- factor(tab$label, levels = orden)
      tab <- tab[order(tab$label), , drop = FALSE]
      tab$label <- as.character(tab$label)
    }

    tab$pct <- tab$n / sum(tab$n)

    .log_resumen(
      "KPI prep OK -> var=", var,
      " | filas=", nrow(tab),
      " | labels=", .safe_chr(tab$label),
      " | n=", .safe_chr(tab$n)
    )

    tab
  }

  .construir_kpi_halfdonut_safe <- function(df, var_kpi) {

    .log_resumen("KPI build -> INICIO var=", var_kpi)

    if (!requireNamespace("plotly", quietly = TRUE)) {
      .log_resumen("KPI build FAIL -> plotly no disponible")
      return(NULL)
    }
    if (!var_kpi %in% names(df)) {
      .log_resumen("KPI build FAIL -> var no existe en df: ", var_kpi)
      return(NULL)
    }

    # 1) Intentar helper original real
    f <- get0(".construir_kpi_halfdonut", mode = "function", ifnotfound = NULL)
    if (is.null(f)) {
      f <- get0("construir_kpi_halfdonut", mode = "function", ifnotfound = NULL)
    }

    .log_resumen("KPI build -> helper original encontrado=", if (is.null(f)) "NO" else "SI")

    if (!is.null(f)) {
      out <- tryCatch(
        f(
          df = df,
          var_kpi = var_kpi,
          instrumento = instrumento,
          colores_apiladas_por_listname = ctx$colores_apiladas_por_listname,
          codigos_perdidos = ctx$codigos_perdidos
        ),
        error = function(e) {
          .log_resumen("KPI helper original ERROR var=", var_kpi, " -> ", conditionMessage(e))
          NULL
        }
      )

      if (!is.null(out)) {
        .log_resumen(
          "KPI helper original retornó lista? ", is.list(out),
          " | plot null? ", is.null(out$plot),
          " | legend null? ", is.null(out$legend),
          " | title null? ", is.null(out$title_html)
        )
      }

      if (!is.null(out) &&
          is.list(out) &&
          !is.null(out$plot) &&
          !is.null(out$legend) &&
          !is.null(out$title_html)) {
        .log_resumen("KPI build OK via helper original -> var=", var_kpi)
        return(out)
      }
    }

    # 2) Fallback robusto
    .log_resumen("KPI build -> usando fallback para var=", var_kpi)

    tab <- .preparar_tabla_kpi_safe(
      df = df,
      var = var_kpi,
      codigos_perdidos = ctx$codigos_perdidos
    )

    if (is.null(tab) || !nrow(tab)) {
      .log_resumen("KPI build FAIL fallback -> tabla nula/vacía para ", var_kpi)
      return(NULL)
    }

    titulo_kpi <- .wrap_titulo_html(
      .obtener_label_var(var_kpi, instrumento, df),
      width = 45
    )

    opcion_levels <- as.character(tab$label)
    paleta <- .resolver_paleta_var_safe(var_kpi, opcion_levels = opcion_levels)

    legend_df <- data.frame(
      label = opcion_levels,
      color = unname(paleta[opcion_levels]),
      stringsAsFactors = FALSE
    )

    .log_resumen(
      "KPI fallback -> labels=", .safe_chr(opcion_levels),
      " | colores=", .safe_chr(unname(paleta[opcion_levels]))
    )

    p <- tryCatch(
      plotly::plot_ly(
        data   = tab,
        labels = ~label,
        values = ~n,
        type   = "pie",
        hole   = 0.68,
        direction = "clockwise",
        rotation  = 180,
        sort      = FALSE,
        textinfo  = "none",
        marker    = list(colors = unname(paleta[opcion_levels])),
        hovertemplate = "%{label}: %{percent}<extra></extra>"
      ) |>
        plotly::layout(
          title = NULL,
          showlegend = FALSE,
          margin = list(l = 10, r = 10, t = 10, b = 5)
        ) |>
        plotly::config(displayModeBar = FALSE, responsive = TRUE),
      error = function(e) {
        .log_resumen("KPI fallback plot ERROR var=", var_kpi, " -> ", conditionMessage(e))
        NULL
      }
    )

    if (is.null(p)) {
      .log_resumen("KPI build FAIL -> plot nulo en fallback para ", var_kpi)
      return(NULL)
    }

    .log_resumen("KPI build OK via fallback -> var=", var_kpi)

    list(
      plot = p,
      legend = legend_df,
      title_html = titulo_kpi
    )
  }

  # ---------------------------------------------------------------------------
  # Filtros
  # ---------------------------------------------------------------------------
  output$filtro_categorias_ui <- shiny::renderUI({
    v <- input$filtro_var
    if (is.null(v) || !nzchar(v) || !v %in% names(data)) return(NULL)

    vals <- sort(unique(as.character(data[[v]])))
    vals <- vals[!is.na(vals)]
    if (!length(vals)) return(NULL)

    surv <- instrumento$survey
    ch   <- instrumento$choices %||% NULL
    label_col <- .get_choice_label_col(ch)

    labels_vals <- vals

    if (!is.null(surv) && all(c("name", "list_name") %in% names(surv)) &&
        !is.null(ch)   && all(c("list_name", "name") %in% names(ch)) &&
        !is.null(label_col) && label_col %in% names(ch)) {

      ln <- .get_list_name_safe(surv, v)
      if (!is.na(ln) && nzchar(ln)) {
        ch_v <- ch[ch$list_name == ln, , drop = FALSE]
        if (nrow(ch_v)) {
          map_code_to_label <- stats::setNames(as.character(ch_v[[label_col]]), as.character(ch_v$name))
          labels_vals <- unname(map_code_to_label[vals])
          labels_vals[is.na(labels_vals) | labels_vals == ""] <- vals[is.na(labels_vals) | labels_vals == ""]
        }
      }
    }

    shiny::checkboxGroupInput(
      inputId  = "filtro_categorias",
      label    = ctx$label_var(v),
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

  # ---------------------------------------------------------------------------
  # Helpers tipo / detección de SM
  # ---------------------------------------------------------------------------
  .has_var_or_dummies <- function(df, var) {
    if (!is.data.frame(df)) return(FALSE)
    if (var %in% names(df)) return(TRUE)
    var_esc <- gsub("([\\W])", "\\\\\\1", var)
    any(grepl(paste0("^", var_esc, "[/\\.]"), names(df)))
  }

  tipo_pregunta <- function(var, survey = NULL, sm_vars_force = NULL, df = NULL) {
    if (!is.null(sm_vars_force) && var %in% sm_vars_force) return("sm")
    if (!is.null(survey) && any(survey$name == var)) {
      tipos <- unique(na.omit(survey$type[survey$name == var]))
      tipos <- tolower(as.character(tipos))
      if (any(grepl("^select_multiple(\\s|$)", tipos))) return("sm")
      if (any(grepl("^select_one(\\s|$)", tipos))) return("so")
    }
    if (!is.null(df) && .has_var_or_dummies(df, var) && !(var %in% names(df))) return("sm")
    "so"
  }

  get_categorias <- function(var, df, survey = NULL, orders_list = NULL, opciones_excluir = NULL) {

    x <- if (var %in% names(df)) df[[var]] else NULL
    lab_attr <- if (!is.null(x)) attr(x, "labels", exact = TRUE) else NULL

    ln <- NA_character_
    if (!is.null(survey) && all(c("name", "list_name") %in% names(survey))) {
      ln <- .get_list_name_safe(survey, var)
    }

    codes  <- character(0)
    labels <- character(0)

    obj <- NULL
    if (!is.null(orders_list)) {
      if (var %in% names(orders_list)) obj <- orders_list[[var]]
      else if (!is.na(ln) && ln %in% names(orders_list)) obj <- orders_list[[ln]]
    }

    if (!is.null(obj)) {
      codes  <- as.character(obj$names)
      labels <- as.character(obj$labels)
    } else if (!is.null(lab_attr) && length(lab_attr) > 0) {
      codes  <- names(lab_attr)
      labels <- as.character(unname(lab_attr))
    } else if (!is.null(x)) {
      codes  <- sort(unique(na.omit(as.character(x))))
      labels <- codes
    }

    ok <- !is.na(codes) & nzchar(codes)
    codes  <- codes[ok]
    labels <- labels[ok]

    if (!is.null(opciones_excluir) && length(opciones_excluir) > 0) {
      ok2 <- !(labels %in% opciones_excluir)
      codes  <- codes[ok2]
      labels <- labels[ok2]
    }

    list(codes = codes, labels = labels, list_name = ln)
  }

  # ---------------------------------------------------------------------------
  # Plot SO
  # ---------------------------------------------------------------------------
  .plot_so_total <- function(df, var, paleta_colores) {

    if (!var %in% names(df)) {
      return(
        plotly::plot_ly(height = BAR_HEIGHT) |>
          plotly::layout(annotations = list(list(text = "Sin variable.", showarrow = FALSE))) |>
          plotly::config(displayModeBar = FALSE, responsive = TRUE)
      )
    }

    x <- as.character(df[[var]])
    x <- x[!is.na(x) & nzchar(x) & x != "NA"]
    if (!length(x)) {
      return(
        plotly::plot_ly(height = BAR_HEIGHT) |>
          plotly::layout(annotations = list(list(text = "Sin datos.", showarrow = FALSE))) |>
          plotly::config(displayModeBar = FALSE, responsive = TRUE)
      )
    }

    tab <- as.data.frame(table(x), stringsAsFactors = FALSE)
    names(tab) <- c("code", "n")
    tab$n   <- as.numeric(tab$n)
    tab$pct <- tab$n / sum(tab$n)

    map_code_to_label <- NULL
    labs <- attr(df[[var]], "labels", exact = TRUE)
    if (!is.null(labs) && length(labs) > 0) {
      map_code_to_label <- stats::setNames(as.character(unname(labs)), as.character(names(labs)))
    }

    tab$label <- if (!is.null(map_code_to_label)) {
      out <- unname(map_code_to_label[tab$code])
      out[is.na(out) | out == ""] <- tab$code[is.na(out) | out == ""]
      out
    } else {
      tab$code
    }

    if (!is.null(paleta_colores) && !is.null(names(paleta_colores)) &&
        all(tab$label %in% names(paleta_colores))) {
      tab$label <- factor(tab$label, levels = names(paleta_colores))
      tab <- tab[order(tab$label), , drop = FALSE]
    } else {
      tab <- tab[order(tab$pct, decreasing = TRUE), , drop = FALSE]
    }

    tab$txt <- paste0("<b>", round(100 * tab$pct, 0), "%</b>")
    tab$hover <- sprintf(
      "%s: %s%%<br>n: %s",
      as.character(tab$label),
      round(100 * tab$pct, 1),
      format(tab$n, big.mark = ",")
    )

    p <- plotly::plot_ly(height = BAR_HEIGHT)

    for (lab in as.character(tab$label)) {
      d <- tab[as.character(tab$label) == lab, , drop = FALSE]
      if (!nrow(d)) next

      col <- if (!is.null(paleta_colores) && !is.null(names(paleta_colores)) &&
                 lab %in% names(paleta_colores)) {
        unname(paleta_colores[[lab]])
      } else NULL

      p <- p |>
        plotly::add_bars(
          data             = d,
          x                = ~pct,
          y                = I("Total"),
          name             = lab,
          orientation      = "h",
          text             = ~txt,
          textposition     = "inside",
          insidetextanchor = "middle",
          textfont         = list(color = "white", size = PCT_FSIZE),
          customdata       = ~hover,
          hovertemplate    = "%{customdata}<extra></extra>",
          marker           = list(color = col, line = list(width = 0))
        )
    }

    p |>
      plotly::layout(
        barmode = "stack",
        xaxis = list(title = "", range = c(0,1), showgrid = FALSE, zeroline = FALSE,
                     showticklabels = FALSE, ticks = ""),
        yaxis = list(title = "", showgrid = FALSE, zeroline = FALSE,
                     showticklabels = FALSE, ticks = ""),
        margin = list(l = 10, r = 10, t = 0, b = 0),
        showlegend = FALSE
      ) |>
      plotly::config(displayModeBar = FALSE, responsive = TRUE)
  }

  # ---------------------------------------------------------------------------
  # Plot SM dummy fill-only
  # ---------------------------------------------------------------------------
  .plot_sm_dummy_fill <- function(df, col_dummy,
                                  col_yes = SM_COLOR_YES,
                                  col_bg  = SM_COLOR_BG,
                                  text_out_color = SM_TEXT_OUT,
                                  pct_inside_threshold = 0.05) {

    .log_resumen("SM plot -> col=", col_dummy, " | text_out_color=", text_out_color)

    if (!col_dummy %in% names(df)) {
      .log_resumen("SM plot FAIL -> dummy no existe: ", col_dummy)
      return(
        plotly::plot_ly(height = BAR_HEIGHT) |>
          plotly::layout(
            annotations = list(list(text = "Sin dummy.", showarrow = FALSE)),
            xaxis = list(visible = FALSE),
            yaxis = list(visible = FALSE),
            margin = list(l = 10, r = 10, t = 0, b = 0)
          ) |>
          plotly::config(displayModeBar = FALSE, responsive = TRUE)
      )
    }

    x <- df[[col_dummy]]

    x2 <- suppressWarnings(as.numeric(as.character(x)))
    if (all(is.na(x2)) && is.logical(x)) x2 <- as.numeric(x)

    ok <- !is.na(x2) & x2 %in% c(0, 1)
    x2 <- x2[ok]

    if (!length(x2)) {
      .log_resumen("SM plot FAIL -> sin datos válidos en ", col_dummy)
      return(
        plotly::plot_ly(height = BAR_HEIGHT) |>
          plotly::layout(
            annotations = list(list(text = "Sin datos.", showarrow = FALSE)),
            xaxis = list(visible = FALSE),
            yaxis = list(visible = FALSE),
            margin = list(l = 10, r = 10, t = 0, b = 0)
          ) |>
          plotly::config(displayModeBar = FALSE, responsive = TRUE)
      )
    }

    N     <- length(x2)
    n_yes <- sum(x2 == 1)
    pct_y <- n_yes / N

    .log_resumen(
      "SM plot -> col=", col_dummy,
      " | N=", N,
      " | n_yes=", n_yes,
      " | pct=", round(pct_y, 4)
    )

    if (pct_y == 0) {
      .log_resumen("SM plot -> pct=0; no se dibuja etiqueta")
      p <- plotly::plot_ly(height = BAR_HEIGHT) |>
        plotly::add_bars(
          x           = 1,
          y           = I("Total"),
          orientation = "h",
          marker      = list(color = col_bg, line = list(width = 0)),
          hovertemplate = paste0(
            "Sí: 0%<br>",
            "n: 0<br>",
            "N: ", format(N, big.mark = ","), "<extra></extra>"
          ),
          showlegend  = FALSE
        ) |>
        plotly::layout(
          barmode = "stack",
          xaxis = list(title = "", range = c(0,1), showgrid = FALSE, zeroline = FALSE,
                       showticklabels = FALSE, ticks = ""),
          yaxis = list(title = "", showgrid = FALSE, zeroline = FALSE,
                       showticklabels = FALSE, ticks = ""),
          margin = list(l = 10, r = 10, t = 0, b = 0),
          showlegend = FALSE
        ) |>
        plotly::config(displayModeBar = FALSE, responsive = TRUE)

      return(p)
    }

    pct_r <- 1 - pct_y

    seg <- data.frame(
      seg   = c("yes", "bg"),
      pct   = c(pct_y, pct_r),
      n_yes = n_yes,
      N     = N,
      stringsAsFactors = FALSE
    )

    pct_num <- round(100 * pct_y, 0)
    pct_txt_plain  <- paste0(pct_num, "%")
    pct_txt_inside <- paste0("<b>", pct_txt_plain, "</b>")

    seg$hover <- c(
      sprintf(
        "Sí: %s%%<br>n: %s<br>N: %s",
        round(100 * pct_y, 1),
        format(n_yes, big.mark = ","),
        format(N, big.mark = ",")
      ),
      ""
    )

    p <- plotly::plot_ly(height = BAR_HEIGHT)

    p <- p |>
      plotly::add_bars(
        data             = seg[seg$seg == "yes", , drop = FALSE],
        x                = ~pct,
        y                = I("Total"),
        orientation      = "h",
        marker           = list(color = col_yes, line = list(width = 0)),
        customdata       = ~hover,
        hovertemplate    = "%{customdata}<extra></extra>",
        showlegend       = FALSE
      )

    p <- p |>
      plotly::add_bars(
        data        = seg[seg$seg == "bg", , drop = FALSE],
        x           = ~pct,
        y           = I("Total"),
        orientation = "h",
        marker      = list(color = col_bg, line = list(width = 0)),
        hoverinfo   = "skip",
        showlegend  = FALSE
      )

    ann <- list()
    if (pct_y >= pct_inside_threshold) {
      .log_resumen("SM plot -> etiqueta INSIDE col=", col_dummy)
      ann <- list(list(
        x = pct_y / 2,
        y = "Total",
        xref = "x",
        yref = "y",
        text = pct_txt_inside,
        showarrow = FALSE,
        xanchor = "center",
        yanchor = "middle",
        align = "center",
        font = list(color = "white", size = PCT_FSIZE)
      ))
    } else {
      .log_resumen(
        "SM plot -> etiqueta OUTSIDE col=", col_dummy,
        " | texto=", pct_txt_plain,
        " | color=", col_yes
      )
      ann <- list(list(
        x = pct_y,
        y = "Total",
        xref = "x",
        yref = "y",
        text = pct_txt_plain,
        showarrow = FALSE,
        xanchor = "left",
        yanchor = "middle",
        align = "left",
        xshift = 6,
        font = list(color = col_yes, size = PCT_FSIZE)
      ))
    }

    p |>
      plotly::layout(
        barmode = "stack",
        xaxis = list(title = "", range = c(0,1), showgrid = FALSE, zeroline = FALSE,
                     showticklabels = FALSE, ticks = ""),
        yaxis = list(title = "", showgrid = FALSE, zeroline = FALSE,
                     showticklabels = FALSE, ticks = ""),
        margin = list(l = 10, r = 28, t = 0, b = 0),
        showlegend = FALSE,
        annotations = ann
      ) |>
      plotly::config(displayModeBar = FALSE, responsive = TRUE)
  }

  # ---------------------------------------------------------------------------
  # Resolver spec SM
  # ---------------------------------------------------------------------------
  .resolver_var_spec_safe <- function(var_madre, ctx, df) {
    f <- get0("resolver_var_spec", mode = "function", ifnotfound = NULL)
    if (is.null(f)) return(list(cols = character(0), map_code_to_label = list()))
    out <- tryCatch(
      f(var_madre = var_madre, ctx = ctx, df = df),
      error = function(e) {
        .log_resumen("resolver_var_spec ERROR var=", var_madre, " -> ", conditionMessage(e))
        list(cols = character(0), map_code_to_label = list())
      }
    )
    if (is.null(out$cols)) out$cols <- character(0)
    if (is.null(out$map_code_to_label)) out$map_code_to_label <- list()
    out
  }

  # ---------------------------------------------------------------------------
  # Título sección
  # ---------------------------------------------------------------------------
  output$section_title_ui <- shiny::renderUI({
    sec <- input$seccion %||% ""
    shiny::HTML(paste0("Resumen de sección: <b>", sec, "</b>"))
  })

  # ---------------------------------------------------------------------------
  # UI: resumen de sección
  # ---------------------------------------------------------------------------
  output$section_summary_ui <- shiny::renderUI({

    shiny::req(input$seccion)
    df <- data_filtrada()
    if (!nrow(df)) {
      return(shiny::div(style = paste0("font-size:12px;color:", MSG_COLOR, ";"), "Sin datos."))
    }

    sec <- input$seccion
    vars_sec <- ctx$secciones_limpias[[sec]] %||% character(0)
    if (!length(vars_sec)) {
      return(shiny::div(style = paste0("font-size:12px;color:", MSG_COLOR, ";"), "Sin variables disponibles."))
    }

    surv <- instrumento$survey %||% NULL

    vars_so <- vars_sec[vapply(vars_sec, function(v)
      tipo_pregunta(v, survey = surv, sm_vars_force = ctx$sm_madres %||% NULL, df = df) == "so",
      logical(1)
    )]

    vars_sm <- vars_sec[vapply(vars_sec, function(v)
      tipo_pregunta(v, survey = surv, sm_vars_force = ctx$sm_madres %||% NULL, df = df) == "sm",
      logical(1)
    )]

    if (length(vars_so) > MAX_SO_ROWS) vars_so <- vars_so[seq_len(MAX_SO_ROWS)]
    vars_show <- c(vars_so, vars_sm)

    if (!length(vars_show)) {
      return(shiny::div(style = paste0("font-size:12px;color:", MSG_COLOR, ";"), "Sin variables resumibles."))
    }

    shiny::div(
      class = "section-summary",
      lapply(seq_along(vars_show), function(i) {

        v <- vars_show[i]
        tp <- tipo_pregunta(v, survey = surv, sm_vars_force = ctx$sm_madres %||% NULL, df = df)

        lab <- .obtener_label_var(v, instrumento, data)
        lab_html <- .wrap_titulo_html(lab, width = 120)

        if (tp == "so") {
          out_id <- paste0("sum_plot_", i)
          return(
            shiny::div(
              class = "summary-row",
              shiny::div(class = "summary-row-title", shiny::HTML(lab_html)),
              shiny::div(
                class = "summary-row-plot",
                plotly::plotlyOutput(out_id, height = paste0(BAR_HEIGHT, "px"))
              )
            )
          )
        }

        spec <- .resolver_var_spec_safe(var_madre = v, ctx = ctx, df = df)
        cols <- spec$cols %||% character(0)

        if (!length(cols)) {
          return(
            shiny::div(
              class = "summary-row",
              shiny::div(class = "summary-row-title", shiny::HTML(lab_html)),
              shiny::div(style = paste0("font-size:12px;color:", MSG_COLOR, ";"), "SM sin dummies disponibles.")
            )
          )
        }

        shiny::div(
          class = "summary-row",
          shiny::div(class = "summary-row-title", shiny::HTML(lab_html)),
          shiny::div(
            class = "summary-row-plot",
            style = "height:auto; overflow:visible;",
            shiny::div(
              class = "sm-card-inner",
              style = "display:flex; flex-direction:column; gap:12px; height:auto; overflow:visible;",
              lapply(seq_along(cols), function(j) {
                colj <- cols[j]

                code <- sub(paste0("^", v, "\\."), "", colj)
                opt_label <- spec$map_code_to_label[[code]] %||% code

                out_id <- paste0("sum_plot_", i, "_", j)

                shiny::div(
                  class = "sm-option-block",
                  style = "height:auto; overflow:visible;",
                  shiny::div(
                    class = "sm-option-title",
                    style = paste0(
                      "color:", SM_OPTION_TEXT, ";",
                      "font-size:12px;",
                      "font-weight:400;",
                      "margin:0 0 6px 0;"
                    ),
                    opt_label
                  ),
                  plotly::plotlyOutput(out_id, height = paste0(BAR_HEIGHT, "px"))
                )
              })
            )
          )
        )
      })
    )
  })

  # ---------------------------------------------------------------------------
  # Render dinámico de plots del resumen
  # ---------------------------------------------------------------------------
  shiny::observe({
    shiny::req(input$seccion)

    df <- data_filtrada()
    sec <- input$seccion
    vars_sec <- ctx$secciones_limpias[[sec]] %||% character(0)
    if (!length(vars_sec)) return()

    surv <- instrumento$survey %||% NULL

    vars_so <- vars_sec[vapply(vars_sec, function(v)
      tipo_pregunta(v, survey = surv, sm_vars_force = ctx$sm_madres %||% NULL, df = df) == "so",
      logical(1)
    )]

    vars_sm <- vars_sec[vapply(vars_sec, function(v)
      tipo_pregunta(v, survey = surv, sm_vars_force = ctx$sm_madres %||% NULL, df = df) == "sm",
      logical(1)
    )]

    if (length(vars_so) > MAX_SO_ROWS) vars_so <- vars_so[seq_len(MAX_SO_ROWS)]
    vars_show <- c(vars_so, vars_sm)

    .log_resumen("Resumen sección=", sec, " | vars_show=", .safe_chr(vars_show))

    for (i in seq_along(vars_show)) {
      local({
        ii <- i
        v  <- vars_show[ii]

        out_so <- paste0("sum_plot_", ii)

        output[[out_so]] <- plotly::renderPlotly({
          df2 <- data_filtrada()
          if (!nrow(df2)) {
            return(
              plotly::plot_ly(height = BAR_HEIGHT) |>
                plotly::layout(annotations = list(list(text = "Sin datos.", showarrow = FALSE))) |>
                plotly::config(displayModeBar = FALSE, responsive = TRUE)
            )
          }

          tp <- tipo_pregunta(v, survey = surv, sm_vars_force = ctx$sm_madres %||% NULL, df = df2)
          if (tp != "so") return(NULL)

          cats <- get_categorias(
            var = v,
            df = df2,
            survey = surv,
            orders_list = (instrumento$orders_list %||% NULL),
            opciones_excluir = NULL
          )

          pal <- .resolver_paleta_var_safe(v, opcion_levels = as.character(cats$labels))
          .plot_so_total(df2, v, paleta_colores = pal)
        })

        tp0 <- tipo_pregunta(v, survey = surv, sm_vars_force = ctx$sm_madres %||% NULL, df = df)
        if (tp0 == "sm") {

          spec0 <- .resolver_var_spec_safe(var_madre = v, ctx = ctx, df = df)
          cols0 <- spec0$cols %||% character(0)
          if (!length(cols0)) return()

          .log_resumen("SM madre=", v, " | dummies=", .safe_chr(cols0))

          for (j in seq_along(cols0)) {
            local({
              jj   <- j
              colj <- cols0[jj]
              out_id <- paste0("sum_plot_", ii, "_", jj)

              output[[out_id]] <- plotly::renderPlotly({
                df2 <- data_filtrada()
                if (!nrow(df2)) {
                  return(
                    plotly::plot_ly(height = BAR_HEIGHT) |>
                      plotly::layout(annotations = list(list(text = "Sin datos.", showarrow = FALSE))) |>
                      plotly::config(displayModeBar = FALSE, responsive = TRUE)
                  )
                }

                .plot_sm_dummy_fill(
                  df = df2,
                  col_dummy = colj,
                  col_yes = SM_COLOR_YES,
                  col_bg  = SM_COLOR_BG,
                  text_out_color = SM_TEXT_OUT,
                  pct_inside_threshold = 0.05
                )
              })
            })
          }
        }
      })
    }
  })

  # ---------------------------------------------------------------------------
  # KPI STATE
  # ---------------------------------------------------------------------------
  kpi_state <- shiny::reactive({
    df <- data_filtrada()
    if (!nrow(df)) return(list(ok = FALSE, msg = "Sin datos."))

    .log_resumen("KPI state -> ctx$kpi_vars raw=", .safe_chr(ctx$kpi_vars))
    .log_resumen("KPI state -> names(df) ejemplo=", .safe_chr(names(df), max_n = 40))

    kpi_vars <- ctx$kpi_vars %||% character(0)
    kpi_vars <- unique(kpi_vars[kpi_vars %in% names(df)])
    if (length(kpi_vars) > 2L) kpi_vars <- kpi_vars[1:2]

    .log_resumen("KPI state -> vars filtradas=", .safe_chr(kpi_vars))

    n_unidades <- if (!is.null(ctx$id_unidad) && ctx$id_unidad %in% names(df)) {
      dplyr::n_distinct(df[[ctx$id_unidad]])
    } else {
      nrow(df)
    }

    n_sufijo <- if (!is.null(ctx$id_unidad) && nzchar(ctx$id_unidad)) ctx$id_unidad else ""
    texto_N  <- paste0(
      "N: ",
      format(n_unidades, big.mark = ",", scientific = FALSE),
      if (nzchar(n_sufijo)) paste0(" ", n_sufijo) else ""
    )

    kpi_obj_1 <- NULL
    kpi_obj_2 <- NULL

    if (length(kpi_vars) >= 1) {
      kpi_obj_1 <- tryCatch(
        .construir_kpi_halfdonut_safe(df = df, var_kpi = kpi_vars[1]),
        error = function(e) {
          .log_resumen("KPI state ERROR obj1 -> ", conditionMessage(e))
          NULL
        }
      )
    }

    if (length(kpi_vars) >= 2) {
      kpi_obj_2 <- tryCatch(
        .construir_kpi_halfdonut_safe(df = df, var_kpi = kpi_vars[2]),
        error = function(e) {
          .log_resumen("KPI state ERROR obj2 -> ", conditionMessage(e))
          NULL
        }
      )
    }

    .log_resumen(
      "KPI state -> obj1 null=", is.null(kpi_obj_1),
      " | obj2 null=", is.null(kpi_obj_2)
    )

    list(
      ok        = TRUE,
      texto_N   = texto_N,
      kpi_vars  = kpi_vars,
      kpi_obj_1 = kpi_obj_1,
      kpi_obj_2 = kpi_obj_2
    )
  })

  # ---------------------------------------------------------------------------
  # RenderPlotly KPIs
  # ---------------------------------------------------------------------------
  output$kpi_plot_1 <- plotly::renderPlotly({
    st <- kpi_state()
    .log_resumen("render kpi_plot_1 -> obj null=", is.null(st$kpi_obj_1))
    if (!isTRUE(st$ok) || is.null(st$kpi_obj_1)) return(NULL)
    st$kpi_obj_1$plot |>
      plotly::config(displayModeBar = FALSE, responsive = TRUE)
  })

  output$kpi_plot_2 <- plotly::renderPlotly({
    st <- kpi_state()
    .log_resumen("render kpi_plot_2 -> obj null=", is.null(st$kpi_obj_2))
    if (!isTRUE(st$ok) || is.null(st$kpi_obj_2)) return(NULL)
    st$kpi_obj_2$plot |>
      plotly::config(displayModeBar = FALSE, responsive = TRUE)
  })

  # ---------------------------------------------------------------------------
  # KPI panel UI
  # ---------------------------------------------------------------------------
  output$kpi_panel <- shiny::renderUI({

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

    st <- kpi_state()
    if (!isTRUE(st$ok)) {
      return(shiny::div(
        style = paste0("font-size:12px;color:", MSG_COLOR, ";padding:10px;text-align:center;"),
        st$msg %||% ""
      ))
    }

    shiny::div(
      class = "kpi-sidebar-stack",

      shiny::div(
        class = "kpi-n-card",
        shiny::div(class = "kpi-n-text", st$texto_N)
      ),

      if (!is.null(st$kpi_obj_1)) shiny::div(
        class = "kpi-cell",
        shiny::div(class = "kpi-donut-title", shiny::HTML(st$kpi_obj_1$title_html)),
        plotly::plotlyOutput("kpi_plot_1", height = "260px"),
        legend_html(st$kpi_obj_1$legend)
      ) else NULL,

      if (!is.null(st$kpi_obj_2)) shiny::div(
        class = "kpi-cell",
        shiny::div(class = "kpi-donut-title", shiny::HTML(st$kpi_obj_2$title_html)),
        plotly::plotlyOutput("kpi_plot_2", height = "260px"),
        legend_html(st$kpi_obj_2$legend)
      ) else NULL,

      if (is.null(st$kpi_obj_1) && is.null(st$kpi_obj_2)) shiny::div(
        style = paste0("font-size:12px;color:", MSG_COLOR, ";padding:10px;text-align:center;"),
        shiny::HTML(
          paste0(
            "No se pudieron construir KPIs.",
            if (!is.null(st$kpi_vars) && length(st$kpi_vars)) {
              paste0("<br><span style='font-size:11px;'>Variables filtradas: ", paste(st$kpi_vars, collapse = ", "), "</span>")
            } else {
              "<br><span style='font-size:11px;'>Variables filtradas: ninguna</span>"
            },
            if (!is.null(ctx$kpi_vars) && length(ctx$kpi_vars)) {
              paste0("<br><span style='font-size:11px;'>ctx$kpi_vars raw: ", paste(ctx$kpi_vars, collapse = ", "), "</span>")
            } else {
              "<br><span style='font-size:11px;'>ctx$kpi_vars raw: vacío</span>"
            }
          )
        )
      ) else NULL
    )
  })

  invisible(NULL)
}
