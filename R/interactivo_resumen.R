# =============================================================================
# Tab 1: Resumen (UI + server) — v3.1 (perfil en sidebar, minimal, barras apiladas)
# -----------------------------------------------------------------------------
# Fixes v3:
# - KPIs (donuts) renderizan desde el primer load (renderPlotly fuera de renderUI)
# - Chip N centrado y con ancho consistente (estructura + CSS ya pegado)
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

      # PERFIL (abajo del sidebar)
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

  # ---------------------------------------------------------------------------
  # Parámetros del resumen
  # ---------------------------------------------------------------------------
  MAX_SO_ROWS <- 16L
  SM_TOP_K    <- 6L
  BAR_HEIGHT  <- 64
  PCT_FSIZE   <- 13

  `%||%` <- get0("%||%", ifnotfound = function(x, y) if (!is.null(x)) x else y)

  .wrap_titulo_html <- get0(
    ".wrap_titulo_html",
    ifnotfound = function(txt, width = 110) {
      if (!requireNamespace("stringr", quietly = TRUE)) return(as.character(txt))
      if (is.null(txt)) return("")
      lineas <- stringr::str_wrap(as.character(txt), width = width)
      paste(lineas, collapse = "<br>")
    }
  )

  .obtener_label_var <- get0(
    ".obtener_label_var",
    ifnotfound = function(var, instrumento, data = NULL) {
      surv <- instrumento$survey
      if (!is.null(surv) && all(c("name","label") %in% names(surv)) && var %in% surv$name) {
        lab <- surv$label[surv$name == var][1]
        if (!is.na(lab) && nzchar(as.character(lab))) return(as.character(lab))
      }
      if (!is.null(data) && var %in% names(data)) {
        vl <- attr(data[[var]], "label", exact = TRUE)
        if (!is.null(vl) && nzchar(as.character(vl))) return(as.character(vl))
      }
      as.character(var)
    }
  )

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
  # Helpers tipo / categorías
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
      if (any(grepl("^select_one(\\s|$)",      tipos))) return("so")
    }
    if (!is.null(df) && .has_var_or_dummies(df, var) && !(var %in% names(df))) return("sm")
    "so"
  }

  col_sm_compact <- function(df, var) {
    v_orig <- paste0(var, "_ORIG")
    if (v_orig %in% names(df)) return(v_orig)
    if (var %in% names(df))    return(var)
    NA_character_
  }

  sm_compact_to_long <- function(x, id) {
    tibble::tibble(id = id, valor = as.character(x)) |>
      tidyr::separate_rows(valor, sep = "\\s*;\\s*", convert = FALSE) |>
      dplyr::mutate(valor = trimws(valor)) |>
      dplyr::filter(!is.na(valor) & nzchar(valor) & valor != "NA")
  }

  get_categorias <- function(var, df, survey = NULL, orders_list = NULL, opciones_excluir = NULL) {

    x <- if (var %in% names(df)) df[[var]] else NULL
    lab_attr <- if (!is.null(x)) attr(x, "labels", exact = TRUE) else NULL

    ln <- NA_character_
    if (!is.null(survey) && all(c("name","list_name") %in% names(survey))) {
      ln <- as.character(survey$list_name[survey$name == var][1])
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
  # Plot SO: barra apilada “Total” (texto blanco centrado)
  # ---------------------------------------------------------------------------
  .plot_so_total <- function(df, var, paleta_colores) {

    df2 <- df
    if (!var %in% names(df2)) {
      return(plotly::plot_ly(height = BAR_HEIGHT) |>
               plotly::layout(annotations = list(list(text="Sin variable.", showarrow=FALSE))))
    }

    x <- as.character(df2[[var]])
    x <- x[!is.na(x) & nzchar(x) & x != "NA"]
    if (!length(x)) {
      return(plotly::plot_ly(height = BAR_HEIGHT) |>
               plotly::layout(annotations = list(list(text="Sin datos.", showarrow=FALSE))))
    }

    tab <- as.data.frame(table(x), stringsAsFactors = FALSE)
    names(tab) <- c("code", "n")
    tab$n <- as.numeric(tab$n)
    tab$pct <- tab$n / sum(tab$n)

    map_code_to_label <- NULL
    labs <- attr(df2[[var]], "labels", exact = TRUE)
    if (!is.null(labs) && length(labs) > 0) {
      map_code_to_label <- stats::setNames(as.character(unname(labs)), as.character(names(labs)))
    }

    tab$label <- if (!is.null(map_code_to_label)) {
      out <- unname(map_code_to_label[tab$code])
      out[is.na(out) | out==""] <- tab$code[is.na(out) | out==""]
      out
    } else {
      tab$code
    }

    if (!is.null(paleta_colores) && !is.null(names(paleta_colores)) && all(tab$label %in% names(paleta_colores))) {
      tab$label <- factor(tab$label, levels = names(paleta_colores))
      tab <- tab[order(tab$label), , drop = FALSE]
    } else {
      tab <- tab[order(tab$pct, decreasing = TRUE), , drop = FALSE]
    }

    tab$txt <- paste0("<b>", round(100 * tab$pct, 0), "%</b>")
    tab$hover <- sprintf("%s: %s%%<br>n: %s",
                         as.character(tab$label),
                         round(100 * tab$pct, 1),
                         format(tab$n, big.mark=","))

    p <- plotly::plot_ly(height = BAR_HEIGHT)

    for (lab in as.character(tab$label)) {
      d <- tab[as.character(tab$label) == lab, , drop = FALSE]
      if (!nrow(d)) next

      col <- if (!is.null(paleta_colores) && !is.null(names(paleta_colores)) && lab %in% names(paleta_colores)) {
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
        xaxis = list(title="", range=c(0,1), showgrid=FALSE, zeroline=FALSE, showticklabels=FALSE, ticks=""),
        yaxis = list(title="", showgrid=FALSE, zeroline=FALSE, showticklabels=FALSE, ticks=""),
        margin = list(l=10, r=10, t=0, b=0),
        showlegend = FALSE
      ) |>
      plotly::config(displayModeBar = FALSE, responsive = TRUE)
  }

  # ---------------------------------------------------------------------------
  # SM: Top-K + Otras + No marcado (completa 100% en personas con alguna mención)
  # ---------------------------------------------------------------------------
  .plot_sm_topk <- function(df, var, paleta_colores, top_k = 6L) {

    if (!.has_var_or_dummies(df, var)) {
      return(plotly::plot_ly(height = BAR_HEIGHT) |>
               plotly::layout(annotations = list(list(text="Sin variable.", showarrow=FALSE))))
    }

    colc <- col_sm_compact(df, var)
    if (is.na(colc) || !colc %in% names(df)) {
      return(plotly::plot_ly(height = BAR_HEIGHT) |>
               plotly::layout(annotations = list(list(text="SM no disponible.", showarrow=FALSE))))
    }

    long <- sm_compact_to_long(df[[colc]], id = seq_len(nrow(df)))
    if (!nrow(long)) {
      return(plotly::plot_ly(height = BAR_HEIGHT) |>
               plotly::layout(annotations = list(list(text="Sin menciones.", showarrow=FALSE))))
    }

    denom_ids <- sort(unique(long$id))
    N <- length(denom_ids)
    if (N <= 0) {
      return(plotly::plot_ly(height = BAR_HEIGHT) |>
               plotly::layout(annotations = list(list(text="Sin menciones.", showarrow=FALSE))))
    }

    counts <- as.data.frame(table(long$valor), stringsAsFactors = FALSE)
    names(counts) <- c("code", "n")
    counts$n <- as.numeric(counts$n)
    counts$pct <- counts$n / N
    counts <- counts[order(counts$pct, decreasing = TRUE), , drop = FALSE]

    map_code_to_label <- NULL
    labs <- if (var %in% names(df)) attr(df[[var]], "labels", exact = TRUE) else NULL
    if (!is.null(labs) && length(labs) > 0) {
      map_code_to_label <- stats::setNames(as.character(unname(labs)), as.character(names(labs)))
    } else {
      surv <- instrumento$survey
      ch   <- instrumento$choices %||% NULL
      if (!is.null(surv) && !is.null(ch) &&
          all(c("name","list_name") %in% names(surv)) &&
          all(c("list_name","name","label") %in% names(ch))) {
        ln <- surv$list_name[surv$name == var][1]
        if (!is.na(ln) && nzchar(ln)) {
          ch_v <- ch[ch$list_name == ln, , drop = FALSE]
          if (nrow(ch_v)) map_code_to_label <- stats::setNames(as.character(ch_v$label), as.character(ch_v$name))
        }
      }
    }

    counts$label <- if (!is.null(map_code_to_label)) {
      out <- unname(map_code_to_label[counts$code])
      out[is.na(out) | out==""] <- counts$code[is.na(out) | out==""]
      out
    } else {
      counts$code
    }

    if (nrow(counts) > top_k) {
      top  <- counts[seq_len(top_k), , drop = FALSE]
      rest <- counts[-seq_len(top_k), , drop = FALSE]
      counts <- rbind(
        top,
        data.frame(code="__otras__", n=sum(rest$n), pct=sum(rest$pct), label="Otras", stringsAsFactors = FALSE)
      )
    }

    resto <- max(0, 1 - sum(counts$pct, na.rm = TRUE))
    if (resto > 1e-8) {
      counts <- rbind(
        counts,
        data.frame(code="__resto__", n=round(resto * N, 0), pct=resto, label="No marcado", stringsAsFactors = FALSE)
      )
    }

    counts$txt <- paste0("<b>", round(100 * counts$pct, 0), "%</b>")
    counts$hover <- sprintf("%s: %s%%<br>n: %s",
                            counts$label,
                            round(100 * counts$pct, 1),
                            format(round(counts$n,0), big.mark=","))

    cols <- paleta_colores
    if (is.null(cols) || !length(cols)) {
      cols <- grDevices::hcl.colors(max(3L, nrow(counts)), "Blues")
      cols <- cols[seq_len(nrow(counts))]
      names(cols) <- counts$label
    } else {
      if (is.null(names(cols))) {
        cols <- rep(cols, length.out = nrow(counts))
        names(cols) <- counts$label
      } else {
        falt <- setdiff(counts$label, names(cols))
        if (length(falt)) {
          extra <- grDevices::hcl.colors(max(3L, length(falt)), "Blues")
          extra <- extra[seq_len(length(falt))]
          cols <- c(cols, stats::setNames(extra, falt))
        }
        cols <- cols[counts$label]
        names(cols) <- counts$label
      }
    }

    p <- plotly::plot_ly(height = BAR_HEIGHT)

    for (lab in counts$label) {
      d <- counts[counts$label == lab, , drop = FALSE]
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
          marker           = list(color = unname(cols[[lab]]), line = list(width = 0))
        )
    }

    p |>
      plotly::layout(
        barmode = "stack",
        xaxis = list(title="", range=c(0,1), showgrid=FALSE, zeroline=FALSE, showticklabels=FALSE, ticks=""),
        yaxis = list(title="", showgrid=FALSE, zeroline=FALSE, showticklabels=FALSE, ticks=""),
        margin = list(l=10, r=10, t=0, b=0),
        showlegend = FALSE
      ) |>
      plotly::config(displayModeBar = FALSE, responsive = TRUE)
  }

  # ---------------------------------------------------------------------------
  # Título sección (minimal)
  # ---------------------------------------------------------------------------
  output$section_title_ui <- shiny::renderUI({
    sec <- input$seccion %||% ""
    shiny::HTML(paste0("Resumen de sección: <b>", sec, "</b>"))
  })

  # ---------------------------------------------------------------------------
  # UI: resumen de sección (solo título + barra)
  # ---------------------------------------------------------------------------
  output$section_summary_ui <- shiny::renderUI({

    shiny::req(input$seccion)
    df <- data_filtrada()
    if (!nrow(df)) {
      return(shiny::div(style="font-size:12px;color:#5f6b7a;", "Sin datos."))
    }

    sec <- input$seccion
    vars_sec <- ctx$secciones_limpias[[sec]] %||% character(0)
    if (!length(vars_sec)) {
      return(shiny::div(style="font-size:12px;color:#5f6b7a;", "Sin variables disponibles."))
    }

    surv <- instrumento$survey %||% NULL

    vars_so <- vars_sec[vapply(vars_sec, function(v) tipo_pregunta(v, survey = surv, sm_vars_force = ctx$sm_madres %||% NULL, df = df) == "so", logical(1))]
    vars_sm <- vars_sec[vapply(vars_sec, function(v) tipo_pregunta(v, survey = surv, sm_vars_force = ctx$sm_madres %||% NULL, df = df) == "sm", logical(1))]

    if (length(vars_so) > MAX_SO_ROWS) vars_so <- vars_so[seq_len(MAX_SO_ROWS)]
    vars_show <- c(vars_so, vars_sm)

    if (!length(vars_show)) {
      return(shiny::div(style="font-size:12px;color:#5f6b7a;", "Sin variables resumibles."))
    }

    shiny::div(
      class = "section-summary",
      lapply(seq_along(vars_show), function(i) {
        v <- vars_show[i]
        out_id <- paste0("sum_plot_", i)

        lab <- .obtener_label_var(v, instrumento, data)
        lab_html <- .wrap_titulo_html(lab, width = 120)

        shiny::div(
          class = "summary-row",
          shiny::div(class="summary-row-title", shiny::HTML(lab_html)),
          shiny::div(class="summary-row-plot", plotly::plotlyOutput(out_id, height = paste0(BAR_HEIGHT, "px")))
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

    vars_so <- vars_sec[vapply(vars_sec, function(v) tipo_pregunta(v, survey = surv, sm_vars_force = ctx$sm_madres %||% NULL, df = df) == "so", logical(1))]
    vars_sm <- vars_sec[vapply(vars_sec, function(v) tipo_pregunta(v, survey = surv, sm_vars_force = ctx$sm_madres %||% NULL, df = df) == "sm", logical(1))]

    if (length(vars_so) > MAX_SO_ROWS) vars_so <- vars_so[seq_len(MAX_SO_ROWS)]
    vars_show <- c(vars_so, vars_sm)

    for (i in seq_along(vars_show)) {
      local({
        ii <- i
        v  <- vars_show[ii]
        out_id <- paste0("sum_plot_", ii)

        output[[out_id]] <- plotly::renderPlotly({
          df2 <- data_filtrada()
          if (!nrow(df2)) {
            return(plotly::plot_ly(height = BAR_HEIGHT) |>
                     plotly::layout(annotations = list(list(text="Sin datos.", showarrow=FALSE))))
          }

          tp <- tipo_pregunta(v, survey = surv, sm_vars_force = ctx$sm_madres %||% NULL, df = df2)

          cats <- get_categorias(
            var = v,
            df = df2,
            survey = surv,
            orders_list = (instrumento$orders_list %||% NULL),
            opciones_excluir = NULL
          )

          pal <- NULL
          if (exists(".resolver_paleta_var", mode = "function")) {
            pal <- tryCatch(
              .resolver_paleta_var(
                var = v,
                instrumento = instrumento,
                colores_apiladas_por_listname = ctx$colores_apiladas_por_listname,
                opcion_levels = as.character(cats$labels)
              ),
              error = function(e) NULL
            )
          }

          if (tp == "so") .plot_so_total(df2, v, paleta_colores = pal)
          else           .plot_sm_topk(df2, v, paleta_colores = pal, top_k = SM_TOP_K)
        })
      })
    }
  })

  # ---------------------------------------------------------------------------
  # KPI STATE (reactive) — clave para que donuts carguen al inicio
  # ---------------------------------------------------------------------------
  kpi_state <- shiny::reactive({
    df <- data_filtrada()
    if (!nrow(df)) return(list(ok = FALSE, msg = "Sin datos."))

    kpi_vars <- ctx$kpi_vars %||% character(0)
    kpi_vars <- unique(kpi_vars[kpi_vars %in% names(df)])
    if (length(kpi_vars) > 2L) kpi_vars <- kpi_vars[1:2]

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
      kpi_obj_1 <- .construir_kpi_halfdonut(
        df = df,
        var_kpi = kpi_vars[1],
        instrumento = instrumento,
        colores_apiladas_por_listname = ctx$colores_apiladas_por_listname,
        codigos_perdidos = ctx$codigos_perdidos
      )
    }

    if (length(kpi_vars) >= 2) {
      kpi_obj_2 <- .construir_kpi_halfdonut(
        df = df,
        var_kpi = kpi_vars[2],
        instrumento = instrumento,
        colores_apiladas_por_listname = ctx$colores_apiladas_por_listname,
        codigos_perdidos = ctx$codigos_perdidos
      )
    }

    list(
      ok        = TRUE,
      texto_N   = texto_N,
      kpi_obj_1 = kpi_obj_1,
      kpi_obj_2 = kpi_obj_2
    )
  })

  # ---------------------------------------------------------------------------
  # RenderPlotly KPIs — fuera del renderUI (fix carga inicial)
  # ---------------------------------------------------------------------------
  output$kpi_plot_1 <- plotly::renderPlotly({
    st <- kpi_state()
    if (!isTRUE(st$ok) || is.null(st$kpi_obj_1)) return(NULL)
    st$kpi_obj_1$plot
  })

  output$kpi_plot_2 <- plotly::renderPlotly({
    st <- kpi_state()
    if (!isTRUE(st$ok) || is.null(st$kpi_obj_2)) return(NULL)
    st$kpi_obj_2$plot
  })

  # ---------------------------------------------------------------------------
  # KPI panel UI (sidebar)
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
      return(shiny::div(style="font-size:12px;color:#5f6b7a;padding:10px;text-align:center;", st$msg %||% ""))
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
        style="font-size:12px;color:#5f6b7a;padding:10px;text-align:center;",
        "No se pudieron construir KPIs."
      ) else NULL
    )
  })

  invisible(NULL)
}
