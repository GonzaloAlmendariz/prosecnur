# =============================================================================
# Tab 1: Resumen (UI + server) — v3.2-hotfix + SM fill-only
# -----------------------------------------------------------------------------
# - Se elimina `config=` en renderPlotly (no soportado en tu plotly).
# - Modebar desactivada usando plotly::config() dentro del plot.
# - Select_multiple en Resumen: una tarjeta por pregunta, con varias barras
#   (una por opción dummy).
# - Cada dummy SM ahora es UNA sola barra 0–100%:
#   * Se pinta SOLO el % de "Sí" en #93C4EB (Pulso light blue)
#   * Resto gris claro (fondo)
#   * Texto % dentro si >=5%, si no -> a la derecha (outside)
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
  BAR_HEIGHT  <- 64
  PCT_FSIZE   <- 13

  # Colores SM fill-only
  SM_COLOR_YES  <- "#1C679D"  # Sí (Pulso light blue)
  SM_COLOR_BG   <- "#EAF2FB"  # Fondo (gris/azul muy claro)
  SM_TEXT_OUT   <- "white"  # Texto cuando va afuera
  SM_SUBTITLE   <- "#1C679D"  # Subtítulo por opción (Pulso blue)

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
      if (any(grepl("^select_one(\\s|$)",      tipos))) return("so")
    }
    if (!is.null(df) && .has_var_or_dummies(df, var) && !(var %in% names(df))) return("sm")
    "so"
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
  # Plot SO: barra apilada “Total” (igual a tu estilo)
  # ---------------------------------------------------------------------------
  .plot_so_total <- function(df, var, paleta_colores) {

    if (!var %in% names(df)) {
      return(
        plotly::plot_ly(height = BAR_HEIGHT) |>
          plotly::layout(annotations = list(list(text="Sin variable.", showarrow=FALSE))) |>
          plotly::config(displayModeBar = FALSE, responsive = TRUE)
      )
    }

    x <- as.character(df[[var]])
    x <- x[!is.na(x) & nzchar(x) & x != "NA"]
    if (!length(x)) {
      return(
        plotly::plot_ly(height = BAR_HEIGHT) |>
          plotly::layout(annotations = list(list(text="Sin datos.", showarrow=FALSE))) |>
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
      out[is.na(out) | out==""] <- tab$code[is.na(out) | out==""]
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
    tab$hover <- sprintf("%s: %s%%<br>n: %s",
                         as.character(tab$label),
                         round(100 * tab$pct, 1),
                         format(tab$n, big.mark=","))

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
        xaxis = list(title="", range=c(0,1), showgrid=FALSE, zeroline=FALSE,
                     showticklabels=FALSE, ticks=""),
        yaxis = list(title="", showgrid=FALSE, zeroline=FALSE,
                     showticklabels=FALSE, ticks=""),
        margin = list(l=10, r=10, t=0, b=0),
        showlegend = FALSE
      ) |>
      plotly::config(displayModeBar = FALSE, responsive = TRUE)
  }

  # ---------------------------------------------------------------------------
  # Plot SM dummy (fill-only):
  # - UNA barra 0–100%
  # - Se pinta SOLO el % de Sí (SM_COLOR_YES)
  # - Resto SM_COLOR_BG
  # - Texto % dentro si >=5%, si no fuera a la derecha
  # - Hover: %sí + n_sí + N válidos
  # ---------------------------------------------------------------------------
  .plot_sm_dummy_fill <- function(df, col_dummy,
                                  col_yes = SM_COLOR_YES,
                                  col_bg  = SM_COLOR_BG,
                                  text_out_color = SM_TEXT_OUT,
                                  pct_inside_threshold = 0.05) {

    if (!col_dummy %in% names(df)) {
      return(
        plotly::plot_ly(height = BAR_HEIGHT) |>
          plotly::layout(annotations = list(list(text="Sin dummy.", showarrow=FALSE)),
                         xaxis = list(visible=FALSE), yaxis = list(visible=FALSE),
                         margin = list(l=10,r=10,t=0,b=0)) |>
          plotly::config(displayModeBar = FALSE, responsive = TRUE)
      )
    }

    x <- df[[col_dummy]]

    # Normalizar a 0/1 cuando sea posible
    x2 <- suppressWarnings(as.numeric(as.character(x)))
    if (all(is.na(x2)) && is.logical(x)) x2 <- as.numeric(x)

    ok <- !is.na(x2) & x2 %in% c(0, 1)
    x2 <- x2[ok]

    if (!length(x2)) {
      return(
        plotly::plot_ly(height = BAR_HEIGHT) |>
          plotly::layout(annotations = list(list(text="Sin datos.", showarrow=FALSE)),
                         xaxis = list(visible=FALSE), yaxis = list(visible=FALSE),
                         margin = list(l=10,r=10,t=0,b=0)) |>
          plotly::config(displayModeBar = FALSE, responsive = TRUE)
      )
    }

    N     <- length(x2)
    n_yes <- sum(x2 == 1)
    pct_y <- n_yes / N
    pct_r <- 1 - pct_y

    # Segmentos (yes + background) para mantener estética "apilada"
    seg <- data.frame(
      seg   = c("yes", "bg"),
      pct   = c(pct_y, pct_r),
      n_yes = n_yes,
      N     = N,
      stringsAsFactors = FALSE
    )

    # Texto: solo en el segmento "yes"
    pct_txt <- paste0("<b>", round(100 * pct_y, 0), "%</b>")
    seg$text <- c(pct_txt, "")

    # Posición texto según umbral
    textpos_yes <- if (pct_y < pct_inside_threshold) "outside" else "inside"
    textfont_yes <- if (pct_y < pct_inside_threshold) {
      list(color = text_out_color, size = PCT_FSIZE)
    } else {
      list(color = "white", size = PCT_FSIZE)
    }

    # Hover solo en yes (en bg vacío)
    seg$hover <- c(
      sprintf("Sí: %s%%<br>n: %s<br>N: %s",
              round(100 * pct_y, 1),
              format(n_yes, big.mark=","),
              format(N, big.mark=",")),
      ""
    )

    p <- plotly::plot_ly(height = BAR_HEIGHT)

    # YES segment
    p <- p |>
      plotly::add_bars(
        data             = seg[seg$seg == "yes", , drop=FALSE],
        x                = ~pct,
        y                = I("Total"),
        orientation      = "h",
        marker           = list(color = col_yes, line = list(width = 0)),
        text             = ~text,
        textposition     = textpos_yes,
        insidetextanchor = "middle",
        textfont         = textfont_yes,
        customdata       = ~hover,
        hovertemplate    = "%{customdata}<extra></extra>",
        cliponaxis       = FALSE
      )

    # BG segment (sin texto ni hover)
    p <- p |>
      plotly::add_bars(
        data        = seg[seg$seg == "bg", , drop=FALSE],
        x           = ~pct,
        y           = I("Total"),
        orientation = "h",
        marker      = list(color = col_bg, line = list(width = 0)),
        hoverinfo   = "skip",
        showlegend  = FALSE
      )

    p |>
      plotly::layout(
        barmode = "stack",
        xaxis = list(title="", range=c(0,1), showgrid=FALSE, zeroline=FALSE,
                     showticklabels=FALSE, ticks=""),
        yaxis = list(title="", showgrid=FALSE, zeroline=FALSE,
                     showticklabels=FALSE, ticks=""),
        margin = list(l=10, r=16, t=0, b=0), # un poco más de r para el texto outside
        showlegend = FALSE
      ) |>
      plotly::config(displayModeBar = FALSE, responsive = TRUE)
  }

  # ---------------------------------------------------------------------------
  # Resolver spec SM (se espera en helpers)
  # ---------------------------------------------------------------------------
  .resolver_var_spec_safe <- function(var_madre, ctx, df) {
    f <- get0("resolver_var_spec", mode = "function", ifnotfound = NULL)
    if (is.null(f)) return(list(cols = character(0), map_code_to_label = list()))
    out <- tryCatch(f(var_madre = var_madre, ctx = ctx, df = df),
                    error = function(e) list(cols = character(0), map_code_to_label = list()))
    if (is.null(out$cols)) out$cols <- character(0)
    if (is.null(out$map_code_to_label)) out$map_code_to_label <- list()
    out
  }

  # ---------------------------------------------------------------------------
  # Título sección (minimal)
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
      return(shiny::div(style="font-size:12px;color:#5f6b7a;", "Sin datos."))
    }

    sec <- input$seccion
    vars_sec <- ctx$secciones_limpias[[sec]] %||% character(0)
    if (!length(vars_sec)) {
      return(shiny::div(style="font-size:12px;color:#5f6b7a;", "Sin variables disponibles."))
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
      return(shiny::div(style="font-size:12px;color:#5f6b7a;", "Sin variables resumibles."))
    }

    shiny::div(
      class = "section-summary",
      lapply(seq_along(vars_show), function(i) {

        v <- vars_show[i]
        tp <- tipo_pregunta(v, survey = surv, sm_vars_force = ctx$sm_madres %||% NULL, df = df)

        lab <- .obtener_label_var(v, instrumento, data)
        lab_html <- .wrap_titulo_html(lab, width = 120)

        # --- SO: 1 barra (como antes)
        if (tp == "so") {
          out_id <- paste0("sum_plot_", i)
          return(
            shiny::div(
              class = "summary-row",
              shiny::div(class="summary-row-title", shiny::HTML(lab_html)),
              shiny::div(
                class="summary-row-plot",
                plotly::plotlyOutput(out_id, height = paste0(BAR_HEIGHT, "px"))
              )
            )
          )
        }

        # --- SM: una tarjeta por pregunta, múltiples barras internas
        spec <- .resolver_var_spec_safe(var_madre = v, ctx = ctx, df = df)
        cols <- spec$cols %||% character(0)

        if (!length(cols)) {
          return(
            shiny::div(
              class = "summary-row",
              shiny::div(class="summary-row-title", shiny::HTML(lab_html)),
              shiny::div(style="font-size:12px;color:#5f6b7a;", "SM sin dummies disponibles.")
            )
          )
        }

        shiny::div(
          class = "summary-row",

          # Columna izquierda (título pregunta)
          shiny::div(class="summary-row-title", shiny::HTML(lab_html)),

          # Columna derecha (IMPORTANTE: dejar crecer el alto)
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
                      "color:", SM_SUBTITLE, ";",
                      "font-size:12px;",
                      "font-weight:400;",
                      "margin:0 0 6px 0;"
                    ),
                    opt_label
                  ),

                  # cada barra mantiene el mismo alto que SO
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

    for (i in seq_along(vars_show)) {
      local({
        ii <- i
        v  <- vars_show[ii]

        # ---- SO output (siempre definido, pero si no es SO retorna NULL)
        out_so <- paste0("sum_plot_", ii)

        output[[out_so]] <- plotly::renderPlotly({
          df2 <- data_filtrada()
          if (!nrow(df2)) {
            return(
              plotly::plot_ly(height = BAR_HEIGHT) |>
                plotly::layout(annotations = list(list(text="Sin datos.", showarrow=FALSE))) |>
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

          .plot_so_total(df2, v, paleta_colores = pal)
        })

        # ---- SM outputs por dummy (fill-only)
        tp0 <- tipo_pregunta(v, survey = surv, sm_vars_force = ctx$sm_madres %||% NULL, df = df)
        if (tp0 == "sm") {

          spec0 <- .resolver_var_spec_safe(var_madre = v, ctx = ctx, df = df)
          cols0 <- spec0$cols %||% character(0)
          if (!length(cols0)) return()

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
                      plotly::layout(annotations = list(list(text="Sin datos.", showarrow=FALSE))) |>
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
  # KPI STATE (reactive) — se mantiene
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
  # RenderPlotly KPIs (sin config= en renderPlotly)
  # ---------------------------------------------------------------------------
  output$kpi_plot_1 <- plotly::renderPlotly({
    st <- kpi_state()
    if (!isTRUE(st$ok) || is.null(st$kpi_obj_1)) return(NULL)
    st$kpi_obj_1$plot |>
      plotly::config(displayModeBar = FALSE, responsive = TRUE)
  })

  output$kpi_plot_2 <- plotly::renderPlotly({
    st <- kpi_state()
    if (!isTRUE(st$ok) || is.null(st$kpi_obj_2)) return(NULL)
    st$kpi_obj_2$plot |>
      plotly::config(displayModeBar = FALSE, responsive = TRUE)
  })

  # ---------------------------------------------------------------------------
  # KPI panel UI (sidebar) — se mantiene
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
      return(shiny::div(style="font-size:12px;color:#5f6b7a;padding:10px;text-align:center;",
                        st$msg %||% ""))
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
