#' Explorador interactivo de frecuencias (Plotly + Shiny)
#'
#' @param data Data.frame con la base ya adaptada (e.g. `rp_data`).
#' @param instrumento Objeto `prosecnur_instrumento` de `reporte_instrumento()`.
#' @param secciones Lista nombrada: nombre de sección -> vector de nombres de variables.
#' @param fuente Texto breve de fuente a mostrar debajo del gráfico.
#' @param titulo Título general del explorador.
#' @param colores_apiladas_por_listname Lista nombrada por `list_name` con paletas
#'        (como las que ya usas en PPT/Word). Se usan solo los colores.
#' @param facet_vars Vector opcional con nombres de variables que se permiten
#'        usar como “Dividir gráfico por”.
#' @param codigos_perdidos Vector opcional de códigos numéricos que deben
#'        tratarse como valores perdidos (ej. 99, 98, 97).
#'
#' @return Lanza una app Shiny.
#' @export
reporte_interactivo <- function(
    data,
    instrumento,
    secciones,
    fuente  = NULL,
    titulo  = "Explorador interactivo",
    colores_apiladas_por_listname = NULL,
    facet_vars = NULL,
    codigos_perdidos = NULL
) {
  # ────────────────────────────────────────────────────────────
  # Dependencias mínimas
  # ────────────────────────────────────────────────────────────
  if (!requireNamespace("shiny",  quietly = TRUE)) stop("Falta 'shiny'.",  call. = FALSE)
  if (!requireNamespace("plotly", quietly = TRUE)) stop("Falta 'plotly'.", call. = FALSE)
  if (!requireNamespace("dplyr",  quietly = TRUE)) stop("Falta 'dplyr'.",  call. = FALSE)
  if (!requireNamespace("tibble", quietly = TRUE)) stop("Falta 'tibble'.", call. = FALSE)

  `%||%` <- function(x, y) if (is.null(x) || length(x) == 0) y else x

  # ────────────────────────────────────────────────────────────
  # Limpieza previa de códigos perdidos numéricos (ej. 99)
  # ────────────────────────────────────────────────────────────
  if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0) {
    codigos_perdidos <- unique(stats::na.omit(suppressWarnings(as.numeric(codigos_perdidos))))
    if (length(codigos_perdidos) > 0) {
      data <- dplyr::mutate(
        data,
        dplyr::across(
          where(is.numeric),
          ~ ifelse(.x %in% codigos_perdidos, NA_real_, .x)
        )
      )
    }
  }

  # ────────────────────────────────────────────────────────────
  # Helpers de metadatos
  # ────────────────────────────────────────────────────────────
  .get_var_meta <- function(var) {
    survey <- instrumento$survey
    fila   <- survey[survey$name == var, , drop = FALSE]

    if (nrow(fila) == 0L) {
      return(list(
        name      = var,
        label     = instrumento$var_labels[[var]] %||% var,
        list_name = NA_character_,
        type      = NA_character_,
        measure   = NA_character_
      ))
    }

    list(
      name      = var,
      label     = if (!is.null(instrumento$var_labels[[var]]))
        instrumento$var_labels[[var]]
      else as.character(fila$label[1]),
      list_name = fila$list_name[1],
      type      = fila$type[1],
      measure   = fila$measure_sugerida[1]
    )
  }

  .get_orders_for_var <- function(var) {
    instrumento$orders_list[[var]] %||% NULL
  }

  .choices_vars <- function(vars) {
    vars <- vars[vars %in% names(instrumento$var_labels)]
    labs <- instrumento$var_labels[vars]
    labs[is.na(labs)] <- vars[is.na(labs)]
    stats::setNames(vars, labs)
  }

  .choices_categorias <- function(var) {
    ord <- .get_orders_for_var(var)
    if (!is.null(ord)) {
      codes  <- ord$names
      labels <- ord$labels
      stats::setNames(codes, labels)
    } else {
      vals <- sort(unique(as.character(data[[var]])))
      stats::setNames(vals, vals)
    }
  }

  .get_palette_for_var <- function(var) {
    meta <- .get_var_meta(var)
    ln   <- meta$list_name
    pal  <- NULL

    if (!is.null(ln) &&
        !is.na(ln) &&
        !is.null(colores_apiladas_por_listname) &&
        ln %in% names(colores_apiladas_por_listname)) {

      pal_raw <- colores_apiladas_por_listname[[ln]]

      if (is.list(pal_raw) && !is.null(pal_raw$colores)) {
        pal <- pal_raw$colores
      } else {
        pal <- pal_raw
      }
    }

    if (is.null(pal)) {
      pal <- c(
        "#39588B", "#88CC88", "#F6D55C", "#F4A261",
        "#E67E73", "#D95F02", "#7570B3", "#B0B0B0"
      )
    }

    pal
  }

  # ────────────────────────────────────────────────────────────
  # Tabla de frecuencias (incluye categorías 0%)
  # ────────────────────────────────────────────────────────────
  .tab_freq_var <- function(var, data_filtrada) {
    meta <- .get_var_meta(var)
    ord  <- .get_orders_for_var(var)

    vec <- as.character(data_filtrada[[var]])

    tab_raw <- dplyr::count(tibble::tibble(val = vec), val, name = "n")

    if (!is.null(ord)) {
      codes  <- ord$names
      labels <- ord$labels
      base   <- dplyr::tibble(code = codes, label = labels)
      tab    <- dplyr::left_join(
        base,
        dplyr::rename(tab_raw, code = val),
        by = "code"
      )
      tab$n[is.na(tab$n)] <- 0L
    } else {
      tab <- dplyr::tibble(
        code  = tab_raw$val,
        label = tab_raw$val,
        n     = tab_raw$n
      )
    }

    total <- sum(tab$n, na.rm = TRUE)
    if (total == 0) total <- 1

    tab <- dplyr::mutate(
      tab,
      pct = n / total
    )

    tab <- dplyr::mutate(
      tab,
      pct_plot  = dplyr::if_else(pct == 0, 0.001, pct),
      pct_label = paste0(round(pct * 100), "%"),
      hover     = sprintf(
        "<b>%s</b><br>n = %s<br>Porcentaje: %s",
        as.character(label), n, pct_label
      )
    )

    tab$label <- as.character(tab$label)
    tab$label_factor <- factor(tab$label, levels = rev(tab$label))

    list(
      meta   = meta,
      tabla  = tab,
      total  = total
    )
  }

  # ────────────────────────────────────────────────────────────
  # Construcción de un gráfico Plotly (una variable)
  # ────────────────────────────────────────────────────────────
  .build_plotly_single <- function(freq_obj,
                                   show_axis_x   = FALSE,
                                   show_labels   = TRUE,
                                   titulo_extra  = NULL) {

    meta <- freq_obj$meta
    tab  <- freq_obj$tabla

    pal <- .get_palette_for_var(meta$name)

    if (!is.null(names(pal))) {
      col_vec <- pal[tab$label]
      col_vec[is.na(col_vec)] <- "#39588B"
    } else {
      col_vec <- rep(pal, length.out = nrow(tab))
    }

    # Título
    titulo_main <- meta$label %||% meta$name
    if (!is.null(titulo_extra)) {
      titulo_main <- paste0(titulo_main, " – ", titulo_extra)
    }

    # Umbral para decidir si la etiqueta va dentro o fuera
    threshold <- 0.15

    # Posiciones para inside / outside / cero
    x_inside  <- ifelse(tab$pct >= threshold, tab$pct / 2, NA_real_)
    x_outside <- ifelse(tab$pct < threshold & tab$pct > 0,
                        pmin(tab$pct + 0.02, 0.98),
                        NA_real_)
    x_zero    <- ifelse(tab$pct == 0, 0.01, NA_real_)

    # Barra principal
    fig <- plotly::plot_ly(
      data = tab,
      type = "bar",
      orientation = "h",
      x = ~pct_plot,
      y = ~label_factor,
      marker = list(color = col_vec),
      hovertext = ~hover,
      hoverinfo = "text",
      cliponaxis = FALSE
    )

    # Etiquetas internas (pct >= threshold)
    if (show_labels && any(!is.na(x_inside))) {
      fig <- fig %>%
        plotly::add_trace(
          type  = "scatter",
          mode  = "text",
          x     = x_inside,
          y     = tab$label_factor,
          text  = ifelse(is.na(x_inside), "", tab$pct_label),
          textposition = "middle center",
          textfont = list(
            color  = "#FFFFFF",
            size   = 13,
            family = "Arial",
            weight = "bold"
          ),
          showlegend = FALSE,
          hoverinfo  = "none",
          inherit    = FALSE
        )
    }

    # Etiquetas pequeñas (0 < pct < threshold) a la derecha
    if (show_labels && any(!is.na(x_outside))) {
      fig <- fig %>%
        plotly::add_trace(
          type  = "scatter",
          mode  = "text",
          x     = x_outside,
          y     = tab$label_factor,
          text  = ifelse(is.na(x_outside), "", tab$pct_label),
          textposition = "middle right",
          textfont = list(
            color  = "#39588B",
            size   = 11,
            family = "Arial"
          ),
          showlegend = FALSE,
          hoverinfo  = "none",
          inherit    = FALSE
        )
    }

    # Etiquetas 0% explícitas
    if (show_labels && any(!is.na(x_zero))) {
      fig <- fig %>%
        plotly::add_trace(
          type  = "scatter",
          mode  = "text",
          x     = x_zero,
          y     = tab$label_factor,
          text  = ifelse(is.na(x_zero), "", "0%"),
          textposition = "middle right",
          textfont = list(
            color  = "#39588B",
            size   = 11,
            family = "Arial"
          ),
          showlegend = FALSE,
          hoverinfo  = "none",
          inherit    = FALSE
        )
    }

    fig <- fig %>%
      plotly::layout(
        title = list(
          text  = titulo_main,
          font  = list(
            family = "Arial",
            size   = 18,
            color  = "#39588B"
          ),
          x     = 0,
          xanchor = "left"
        ),
        xaxis = list(
          title          = "",
          range          = c(0, 1),
          tickformat     = ".0%",
          showgrid       = TRUE,
          zeroline       = FALSE,
          showticklabels = TRUE,
          gridcolor      = "rgba(0,0,0,0.06)",
          ticklen        = 6,
          tickpad        = 4,
          tickfont       = list(
            family = "Arial",
            size   = 11,
            color  = "#444444"
          )
        ),
        yaxis = list(
          title     = "",
          autorange = "reversed",
          ticklen   = 6,
          tickpad   = 6,
          tickfont  = list(
            family = "Arial",
            size   = 11,
            color  = "#444444"
          )
        ),
        margin = list(l = 130, r = 40, t = 70, b = 60),
        bargap = 0.25,
        plot_bgcolor  = "rgba(0,0,0,0)",
        paper_bgcolor = "rgba(0,0,0,0)",
        showlegend = FALSE,
        uniformtext = list(
          minsize = 9,
          mode    = "hide"
        ),
        transition = list(
          duration = 300,
          easing   = "cubic-in-out"
        )
      ) %>%
      plotly::config(
        displaylogo = FALSE,
        modeBarButtonsToRemove = c(
          "lasso2d", "select2d", "autoScale2d",
          "zoomIn2d", "zoomOut2d", "resetScale2d"
        )
      )

    fig
  }

  .build_plotly_for_var <- function(var,
                                    data_filtrada,
                                    usar_facet,
                                    facet_var,
                                    show_axis_x,
                                    show_labels) {

    # Sin desagregación
    if (!usar_facet || is.null(facet_var) || identical(facet_var, "")) {
      freq_obj <- .tab_freq_var(var, data_filtrada)
      return(
        .build_plotly_single(
          freq_obj,
          show_axis_x = show_axis_x,
          show_labels = show_labels
        )
      )
    }

    # Con desagregación (facet)
    meta_facet <- .get_var_meta(facet_var)
    ord_facet  <- .get_orders_for_var(facet_var)

    if (!is.null(ord_facet)) {
      codes_f  <- ord_facet$names
      labels_f <- ord_facet$labels
    } else {
      vals     <- sort(unique(as.character(data_filtrada[[facet_var]])))
      codes_f  <- vals
      labels_f <- vals
    }

    if (length(codes_f) > 6) {
      codes_f  <- codes_f[1:6]
      labels_f <- labels_f[1:6]
    }

    plots <- list()
    for (i in seq_along(codes_f)) {
      code_i  <- codes_f[i]
      label_i <- labels_f[i]

      sub <- data_filtrada[as.character(data_filtrada[[facet_var]]) == code_i,
                           , drop = FALSE]
      if (nrow(sub) == 0) next

      freq_i <- .tab_freq_var(var, sub)
      p_i    <- .build_plotly_single(
        freq_i,
        show_axis_x  = TRUE,
        show_labels  = show_labels,
        titulo_extra = label_i
      )
      plots[[length(plots) + 1]] <- p_i
    }

    if (length(plots) == 0) {
      freq_obj <- .tab_freq_var(var, data_filtrada)
      fig <- .build_plotly_single(
        freq_obj,
        show_axis_x = show_axis_x,
        show_labels = show_labels
      )
      return(fig)
    }

    n_panels <- length(plots)
    n_cols   <- if (n_panels <= 2) 1 else 2
    n_rows   <- ceiling(n_panels / n_cols)

    fig <- plotly::subplot(
      plots,
      nrows  = n_rows,
      shareX = TRUE,
      titleX = TRUE,
      margin = 0.06
    ) %>%
      plotly::layout(
        title = list(
          text  = (.get_var_meta(var)$label %||% var),
          font  = list(
            family = "Arial",
            size   = 18,
            color  = "#39588B"
          ),
          x     = 0,
          xanchor = "left"
        ),
        plot_bgcolor  = "rgba(0,0,0,0)",
        paper_bgcolor = "rgba(0,0,0,0)",
        transition = list(
          duration = 300,
          easing   = "cubic-in-out"
        ),
        uniformtext = list(
          minsize = 9,
          mode    = "hide"
        )
      ) %>%
      plotly::config(
        displaylogo = FALSE,
        modeBarButtonsToRemove = c(
          "lasso2d", "select2d", "autoScale2d",
          "zoomIn2d", "zoomOut2d", "resetScale2d"
        )
      )

    fig
  }

  # ────────────────────────────────────────────────────────────
  # UI (barra de configuración unificada arriba)
  # ────────────────────────────────────────────────────────────

  todas_vars <- unique(unlist(secciones, use.names = FALSE))
  todas_vars <- intersect(todas_vars, names(data))

  sec_choices <- stats::setNames(names(secciones), names(secciones))

  ui <- shiny::fluidPage(
    shiny::tags$head(
      shiny::tags$style(
        shiny::HTML("
          body {
            background-color: #F2F3F8;
          }
          .config-wrapper {
            background-color: #FFFFFF;
            border-radius: 10px;
            border: 1px solid #E0E0E0;
            padding: 10px 14px 6px 14px;
            margin-bottom: 12px;
          }
          .config-wrapper h4 {
            margin-top: 2px;
            margin-bottom: 10px;
            font-weight: 600;
            color: #39588B;
          }
          .config-block {
            margin-bottom: 6px;
            padding: 6px 10px 4px 10px;
            border-right: 1px solid #EEEEEE;
          }
          .config-block-last {
            border-right: none;
          }
          .config-block h5 {
            margin-top: 0;
            margin-bottom: 6px;
            font-weight: 600;
            color: #39588B;
            font-size: 14px;
          }
          .fuente-text {
            font-size: 11px;
            color: #666666;
            margin-top: 6px;
          }
        ")
      )
    ),
    shiny::titlePanel(titulo),
    shiny::fluidRow(
      shiny::column(
        width = 12,
        shiny::div(
          class = "config-wrapper",
          shiny::h4("Configuración"),
          shiny::fluidRow(
            shiny::column(
              width = 4,
              shiny::div(
                class = "config-block",
                shiny::h5("Variable a graficar"),
                shiny::selectInput(
                  inputId = "sec_select",
                  label   = "Sección:",
                  choices = sec_choices,
                  selected = names(secciones)[1]
                ),
                shiny::uiOutput("var_select_ui")
              )
            ),
            shiny::column(
              width = 4,
              shiny::div(
                class = "config-block",
                shiny::h5("Filtro"),
                shiny::checkboxInput("usar_filtro", "Usar filtro", value = FALSE),
                shiny::conditionalPanel(
                  condition = "input.usar_filtro == true",
                  shiny::selectInput(
                    inputId = "filtro_var",
                    label   = "Variable de filtro:",
                    choices = .choices_vars(todas_vars),
                    selected = NULL
                  ),
                  shiny::uiOutput("filtro_categorias_ui")
                )
              )
            ),
            shiny::column(
              width = 4,
              shiny::div(
                class = "config-block config-block-last",
                shiny::h5("Desagregación del gráfico"),
                shiny::checkboxInput("usar_facet", "Dividir gráfico por otra variable", value = FALSE),
                shiny::conditionalPanel(
                  condition = "input.usar_facet == true",
                  shiny::selectInput(
                    inputId = "facet_var",
                    label   = "Variable para dividir el gráfico:",
                    choices = character(0),
                    selected = NULL
                  )
                )
              )
            )
          )
        )
      )
    ),
    shiny::fluidRow(
      shiny::column(
        width = 12,
        shiny::div(
          plotly::plotlyOutput("plot_1", height = "720px"),
          shiny::div(
            class = "fuente-text",
            if (!is.null(fuente))
              shiny::HTML(paste0("Fuente: ", fuente))
            else
              shiny::HTML("")
          )
        )
      )
    )
  )

  # ────────────────────────────────────────────────────────────
  # SERVER
  # ────────────────────────────────────────────────────────────
  server <- function(input, output, session) {

    # Variables por sección (selector de variable)
    output$var_select_ui <- shiny::renderUI({
      sec <- input$sec_select
      vars_sec <- secciones[[sec]]
      vars_sec <- intersect(vars_sec, names(data))
      shiny::selectInput(
        inputId = "var_select",
        label   = "Variable:",
        choices = .choices_vars(vars_sec),
        selected = vars_sec[1]
      )
    })

    # Poblar facet_var solo con las facet_vars permitidas
    shiny::observe({
      if (is.null(facet_vars) || length(facet_vars) == 0) {
        shiny::updateSelectInput(
          session, "facet_var",
          choices = character(0),
          selected = NULL
        )
      } else {
        fv <- intersect(facet_vars, todas_vars)
        shiny::updateSelectInput(
          session, "facet_var",
          choices = .choices_vars(fv),
          selected = NULL
        )
      }
    })

    # UI de filtro (categórico vs numérico)
    output$filtro_categorias_ui <- shiny::renderUI({
      shiny::req(input$usar_filtro)
      var_f <- input$filtro_var
      if (is.null(var_f) || var_f == "" || !var_f %in% names(data)) return(NULL)

      meta_f  <- .get_var_meta(var_f)
      measure <- meta_f$measure

      if (!is.null(measure) && !is.na(measure) && measure == "scale") {
        v <- suppressWarnings(as.numeric(data[[var_f]]))
        v <- v[is.finite(v)]
        if (length(v) == 0) return(NULL)
        rng <- range(v, na.rm = TRUE)
        shiny::sliderInput(
          inputId = "filtro_rango",
          label   = "Rango:",
          min     = floor(rng[1]),
          max     = ceiling(rng[2]),
          value   = rng,
          step    = max(1, (ceiling(rng[2]) - floor(rng[1])) / 20)
        )
      } else {
        choices <- .choices_categorias(var_f)
        shiny::checkboxGroupInput(
          inputId = "filtro_categorias",
          label   = "Categorías",
          choices = choices,
          selected = names(choices)
        )
      }
    })

    # Data filtrada
    data_filtrada <- shiny::reactive({
      df <- data
      if (isTRUE(input$usar_filtro)) {
        var_f <- input$filtro_var
        if (!is.null(var_f) && var_f != "" && var_f %in% names(df)) {
          meta_f  <- .get_var_meta(var_f)
          measure <- meta_f$measure

          if (!is.null(measure) && !is.na(measure) && measure == "scale") {
            rng <- input$filtro_rango
            if (!is.null(rng) && length(rng) == 2) {
              v <- suppressWarnings(as.numeric(df[[var_f]]))
              keep <- !is.na(v) & v >= rng[1] & v <= rng[2]
              df <- df[keep, , drop = FALSE]
            }
          } else {
            cats_sel <- input$filtro_categorias
            if (!is.null(cats_sel) && length(cats_sel) > 0) {
              v <- as.character(df[[var_f]])
              df <- df[v %in% cats_sel, , drop = FALSE]
            }
          }
        }
      }
      df
    })

    # Plot principal
    output$plot_1 <- plotly::renderPlotly({
      shiny::req(input$var_select)
      var <- input$var_select
      df  <- data_filtrada()

      if (!var %in% names(df)) {
        return(NULL)
      }

      usar_facet  <- isTRUE(input$usar_facet)
      facet_var   <- input$facet_var

      show_axis_x <- FALSE
      show_labels <- TRUE

      .build_plotly_for_var(
        var           = var,
        data_filtrada = df,
        usar_facet    = usar_facet,
        facet_var     = facet_var,
        show_axis_x   = show_axis_x,
        show_labels   = show_labels
      )
    })
  }

  shiny::shinyApp(ui = ui, server = server)
}
