# =============================================================================
# Explorador interactivo simple: reporte_interactivo()
# Fase 1 – Una variable principal, filtros y cruce opcional
# =============================================================================

# Helper genérico
`%||%` <- function(x, y) if (!is.null(x)) x else y

# -----------------------------------------------------------------------------
# Helpers internos
# -----------------------------------------------------------------------------

#' Obtener etiqueta legible de una variable
#'
#' Busca primero en `instrumento$survey$label`, luego en el atributo
#' `label` de la columna en `data`, y finalmente retorna el nombre
#' de la variable si no encuentra nada.
#'
#' @keywords internal
#' @noRd
.obtener_label_var <- function(var, instrumento, data = NULL) {
  survey <- instrumento$survey

  # 1) survey$label
  if (!is.null(survey) &&
      all(c("name", "label") %in% names(survey)) &&
      var %in% survey$name) {

    lab <- survey$label[survey$name == var][1]
    if (!is.na(lab) && nzchar(as.character(lab))) {
      return(as.character(lab))
    }
  }

  # 2) attr(label) en data
  if (!is.null(data) && var %in% names(data)) {
    vl <- attr(data[[var]], "label", exact = TRUE)
    if (!is.null(vl) && nzchar(as.character(vl))) {
      return(as.character(vl))
    }
  }

  # 3) fallback: nombre de variable
  as.character(var)
}

#' Obtener etiqueta legible de una categoría (choices)
#'
#' Usa el componente `choices` del instrumento si existe.
#'
#' @keywords internal
#' @noRd
.obtener_label_choice <- function(list_name, value, instrumento) {
  choices <- instrumento$choices
  if (is.null(choices) ||
      !all(c("list_name", "name", "label") %in% names(choices)) ||
      is.na(list_name) || !nzchar(list_name)) {

    return(as.character(value))
  }

  ch <- choices[choices$list_name == list_name &
                  choices$name == value, , drop = FALSE]

  if (nrow(ch) == 0L) {
    return(as.character(value))
  }

  lab <- ch$label[1]
  if (is.na(lab) || !nzchar(as.character(lab))) {
    return(as.character(value))
  }

  as.character(lab)
}

#' Wrap del título en HTML (usa <br>)
#'
#' @keywords internal
#' @noRd
.wrap_titulo_html <- function(txt, width = 120) {
  if (!requireNamespace("stringr", quietly = TRUE)) {
    return(txt)
  }
  txt <- as.character(txt)
  if (!nzchar(txt)) return(txt)

  lineas <- stringr::str_wrap(txt, width = width)
  paste(lineas, collapse = "<br>")
}

#' Construir tabla de proporciones para variable principal (simple o cruzada)
#'
#' Devuelve un data.frame largo con:
#'  - estrato_label : etiqueta de la fila (Total o categoría del cruce)
#'  - opcion_label  : etiqueta de la respuesta (No, Sí, etc.)
#'  - pct           : proporción 0–1
#'  - n             : conteo de casos en cada combinación
#'
#' @keywords internal
#' @noRd
.preparar_tabla_proporciones <- function(data,
                                         instrumento,
                                         var,
                                         var_cruce = NULL,
                                         codigos_perdidos = NULL) {

  if (!requireNamespace("dplyr", quietly = TRUE)) {
    stop("Se requiere el paquete 'dplyr' para `reporte_interactivo()`.",
         call. = FALSE)
  }

  survey  <- instrumento$survey
  choices <- instrumento$choices %||% NULL

  if (is.null(survey) || !"name" %in% names(survey)) {
    stop("El `instrumento` debe contener un componente `survey` válido.",
         call. = FALSE)
  }

  # Info de la variable principal
  fila_var <- survey[survey$name == var, , drop = FALSE]
  if (nrow(fila_var) == 0L) {
    stop("La variable '", var, "' no está en `instrumento$survey`.",
         call. = FALSE)
  }
  list_main <- fila_var$list_name[1]

  # Mapeo código → etiqueta para la variable principal
  if (!is.null(choices) &&
      all(c("list_name", "name", "label") %in% names(choices)) &&
      !is.na(list_main) && nzchar(list_main)) {

    ch_main <- choices[choices$list_name == list_main, , drop = FALSE]
    codigos_main <- as.character(ch_main$name)
    labels_main  <- as.character(ch_main$label)

  } else {
    # Fallback: usar valores únicos en data
    codigos_main <- sort(unique(as.character(data[[var]])))
    labels_main  <- codigos_main
  }

  map_main <- stats::setNames(labels_main, codigos_main)

  # Filtrar NA y códigos perdidos en variable principal
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
    stop("No hay datos válidos para la variable '", var, "'.", call. = FALSE)
  }

  # Si no hay cruce → una sola barra ("Total")
  if (is.null(var_cruce) || !nzchar(var_cruce)) {

    df_tab <- df |>
      dplyr::count(.data[[var]], name = "n") |>
      dplyr::mutate(
        pct          = n / sum(n),
        opcion_code  = as.character(.data[[var]]),
        opcion_label = map_main[opcion_code] %||% opcion_code,
        estrato_label = "Total"
      ) |>
      dplyr::select(estrato_label, opcion_label, pct, n)

    # Ordenar categorías según labels_main
    orden_lvls <- map_main[codigos_main]
    df_tab$opcion_label <- factor(
      df_tab$opcion_label,
      levels = unique(orden_lvls[!is.na(orden_lvls)])
    )
    df_tab <- df_tab[order(df_tab$opcion_label), , drop = FALSE]

    return(df_tab)
  }

  # ---------------------------------------------------------
  # Con cruce
  # ---------------------------------------------------------
  if (!var_cruce %in% names(df)) {
    stop("La variable de cruce '", var_cruce, "' no existe en `data`.",
         call. = FALSE)
  }

  df[[var_cruce]] <- as.character(df[[var_cruce]])

  # Mapeo código → etiqueta para la variable de cruce (si aplica)
  fila_cruce <- survey[survey$name == var_cruce, , drop = FALSE]
  list_cruce <- if (nrow(fila_cruce)) fila_cruce$list_name[1] else NA_character_

  if (!is.null(choices) &&
      all(c("list_name", "name", "label") %in% names(choices)) &&
      !is.na(list_cruce) && nzchar(list_cruce)) {

    ch_cruce <- choices[choices$list_name == list_cruce, , drop = FALSE]
    map_cruce <- stats::setNames(
      as.character(ch_cruce$label),
      as.character(ch_cruce$name)
    )
  } else {
    # Fallback: usar los valores tal cual
    niveles_cruce <- sort(unique(df[[var_cruce]]))
    map_cruce <- stats::setNames(niveles_cruce, niveles_cruce)
  }

  df_tab <- df |>
    dplyr::count(.data[[var_cruce]], .data[[var]], name = "n") |>
    dplyr::group_by(.data[[var_cruce]]) |>
    dplyr::mutate(
      pct = n / sum(n)
    ) |>
    dplyr::ungroup() |>
    dplyr::mutate(
      opcion_code   = as.character(.data[[var]]),
      opcion_label  = map_main[opcion_code] %||% opcion_code,
      estrato_code  = as.character(.data[[var_cruce]]),
      estrato_label = map_cruce[estrato_code] %||% estrato_code
    ) |>
    dplyr::select(estrato_label, opcion_label, pct, n)

  # Orden sensible: primero por estrato_label, luego por orden de labels_main
  orden_lvls_main <- map_main[codigos_main]
  df_tab$opcion_label <- factor(
    df_tab$opcion_label,
    levels = unique(orden_lvls_main[!is.na(orden_lvls_main)])
  )

  df_tab$estrato_label <- factor(
    df_tab$estrato_label,
    levels = sort(unique(df_tab$estrato_label))
  )

  df_tab[order(df_tab$estrato_label, df_tab$opcion_label), , drop = FALSE]
}

#' Construir gráfico plotly a partir de la tabla larga de proporciones
#'
#' @keywords internal
#' @noRd
.construir_plotly_barras <- function(df_tab,
                                     titulo,
                                     paleta_colores = NULL,
                                     height = 600) {

  if (!requireNamespace("plotly", quietly = TRUE)) {
    stop("Se requiere el paquete 'plotly' para `reporte_interactivo()`.",
         call. = FALSE)
  }

  # Asegurar proporciones y conteos válidos
  df_tab$pct[is.na(df_tab$pct)] <- 0
  df_tab$pct[df_tab$pct < 0] <- 0
  df_tab$pct[df_tab$pct > 1] <- 1
  df_tab$n[is.na(df_tab$n)]   <- 0

  # Re-normalizar para que los porcentajes enteros sumen 100% por estrato
  df_tab <- df_tab |>
    dplyr::group_by(.data$estrato_label) |>
    dplyr::mutate(
      porc_raw = ifelse(sum(pct) > 0, pct / sum(pct), 0),
      porc_int = round(porc_raw * 100),
      diff     = 100 - sum(porc_int),
      porc_int = ifelse(
        dplyr::row_number() == which.max(pct),
        porc_int + diff,
        porc_int
      )
    ) |>
    dplyr::ungroup()

  df_tab$texto_pct      <- paste0(df_tab$porc_int, "%")
  df_tab$texto_pct_html <- paste0("<b>", df_tab$porc_int, "%</b>")

  # Ordenar niveles
  opcion_levels  <- levels(df_tab$opcion_label)
  estrato_levels <- levels(df_tab$estrato_label)

  if (is.null(opcion_levels)) {
    opcion_levels <- unique(df_tab$opcion_label)
  }
  if (is.null(estrato_levels)) {
    estrato_levels <- unique(df_tab$estrato_label)
  }

  df_tab$opcion_label  <- factor(df_tab$opcion_label,  levels = opcion_levels)
  df_tab$estrato_label <- factor(df_tab$estrato_label, levels = estrato_levels)

  opcion_levels <- levels(df_tab$opcion_label)

  # Paleta de colores
  if (is.null(paleta_colores) || length(paleta_colores) == 0L) {
    n_cols <- max(3L, length(opcion_levels))
    paleta_colores <- grDevices::hcl.colors(n_cols, palette = "Blues")
  }

  if (length(paleta_colores) < length(opcion_levels)) {
    paleta_colores <- rep(paleta_colores, length.out = length(opcion_levels))
  }

  names(paleta_colores) <- opcion_levels

  # Construir gráfico
  p <- plotly::plot_ly()

  for (opt in opcion_levels) {
    df_opt <- df_tab[df_tab$opcion_label == opt, , drop = FALSE]

    if (nrow(df_opt) == 0L) next

    p <- p |>
      plotly::add_bars(
        data        = df_opt,
        x           = ~pct,
        y           = ~estrato_label,
        name        = as.character(opt),
        orientation = "h",
        text        = ~texto_pct_html,
        textposition = "inside",
        insidetextanchor = "middle",
        textfont    = list(
          color = "white",
          size  = 11
        ),
        customdata = ~n,
        marker      = list(
          color = paleta_colores[as.character(opt)],
          line  = list(width = 0)
        ),
        hovertemplate = paste0(
          "%{y}<br>",
          as.character(opt), ": %{text}<br>",
          "N: %{customdata}",
          "<extra></extra>"
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
        title      = "",
        automargin = TRUE
      ),
      legend = list(
        orientation = "h",
        x           = 0.5,
        xanchor     = "center",
        y           = -0.12
      ),
      margin = list(
        l = 170,
        r = 30,
        t = 80,
        b = 40
      ),
      height = height,
      title  = list(
        text    = titulo,
        x       = 0,
        xanchor = "left"
      ),
      uniformtext = list(
        minsize = 10,
        mode    = "hide"
      ),
      hovermode = "closest",
      transition = list(
        duration = 350,
        easing   = "cubic-in-out"
      )
    ) |>
    plotly::config(
      displayModeBar = FALSE
    )

  p
}

# -----------------------------------------------------------------------------
# App principal: reporte_interactivo()
# -----------------------------------------------------------------------------

#' Explorador interactivo de resultados (fase 1)
#'
#' Genera una aplicación \pkg{shiny} que permite:
#' \itemize{
#'   \item Seleccionar una variable principal (categórica) y visualizar su
#'         distribución como barra horizontal apilada (100\%).
#'   \item Aplicar filtros sobre una variable categórica (checkboxes tipo chip).
#'   \item Cruzar la variable principal por una variable categórica (estratos),
#'         mostrando una barra por categoría del cruce.
#' }
#'
#' El título del gráfico siempre corresponde únicamente a la etiqueta de la
#' variable principal, independientemente de si hay cruce.
#'
#' @param data Base de reporte (idealmente salida de `reporte_data()`).
#' @param instrumento Objeto devuelto por `reporte_instrumento()`, que debe
#'   contener al menos el componente `survey` y, opcionalmente, `choices`.
#' @param secciones Lista nombrada de vectores de variables por sección,
#'   usada para poblar el selector de la variable principal.
#' @param fuente (Por ahora no se muestra, reservado para futuras versiones).
#' @param titulo Título principal de la aplicación.
#' @param colores_apiladas_por_listname Lista nombrada de paletas de
#'   colores por `list_name` del instrumento (para las barras apiladas).
#' @param codigos_perdidos Vector de códigos que deben excluirse de la
#'   variable principal (por ejemplo, 96, 97, 98, 99).
#' @param facet_vars Vector de nombres de variables categóricas candidatas
#'   a filtro y cruce.
#'
#' @return Un objeto \code{shiny.appobj}.
#' @export
#'
#' @importFrom dplyr count mutate group_by ungroup
#' @importFrom stats setNames
reporte_interactivo <- function(
    data,
    instrumento,
    secciones,
    fuente      = NULL,
    titulo      = "Explorador interactivo",
    colores_apiladas_por_listname = NULL,
    codigos_perdidos = NULL,
    facet_vars = NULL
) {

  if (!requireNamespace("shiny", quietly = TRUE) ||
      !requireNamespace("plotly", quietly = TRUE) ||
      !requireNamespace("dplyr",  quietly = TRUE)) {
    stop("Se requieren los paquetes 'shiny', 'plotly' y 'dplyr' para `reporte_interactivo()`.",
         call. = FALSE)
  }

  survey <- instrumento$survey
  if (is.null(survey) || !"name" %in% names(survey)) {
    stop("El `instrumento` debe contener un `survey` válido.", call. = FALSE)
  }

  # ----------------------- Variables disponibles ------------------------------
  vars_disponibles <- unique(unlist(secciones, use.names = FALSE))
  vars_disponibles <- vars_disponibles[vars_disponibles %in% names(data)]

  if (!length(vars_disponibles)) {
    stop("No hay variables disponibles en `secciones` que estén en `data`.",
         call. = FALSE)
  }

  label_var <- function(v) .obtener_label_var(v, instrumento, data)

  var_choices <- stats::setNames(
    vars_disponibles,
    vapply(vars_disponibles, label_var, character(1))
  )

  var_default <- vars_disponibles[1]

  # Candidatas a filtro / cruce
  facet_vars <- facet_vars %||% character(0)
  facet_vars <- facet_vars[facet_vars %in% names(data)]

  facet_choices <- stats::setNames(
    facet_vars,
    vapply(facet_vars, label_var, character(1))
  )

  # ------------------------------- UI ----------------------------------------
  ui <- shiny::fluidPage(
    shiny::titlePanel(title = titulo),

    shiny::sidebarLayout(
      shiny::sidebarPanel(
        width = 3,

        # ---------------------- VARIABLES -------------------------------------
        shiny::h3("Variables"),

        shiny::p("Seleccione la variable principal a visualizar."),

        shiny::selectInput(
          inputId  = "var_principal",
          label    = "Variable principal",
          choices  = var_choices,
          selected = var_default
        ),

        shiny::hr(),

        # ------------------------ FILTROS ------------------------------------
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

        # ------------------------- CRUCE -------------------------------------
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
        plotly::plotlyOutput("plot_principal", height = "620px")
      )
    )
  )

  # ------------------------------- SERVER ------------------------------------
  server <- function(input, output, session) {

    # --------- UI dinámico para categorías del filtro ------------------------
    output$filtro_categorias_ui <- shiny::renderUI({
      v <- input$filtro_var
      if (is.null(v) || !nzchar(v) || !v %in% names(data)) {
        return(NULL)
      }

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

    # --------- Limpiar filtros -----------------------------------------------
    shiny::observeEvent(input$limpiar_filtros, {
      v <- input$filtro_var
      if (is.null(v) || !nzchar(v) || !v %in% names(data)) {
        return()
      }
      vals <- sort(unique(as.character(data[[v]])))
      vals <- vals[!is.na(vals)]
      shiny::updateCheckboxGroupInput(
        session,
        inputId  = "filtro_categorias",
        selected = vals
      )
    })

    # --------- Limpiar cruce -------------------------------------------------
    shiny::observeEvent(input$limpiar_cruce, {
      shiny::updateSelectInput(
        session,
        inputId  = "var_cruce",
        selected = ""
      )
    })

    # --------- Data filtrada -------------------------------------------------
    data_filtrada <- shiny::reactive({
      df <- data

      v_filtro <- input$filtro_var

      if (!is.null(v_filtro) &&
          nzchar(v_filtro) &&
          v_filtro %in% names(df) &&
          !is.null(input$filtro_categorias)) {

        vals_sel <- input$filtro_categorias
        if (length(vals_sel) > 0L) {
          df <- df[df[[v_filtro]] %in% vals_sel, , drop = FALSE]
        }
      }

      df
    })

    # --------- Gráfico principal ---------------------------------------------
    output$plot_principal <- plotly::renderPlotly({
      req(input$var_principal)

      var_main  <- input$var_principal
      df        <- data_filtrada()
      var_cruce <- input$var_cruce
      if (!nzchar(var_cruce)) var_cruce <- NULL

      if (nrow(df) == 0L) {
        shiny::validate(
          shiny::need(FALSE, "No hay datos después de aplicar los filtros.")
        )
      }

      # Título SIEMPRE solo de la variable principal
      titulo_plot <- .wrap_titulo_html(
        .obtener_label_var(var_main, instrumento, data),
        width = 120
      )

      # Paleta por list_name si existe
      survey  <- instrumento$survey

      paleta <- NULL
      if (!is.null(colores_apiladas_por_listname) &&
          all(c("name", "list_name") %in% names(survey))) {

        ln_main <- survey$list_name[survey$name == var_main][1]
        if (!is.na(ln_main) && ln_main %in% names(colores_apiladas_por_listname)) {
          paleta <- colores_apiladas_por_listname[[ln_main]]
        }
      }

      df_tab <- .preparar_tabla_proporciones(
        data              = df,
        instrumento       = instrumento,
        var               = var_main,
        var_cruce         = var_cruce,
        codigos_perdidos  = codigos_perdidos
      )

      .construir_plotly_barras(
        df_tab          = df_tab,
        titulo          = titulo_plot,
        paleta_colores  = paleta,
        height          = 600
      )
    })
  }

  shiny::shinyApp(ui = ui, server = server)
}
