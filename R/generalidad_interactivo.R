# =============================================================================
# Explorador interactivo: reporte_interactivo()
# Fase 2 – Gráfico principal + tabla + bloque de perfil (N + 2 KPIs)
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
                  choices$name      == value, , drop = FALSE]

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

#' Recalcular porcentajes ENTEROS por estrato (suma exacta = 100)
#'
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
      if (k > 0) {
        base[ord[seq_len(k)]] <- base[ord[seq_len(k)]] + 1L
      }
    } else if (rem < 0) {
      ord <- order(frac, decreasing = FALSE, na.last = NA)
      k   <- min(-rem, length(ord))
      if (k > 0) {
        base[ord[seq_len(k)]] <- pmax(0L, base[ord[seq_len(k)]] - 1L)
      }
    }

    df_g$porc_raw <- pct_norm
    df_g$porc_int <- base
    df_g
  })

  dplyr::bind_rows(df_list)
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

    ch_main      <- choices[choices$list_name == list_main, , drop = FALSE]
    codigos_main <- as.character(ch_main$name)
    labels_main  <- as.character(ch_main$label)

  } else {
    codigos_main <- sort(unique(as.character(data[[var]])))
    labels_main  <- codigos_main
  }

  map_main <- stats::setNames(labels_main, codigos_main)

  df <- data
  if (!var %in% names(df)) {
    stop("La variable '", var, "' no existe en `data`.", call. = FALSE)
  }

  df[[var]] <- as.character(df[[var]])
  df        <- df[!is.na(df[[var]]), , drop = FALSE]

  if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0) {
    df <- df[!(df[[var]] %in% as.character(codigos_perdidos)), , drop = FALSE]
  }

  if (nrow(df) == 0L) {
    stop("No hay datos válidos para la variable '", var, "'.", call. = FALSE)
  }

  # ------------------------ SIN CRUCE ----------------------------------------
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
    df_tab$opcion_label <- factor(
      df_tab$opcion_label,
      levels = unique(orden_lvls[!is.na(orden_lvls)])
    )
    df_tab <- df_tab[order(df_tab$opcion_label), , drop = FALSE]

    return(df_tab)
  }

  # ------------------------ CON CRUCE ----------------------------------------
  if (!var_cruce %in% names(df)) {
    stop("La variable de cruce '", var_cruce, "' no existe en `data`.",
         call. = FALSE)
  }

  df[[var_cruce]] <- as.character(df[[var_cruce]])

  fila_cruce <- survey[survey$name == var_cruce, , drop = FALSE]
  list_cruce <- if (nrow(fila_cruce)) fila_cruce$list_name[1] else NA_character_

  if (!is.null(choices) &&
      all(c("list_name", "name", "label") %in% names(choices)) &&
      !is.na(list_cruce) && nzchar(list_cruce)) {

    ch_cruce  <- choices[choices$list_name == list_cruce, , drop = FALSE]
    map_cruce <- stats::setNames(
      as.character(ch_cruce$label),
      as.character(ch_cruce$name)
    )
  } else {
    niveles_cruce <- sort(unique(df[[var_cruce]]))
    map_cruce     <- stats::setNames(niveles_cruce, niveles_cruce)
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

  orden_lvls_main <- map_main[codigos_main]
  df_tab$opcion_label <- factor(
    df_tab$opcion_label,
    levels = unique(orden_lvls_main[!is.na(orden_lvls_main)])
  )

  df_tab$estrato_label <- factor(
    df_tab$estrato_label,
    levels = sort(unique(df_tab$estrato_label))
  )

  # Parche por si solo hubiera un estrato "Total"
  if (length(unique(df_tab$estrato_label)) == 1 &&
      unique(as.character(df_tab$estrato_label)) %in% c("Total", "TOTAL", "total")) {

    df_tab$estrato_label <- factor(rep("", nrow(df_tab)))
  }

  df_tab[order(df_tab$estrato_label, df_tab$opcion_label), , drop = FALSE]
}

#' Construir tabla resumen para mostrar en bloque 2
#'
#' @keywords internal
#' @noRd
.construir_tabla_resumen <- function(df_tab) {

  if (!requireNamespace("dplyr", quietly = TRUE)) {
    stop("Se requiere 'dplyr' para la tabla resumen.", call. = FALSE)
  }

  df_tab <- .anotar_porcentajes_enteros(df_tab)

  if (all(as.character(df_tab$estrato_label) %in% c("", NA))) {
    out <- df_tab |>
      dplyr::arrange(opcion_label) |>
      dplyr::transmute(
        Respuesta  = as.character(.data$opcion_label),
        N          = .data$n,
        Porcentaje = paste0(.data$porc_int, "%")
      )
  } else {
    out <- df_tab |>
      dplyr::arrange(estrato_label, opcion_label) |>
      dplyr::transmute(
        Estrato    = as.character(.data$estrato_label),
        Respuesta  = as.character(.data$opcion_label),
        N          = .data$n,
        Porcentaje = paste0(.data$porc_int, "%")
      )
  }

  out
}

#' Graficador principal (barras apiladas horizontales) → plotly
#'
#' @keywords internal
#' @noRd
.construir_plotly_barras <- function(df_tab,
                                     titulo,
                                     paleta_colores = NULL,
                                     height = NULL,
                                     mostrar_leyenda = TRUE) {

  if (!requireNamespace("plotly", quietly = TRUE)) {
    stop("Se requiere el paquete 'plotly' para `reporte_interactivo()`.",
         call. = FALSE)
  }

  df_tab$pct[is.na(df_tab$pct)] <- 0
  df_tab$pct[df_tab$pct < 0]    <- 0
  df_tab$n[is.na(df_tab$n)]     <- 0

  df_tab <- .anotar_porcentajes_enteros(df_tab)

  df_tab$texto_pct      <- paste0(df_tab$porc_int, "%")
  df_tab$texto_pct_html <- paste0("<b>", df_tab$porc_int, "%</b>")

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

  # Paleta
  if (is.null(paleta_colores) || length(paleta_colores) == 0L) {
    n_cols <- max(3L, length(opcion_levels))
    paleta_colores <- grDevices::hcl.colors(n_cols, palette = "Blues")
  }

  if (length(paleta_colores) < length(opcion_levels)) {
    paleta_colores <- rep(paleta_colores, length.out = length(opcion_levels))
  }

  names(paleta_colores) <- opcion_levels

  n_estratos <- length(unique(df_tab$estrato_label))
  if (is.null(height)) {
    height <- max(220, min(650, 160 + 60 * n_estratos))
  }

  # ¿Solo una barra "Total" (estrato vacío)?
  solo_total <- all(as.character(df_tab$estrato_label) %in% c("", NA))

  # Título y márgenes distintos para gráfico grande vs KPI
  if (mostrar_leyenda) {
    # Gráfico principal
    titulo_font_size  <- 14
    titulo_margin_top <- 60
    # reducción fuerte cuando sólo hay una barra para evitar franja blanca
    margin_left       <- if (solo_total) 20 else 170
    margin_right      <- 30
    margin_bottom     <- 40
  } else {
    # KPI horizontal (si se usara)
    titulo_font_size  <- 11
    titulo_margin_top <- 30
    margin_left       <- if (solo_total) 20 else 120
    margin_right      <- 10
    margin_bottom     <- 25
  }

  p <- plotly::plot_ly(height = height)

  for (opt in opcion_levels) {
    df_opt <- df_tab[df_tab$opcion_label == opt, , drop = FALSE]
    if (nrow(df_opt) == 0L) next

    # Hover: sin "Total" explícito cuando no hay cruce
    if (solo_total) {
      df_opt$hover_text <- sprintf(
        "%s: %s<br>N: %s",
        as.character(opt),
        df_opt$texto_pct,
        df_opt$n
      )
    } else {
      df_opt$hover_text <- sprintf(
        "%s<br>%s: %s<br>N: %s",
        as.character(df_opt$estrato_label),
        as.character(opt),
        df_opt$texto_pct,
        df_opt$n
      )
    }

    p <- p |>
      plotly::add_bars(
        data             = df_opt,
        x                = ~pct,
        y                = ~estrato_label,
        name             = as.character(opt),
        orientation      = "h",
        text             = ~texto_pct_html,
        textposition     = "inside",
        insidetextanchor = "middle",
        textfont         = list(
          color = "white",
          size  = 11
        ),
        customdata       = ~hover_text,
        hovertemplate    = "%{customdata}<extra></extra>",
        marker           = list(
          color = paleta_colores[as.character(opt)],
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
        x           = 0.5,
        xanchor     = "center",
        y           = -0.12
      ),
      margin = list(
        l = margin_left,
        r = margin_right,
        t = titulo_margin_top,
        b = margin_bottom
      ),
      title  = list(
        text    = titulo,
        x       = 0,
        xanchor = "left",
        font    = list(size = titulo_font_size)
      ),
      uniformtext = list(
        minsize = 10,
        mode    = "hide"
      ),
      hovermode  = "closest",
      showlegend = mostrar_leyenda,
      transition = list(
        duration = 400,
        easing   = "cubic-in-out"
      )
    ) |>
    plotly::config(
      displayModeBar = FALSE
    )

  # Opciones de animación generales
  p <- plotly::animation_opts(
    p,
    frame      = 1000,
    transition = 400,
    easing     = "cubic-in-out",
    redraw     = TRUE
  )

  p
}

# -----------------------------------------------------------------------------
# App principal
# -----------------------------------------------------------------------------

#' Explorador interactivo de resultados (fase 2)
#'
#' Genera una aplicación \pkg{shiny} con tres bloques:
#' \itemize{
#'   \item Bloque 1 (superior): gráfico principal de barras horizontales
#'         apiladas al 100\% para una variable categórica seleccionada,
#'         con opción de cruce por otra variable categórica.
#'   \item Bloque 2 (inferior izquierda): tabla resumen de frecuencias y
#'         porcentajes (coherente con el gráfico y los filtros aplicados),
#'         en un contenedor de altura fija con scroll interno.
#'   \item Bloque 3 (inferior derecha): perfil dinámico de la muestra para
#'         la pregunta seleccionada, que incluye una tarjeta con el N de
#'         casos válidos de la pregunta y hasta dos gráficos KPI (barras
#'         apiladas verticales sin leyenda) para variables definidas en
#'         \code{kpi_vars}.
#' }
#'
#' El perfil se calcula siempre sobre los mismos casos que alimentan el
#' gráfico principal: sólo casos con respuesta válida (no NA) en la
#' variable principal, después de aplicar filtros.
#'
#' @param data Base de reporte (idealmente salida de `reporte_data()`).
#' @param instrumento Objeto devuelto por `reporte_instrumento()`, que debe
#'   contener al menos el componente `survey` y, opcionalmente, `choices`.
#' @param secciones Lista nombrada de vectores de variables por sección,
#'   usada para poblar el selector de sección y variable principal.
#' @param fuente (Por ahora no se muestra, reservado para futuras versiones).
#' @param titulo Título principal de la aplicación.
#' @param colores_apiladas_por_listname Lista nombrada de paletas de
#'   colores por `list_name` del instrumento (para las barras apiladas).
#' @param codigos_perdidos Vector de códigos que deben excluirse de la
#'   variable principal y de los KPIs (por ejemplo, 96, 97, 98, 99).
#' @param facet_vars Vector de nombres de variables categóricas candidatas
#'   a filtro y cruce.
#' @param id_unidad Nombre de la variable que identifica la unidad de
#'   análisis (por ejemplo, código de EESS). Si se especifica, el N del
#'   perfil corresponde al número de unidades distintas; en caso contrario,
#'   corresponde al número de filas.
#' @param kpi_vars Vector con 0, 1 o 2 nombres de variables categóricas a
#'   mostrar como KPIs en el Bloque 3. Si hay más de 2, se usan solo las
#'   dos primeras.
#'
#' @return Un objeto \code{shiny.appobj}.
#' @export
#'
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
    stop("Se requieren 'shiny', 'plotly' y 'dplyr' para `reporte_interactivo()`.",
         call. = FALSE)
  }

  survey <- instrumento$survey
  if (is.null(survey) || !"name" %in% names(survey)) {
    stop("El `instrumento` debe contener un `survey` válido.", call. = FALSE)
  }

  usa_DT <- requireNamespace("DT", quietly = TRUE)

  if (is.null(secciones) || !length(secciones)) {
    stop("`secciones` debe ser una lista nombrada con vectores de variables.",
         call. = FALSE)
  }

  # Secciones: sólo variables presentes en data
  secciones_limpias <- lapply(secciones, function(v) v[v %in% names(data)])
  secciones_limpias <- secciones_limpias[vapply(secciones_limpias, length, integer(1)) > 0]

  if (!length(secciones_limpias)) {
    stop("Ninguna sección de `secciones` tiene variables presentes en `data`.",
         call. = FALSE)
  }

  secciones_nombres <- names(secciones_limpias)

  label_var <- function(v) .obtener_label_var(v, instrumento, data)

  # Filtros / cruces
  facet_vars <- facet_vars %||% character(0)
  facet_vars <- facet_vars[facet_vars %in% names(data)]

  facet_choices <- stats::setNames(
    facet_vars,
    vapply(facet_vars, label_var, character(1))
  )

  # KPIs (máx 2)
  kpi_vars <- kpi_vars %||% character(0)
  kpi_vars <- kpi_vars[kpi_vars %in% names(data)]
  kpi_vars <- unique(kpi_vars)
  if (length(kpi_vars) > 2L) {
    kpi_vars <- kpi_vars[1:2]
  }

  # -------------------------------- UI ---------------------------------------
  ui <- shiny::fluidPage(
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

        # Bloque 1: gráfico principal (ancho completo)
        shiny::fluidRow(
          shiny::column(
            width = 12,
            plotly::plotlyOutput("plot_principal", height = "420px")
          )
        ),

        shiny::br(),

        # Bloque 2 (tabla) y Bloque 3 (perfil) abajo, misma altura + scroll
        shiny::fluidRow(
          shiny::column(
            width = 6,
            shiny::div(
              style = "height: 360px; overflow-y: auto; border: 1px solid #eee; border-radius: 6px; padding: 5px;",
              if (usa_DT) {
                DT::dataTableOutput("tabla_principal")
              } else {
                shiny::tableOutput("tabla_principal")
              }
            )
          ),
          shiny::column(
            width = 6,
            shiny::div(
              style = paste(
                "height: 360px;",  # mismo alto que la tabla
                "border: 1px solid #eee; border-radius: 6px; padding: 5px;",
                # Flex en columna, pero que los hijos USEN TODO EL ANCHO
                "display: flex; flex-direction: column; align-items: stretch;",
                # No recortar tooltips de plotly
                "overflow-y: visible; overflow-x: visible;"
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

    # Actualizar variable principal según sección
    shiny::observe({
      sec      <- input$seccion
      vars_sec <- secciones_limpias[[sec]]

      if (is.null(vars_sec) || !length(vars_sec)) {
        shiny::updateSelectInput(
          session,
          inputId  = "var_principal",
          choices  = c(),
          selected = ""
        )
      } else {
        choices_sec <- stats::setNames(
          vars_sec,
          vapply(vars_sec, label_var, character(1))
        )
        shiny::updateSelectInput(
          session,
          inputId  = "var_principal",
          choices  = choices_sec,
          selected = vars_sec[1]
        )
      }
    })

    # UI dinámico de categorías del filtro
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

    # Limpiar filtros
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

    # Limpiar cruce
    shiny::observeEvent(input$limpiar_cruce, {
      shiny::updateSelectInput(
        session,
        inputId = "var_cruce",
        selected = ""
      )
    })

    # Data filtrada (según filtro_var + categorías)
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

    # ------------------ Bloque 1: gráfico principal --------------------------
    output$plot_principal <- plotly::renderPlotly({
      req(input$var_principal)

      var_main <- input$var_principal
      df_all   <- data_filtrada()

      # PERFIL dinámico: sólo casos que responden la pregunta principal
      if (var_main %in% names(df_all)) {
        df <- df_all[!is.na(df_all[[var_main]]), , drop = FALSE]
      } else {
        df <- df_all
      }

      var_cruce <- input$var_cruce
      if (!nzchar(var_cruce)) var_cruce <- NULL

      if (nrow(df) == 0L) {
        shiny::validate(
          shiny::need(FALSE, "No hay datos válidos para la pregunta seleccionada (después de filtros).")
        )
      }

      titulo_plot <- .wrap_titulo_html(
        .obtener_label_var(var_main, instrumento, data),
        width = 120
      )

      survey_local <- instrumento$survey

      paleta <- NULL
      if (!is.null(colores_apiladas_por_listname) &&
          all(c("name", "list_name") %in% names(survey_local))) {

        ln_main <- survey_local$list_name[survey_local$name == var_main][1]
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
        height          = 420,
        mostrar_leyenda = TRUE
      )
    })

    # ------------------ Bloque 2: tabla resumen ------------------------------
    if (usa_DT) {

      output$tabla_principal <- DT::renderDataTable({
        req(input$var_principal)
        var_main <- input$var_principal
        df_all   <- data_filtrada()

        # mismo criterio que el gráfico: sólo casos que responden la pregunta
        if (var_main %in% names(df_all)) {
          df <- df_all[!is.na(df_all[[var_main]]), , drop = FALSE]
        } else {
          df <- df_all
        }

        var_cruce <- input$var_cruce
        if (!nzchar(var_cruce)) var_cruce <- NULL

        if (nrow(df) == 0L) {
          return(NULL)
        }

        df_tab <- .preparar_tabla_proporciones(
          data              = df,
          instrumento       = instrumento,
          var               = var_main,
          var_cruce         = var_cruce,
          codigos_perdidos  = codigos_perdidos
        )

        tabla <- .construir_tabla_resumen(df_tab)

        DT::datatable(
          tabla,
          rownames = FALSE,
          options  = list(
            paging    = FALSE,
            searching = FALSE,
            info      = FALSE
          )
        )
      })

    } else {

      output$tabla_principal <- shiny::renderTable({
        req(input$var_principal)
        var_main <- input$var_principal
        df_all   <- data_filtrada()

        # mismo criterio que el gráfico: sólo casos que responden la pregunta
        if (var_main %in% names(df_all)) {
          df <- df_all[!is.na(df_all[[var_main]]), , drop = FALSE]
        } else {
          df <- df_all
        }

        var_cruce <- input$var_cruce
        if (!nzchar(var_cruce)) var_cruce <- NULL

        if (nrow(df) == 0L) {
          return(NULL)
        }

        df_tab <- .preparar_tabla_proporciones(
          data              = df,
          instrumento       = instrumento,
          var               = var_main,
          var_cruce         = var_cruce,
          codigos_perdidos  = codigos_perdidos
        )

        .construir_tabla_resumen(df_tab)
      })
    }

    # ------------------ Bloque 3: perfil (N + KPIs) --------------------------
    output$kpi_panel <- shiny::renderUI({
      df_all   <- data_filtrada()
      var_main <- input$var_principal

      # PERFIL dinámico: sólo casos que responden la pregunta principal
      if (!is.null(var_main) && nzchar(var_main) && var_main %in% names(df_all)) {
        df <- df_all[!is.na(df_all[[var_main]]), , drop = FALSE]
      } else {
        df <- df_all
      }

      if (nrow(df) == 0L) {
        return(shiny::div("Sin datos para la pregunta seleccionada."))
      }

      # N dinámico: unidades o filas
      n_unidades <- if (!is.null(id_unidad) && id_unidad %in% names(df)) {
        dplyr::n_distinct(df[[id_unidad]])
      } else {
        nrow(df)
      }

      # Tarjeta de N centrada arriba
      tarjeta_N <- shiny::div(
        style = paste(
          "border: 1px solid #ddd; border-radius: 6px;",
          "padding: 10px; margin-bottom: 10px;",
          "background-color: #f9f9f9; text-align: center; width: 90%;"
        ),
        shiny::div(
          style = "font-size: 11px; text-transform: uppercase; color: #666; letter-spacing: 0.08em;",
          "N de casos (pregunta actual)"
        ),
        shiny::div(
          style = "font-size: 22px; font-weight: 700; color: #333; margin-top: 3px;",
          format(n_unidades, big.mark = ",", scientific = FALSE)
        )
      )

      kpi_elems <- list(tarjeta_N)

      # Helper para gráfico KPI (BARRA APILADA VERTICAL, con leyenda a la izquierda)
      construir_kpi_plot <- function(df, var_kpi) {
        if (!var_kpi %in% names(df)) return(NULL)

        df_kpi <- df[!is.na(df[[var_kpi]]), , drop = FALSE]
        if (!nrow(df_kpi)) return(NULL)

        df_tab_kpi <- .preparar_tabla_proporciones(
          data             = df_kpi,
          instrumento      = instrumento,
          var              = var_kpi,
          var_cruce        = NULL,
          codigos_perdidos = codigos_perdidos
        )

        df_tab_kpi <- .anotar_porcentajes_enteros(df_tab_kpi)

        titulo_kpi <- .wrap_titulo_html(
          .obtener_label_var(var_kpi, instrumento, data),
          width = 60
        )

        # Niveles de opción
        opcion_levels <- levels(df_tab_kpi$opcion_label)
        if (is.null(opcion_levels)) {
          opcion_levels <- unique(df_tab_kpi$opcion_label)
          df_tab_kpi$opcion_label <- factor(
            df_tab_kpi$opcion_label,
            levels = opcion_levels
          )
        }

        # --------- PALETA: respeta colores_apiladas_por_listname si aplica ----
        paleta <- NULL

        survey_local <- instrumento$survey
        if (!is.null(colores_apiladas_por_listname) &&
            !is.null(survey_local) &&
            all(c("name", "list_name") %in% names(survey_local))) {

          ln_kpi <- survey_local$list_name[survey_local$name == var_kpi][1]
          if (!is.na(ln_kpi) && ln_kpi %in% names(colores_apiladas_por_listname)) {
            paleta <- colores_apiladas_por_listname[[ln_kpi]]
          }
        }

        # Fallback si no hay paleta específica para este list_name
        if (is.null(paleta) || length(paleta) == 0L) {
          n_cols <- max(3L, length(opcion_levels))
          paleta <- grDevices::hcl.colors(n_cols, "Blues")
        }

        if (length(paleta) < length(opcion_levels)) {
          paleta <- rep(paleta, length.out = length(opcion_levels))
        }
        names(paleta) <- opcion_levels

        # Una sola barra apilada vertical (x = "KPI")
        df_tab_kpi$x_dummy <- "KPI"

        p <- plotly::plot_ly(height = 220)

        for (opt in opcion_levels) {
          df_opt <- df_tab_kpi[df_tab_kpi$opcion_label == opt, , drop = FALSE]

          p <- p |>
            plotly::add_bars(
              data             = df_opt,
              x                = ~x_dummy,
              y                = ~porc_int,
              name             = as.character(opt),         # para la leyenda
              marker           = list(color = paleta[as.character(opt)]),
              text             = ~paste0(porc_int, "%"),
              textposition     = "inside",
              insidetextanchor = "middle",
              hovertemplate    = paste0(
                "<b>", titulo_kpi, "</b><br>",
                "%{name}: %{text}<extra></extra>"
              )
            )
        }

        p <- p |>
          plotly::layout(
            barmode = "stack",
            title = list(
              text    = titulo_kpi,
              y       = 0.98,
              x       = 0.5,
              xanchor = "center",
              font    = list(size = 11)
            ),
            xaxis = list(
              title          = "",
              showticklabels = FALSE,
              showgrid       = FALSE,
              zeroline       = FALSE
            ),
            yaxis = list(
              title          = "",
              showticklabels = FALSE,
              showgrid       = FALSE,
              zeroline       = FALSE,
              range          = c(0, 100)
            ),
            # Leyenda vertical a la izquierda
            showlegend = F,
            legend = list(
              orientation = "v",
              x           = 0,
              xanchor     = "left",
              y           = 0.5,
              yanchor     = "middle",
              font        = list(size = 9)
            ),
            margin = list(l = 80, r = 10, t = 40, b = 20)
          ) |>
          plotly::config(displayModeBar = FALSE)

        p
      }

      # Render de KPIs
      if (length(kpi_vars) >= 1) {
        output$kpi_plot_1 <- plotly::renderPlotly({
          construir_kpi_plot(df = df, var_kpi = kpi_vars[1])
        })
      }
      if (length(kpi_vars) >= 2) {
        output$kpi_plot_2 <- plotly::renderPlotly({
          construir_kpi_plot(df = df, var_kpi = kpi_vars[2])
        })
      }

      # Fila de KPIs: uno al lado del otro
      if (length(kpi_vars) >= 1) {
        fila_kpis <- shiny::fluidRow(
          style = "width: 100%; margin-top: 10px;",
          shiny::column(
            width = if (length(kpi_vars) >= 2) 6 else 12,
            shiny::div(
              style = "width: 100%; padding: 0 5px;",
              plotly::plotlyOutput("kpi_plot_1", height = "220px", width = "100%")
            )
          ),
          if (length(kpi_vars) >= 2) {
            shiny::column(
              width = 6,
              shiny::div(
                style = "width: 100%; padding: 0 5px;",
                plotly::plotlyOutput("kpi_plot_2", height = "220px", width = "100%")
              )
            )
          }
        )

        kpi_elems[[length(kpi_elems) + 1]] <- fila_kpis
      }

      do.call(shiny::tagList, kpi_elems)
    })
  }

  shiny::shinyApp(ui = ui, server = server)
}
