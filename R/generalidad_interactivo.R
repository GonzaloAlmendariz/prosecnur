# =============================================================================
# Explorador interactivo de resultados (reporte_interactivo)
# =============================================================================

# Pequeño helper genérico
`%||%` <- function(x, y) if (!is.null(x)) x else y

# -----------------------------------------------------------------------------
# Helpers internos para labels y tipo de variable
# -----------------------------------------------------------------------------

#' Obtener etiqueta legible de una variable
#'
#' Busca primero en `instrumento$survey$label`, luego en el atributo
#' `label` de la columna en `data`, y finalmente retorna el nombre
#' de la variable si no encuentra nada.
#'
#' @param var Nombre de variable (string).
#' @param instrumento Objeto del instrumento (con componente `survey`).
#' @param data (Opcional) data.frame con los datos, para buscar
#'   atributos de etiqueta.
#'
#' @return Un string con la etiqueta legible.
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

#' Determinar tipo de variable para el explorador interactivo
#'
#' Wrapper simple sobre `tipo_pregunta_spss()` para dejar abierta
#' la posibilidad de incluir lógica adicional específica del
#' explorador.
#'
#' @param var Nombre de variable.
#' @param data Base de datos.
#' @param instrumento Instrumento con componente `survey`.
#' @param sm_vars_force (Opcional) vector de variables que deben
#'   tratarse como `select_multiple`.
#'
#' @return Lista con elemento `tipo`.
#' @keywords internal
#' @noRd
tipo_var_interactivo <- function(var,
                                 data,
                                 instrumento,
                                 sm_vars_force = NULL) {
  survey <- instrumento$survey

  tipo <- tipo_pregunta_spss(
    var           = var,
    survey        = survey,
    sm_vars_force = sm_vars_force
  )

  list(tipo = tipo)
}

# -----------------------------------------------------------------------------
# Graficador: vista simple (una barra apilada con distribución total)
# -----------------------------------------------------------------------------

#' Construir gráfico ggplot para una variable (vista simple)
#'
#' Genera una barra horizontal apilada que representa la distribución
#' de la variable sobre el total de casos, usando `graficar_barras_apiladas()`.
#'
#' @param data Base de datos (filtrada).
#' @param instrumento Instrumento con `survey` y `orders_list`.
#' @param var Nombre de la variable a graficar.
#' @param fuente Texto de fuente para el pie de página.
#' @param colores_apiladas_por_listname Lista de paletas por `list_name`.
#' @param codigos_perdidos Vector de códigos a excluir para la variable.
#'
#' @return Un objeto `ggplot`.
#' @keywords internal
#' @noRd
build_gg_for_var <- function(data,
                             instrumento,
                             var,
                             fuente                         = NULL,
                             colores_apiladas_por_listname  = NULL,
                             codigos_perdidos               = NULL) {

  survey <- instrumento$survey
  orders <- instrumento$orders_list %||% NULL

  tipo_info <- tipo_var_interactivo(
    var           = var,
    data          = data,
    instrumento   = instrumento,
    sm_vars_force = NULL
  )
  tipo <- tipo_info$tipo

  if (tipo == "ninguno") {
    stop("La variable '", var, "' no se puede graficar en este explorador.",
         call. = FALSE)
  }

  d_var <- data

  # excluir códigos perdidos SOLO para esta variable
  if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0 &&
      var %in% names(d_var)) {

    col_vals <- d_var[[var]]
    if (is.numeric(col_vals)) {
      d_var <- d_var[!(col_vals %in% codigos_perdidos), , drop = FALSE]
    } else {
      d_var <- d_var[!(as.character(col_vals) %in% as.character(codigos_perdidos)),
                     , drop = FALSE]
    }
  }

  tab_freq <- freq_table_spss(
    data          = d_var,
    var           = var,
    survey        = survey,
    sm_vars_force = NULL,
    orders_list   = orders,
    mostrar_todo  = TRUE
  )

  if (!nrow(tab_freq)) {
    stop("La variable '", var, "' no tiene datos válidos tras filtros/aplicación de pesos.",
         call. = FALSE)
  }

  is_total <- tab_freq$Opciones == "Total"
  body     <- tab_freq[!is_total, , drop = FALSE]
  total    <- tab_freq[ is_total, , drop = FALSE]

  if (!nrow(body)) {
    stop("La variable '", var, "' solo tiene fila Total; no se puede graficar.",
         call. = FALSE)
  }

  # ---------------------------------------------------------------------------
  # Tabla ancha: UNA barra "Distribución"
  # ---------------------------------------------------------------------------
  dummy_cat   <- "Distribución"
  var_cat_col <- "categoria"

  wide <- body |>
    dplyr::mutate(
      !!var_cat_col := dummy_cat,
      pct = as.numeric(.data$pct)
    )

  cols_pct  <- paste0("pct_", seq_len(nrow(wide)))
  etiquetas <- as.character(wide$Opciones)

  for (i in seq_len(nrow(wide))) {
    nm_col <- cols_pct[i]
    wide[[nm_col]] <- 0
    wide[[nm_col]][1] <- wide$pct[i]
  }

  wide <- wide[1, c(var_cat_col, "n", cols_pct), drop = FALSE]
  wide$n[1] <- if (nrow(total)) as.numeric(total$n[1]) else sum(body$n, na.rm = TRUE)

  names(etiquetas) <- cols_pct

  titulo_var_plot <- .obtener_label_var(var, instrumento, data = data)

  nota_pie <- fuente %||% NULL
  nota_pie_derecha <- if (nrow(total)) {
    paste0("N total = ",
           format(total$n[1], big.mark = ",", scientific = FALSE))
  } else {
    NULL
  }

  # ---------------------------------------------------------------------------
  # Colores por list_name (etiquetas, no códigos)
  # ---------------------------------------------------------------------------
  colores_grupos <- NULL
  if (!is.null(colores_apiladas_por_listname) &&
      !is.null(survey) &&
      all(c("name", "list_name") %in% names(survey)) &&
      var %in% survey$name) {

    ln <- survey$list_name[survey$name == var][1]
    if (!is.na(ln) && ln %in% names(colores_apiladas_por_listname)) {
      pal <- colores_apiladas_por_listname[[ln]]
      if (!is.null(pal) && length(pal)) {
        n_etq <- length(etiquetas)
        if (length(pal) < n_etq) {
          pal <- rep(pal, length.out = n_etq)
        } else if (length(pal) > n_etq) {
          pal <- pal[seq_len(n_etq)]
        }
        names(pal) <- etiquetas
        colores_grupos <- pal
      }
    }
  }

  graficar_barras_apiladas(
    data                 = wide,
    var_categoria        = var_cat_col,
    var_n                = "n",
    cols_porcentaje      = cols_pct,
    etiquetas_grupos     = etiquetas,
    escala_valor         = "proporcion_1",
    colores_grupos       = colores_grupos,
    mostrar_valores      = TRUE,
    decimales            = 1,
    umbral_etiqueta      = 0.03,
    umbral_etiqueta_peq  = 0.015,
    mostrar_barra_extra  = TRUE,
    barra_extra_preset   = "totales",
    prefijo_barra_extra  = NULL,
    titulo_barra_extra   = NULL,
    barra_extra_vjust    = NULL,
    titulo               = titulo_var_plot,
    subtitulo            = NULL,
    nota_pie             = nota_pie,
    nota_pie_derecha     = nota_pie_derecha,
    pos_titulo           = "izquierda",
    pos_nota_pie         = "derecha",
    centro_cowplot       = NA_real_,
    color_titulo         = "#000000",
    size_titulo          = 11,
    color_subtitulo      = "#000000",
    size_subtitulo       = 9,
    color_nota_pie       = "#000000",
    size_nota_pie        = 8,
    color_leyenda        = "#000000",
    size_leyenda         = 8,
    color_texto_barras   = "white",
    size_texto_barras    = 3,
    size_texto_barras_peq = 2.5,
    color_barra_extra    = "#000000",
    size_barra_extra     = 3,
    color_ejes           = "#000000",
    size_ejes            = 9,
    color_fondo          = NA,
    grosor_barras        = 0.7,
    extra_derecha_rel    = 0.12,
    espacio_izquierda_rel = 0,
    ancho_max_eje_y      = NULL,
    mostrar_leyenda      = TRUE,
    usar_leyenda_cowplot = FALSE,   # importante para compatibilidad con plotly
    invertir_leyenda     = FALSE,
    invertir_barras      = FALSE,
    invertir_segmentos   = FALSE,
    textos_negrita       = c("titulo", "porcentajes", "barra_extra"),
    exportar             = "rplot"
  )
}

# -----------------------------------------------------------------------------
# Graficador: vista de cruces (una barra por estrato)
# -----------------------------------------------------------------------------

#' Construir gráfico ggplot para una variable cruzada por estrato
#'
#' Genera un gráfico de barras apiladas donde cada barra corresponde
#' a una categoría de la variable de cruce (estrato), y los segmentos
#' representan la distribución de la variable de interés.
#'
#' @param data Base de datos (filtrada).
#' @param instrumento Instrumento con `survey` y `orders_list`.
#' @param var Variable principal a graficar.
#' @param var_cruce Variable de cruce (estrato).
#' @param fuente Texto de fuente para el pie de página.
#' @param colores_apiladas_por_listname Lista de paletas por `list_name`.
#' @param codigos_perdidos Vector de códigos a excluir para la variable.
#'
#' @return Un objeto `ggplot`.
#' @keywords internal
#' @noRd
build_gg_for_var_cruce <- function(data,
                                   instrumento,
                                   var,
                                   var_cruce,
                                   fuente                         = NULL,
                                   colores_apiladas_por_listname  = NULL,
                                   codigos_perdidos               = NULL) {

  survey <- instrumento$survey
  orders <- instrumento$orders_list %||% NULL

  if (!var_cruce %in% names(data)) {
    stop("La variable de cruce '", var_cruce, "' no existe en la base.", call. = FALSE)
  }

  tipo_info <- tipo_var_interactivo(
    var           = var,
    data          = data,
    instrumento   = instrumento,
    sm_vars_force = NULL
  )
  tipo <- tipo_info$tipo
  if (tipo == "ninguno") {
    stop("La variable '", var, "' no se puede graficar en este explorador.",
         call. = FALSE)
  }

  d_var <- data

  # excluir códigos perdidos para la variable principal
  if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0 &&
      var %in% names(d_var)) {

    col_vals <- d_var[[var]]
    if (is.numeric(col_vals)) {
      d_var <- d_var[!(col_vals %in% codigos_perdidos), , drop = FALSE]
    } else {
      d_var <- d_var[!(as.character(col_vals) %in% as.character(codigos_perdidos)),
                     , drop = FALSE]
    }
  }

  # niveles del cruce
  cruce_vals    <- d_var[[var_cruce]]
  niveles_cruce <- sort(unique(cruce_vals[!is.na(cruce_vals)]))
  if (!length(niveles_cruce)) {
    stop("La variable de cruce '", var_cruce, "' no tiene categorías válidas.",
         call. = FALSE)
  }

  # categorías globales de la variable (para asegurar mismo orden)
  tab_global <- freq_table_spss(
    data          = d_var,
    var           = var,
    survey        = survey,
    sm_vars_force = NULL,
    orders_list   = orders,
    mostrar_todo  = TRUE
  )
  is_total_global <- tab_global$Opciones == "Total"
  cats_global     <- as.character(tab_global$Opciones[!is_total_global])

  if (!length(cats_global)) {
    stop("La variable '", var, "' no tiene categorías de respuesta válidas.",
         call. = FALSE)
  }

  # ---------------------------------------------------------------------------
  # Construir tabla ancha: una fila por estrato (cruce)
  # ---------------------------------------------------------------------------
  lista_filas <- vector("list", length(niveles_cruce))

  for (i in seq_along(niveles_cruce)) {
    val_i <- niveles_cruce[i]
    sub_i <- d_var[d_var[[var_cruce]] == val_i, , drop = FALSE]

    tab_i <- freq_table_spss(
      data          = sub_i,
      var           = var,
      survey        = survey,
      sm_vars_force = NULL,
      orders_list   = orders,
      mostrar_todo  = TRUE
    )

    if (!nrow(tab_i)) {
      next
    }

    is_total_i <- tab_i$Opciones == "Total"
    body_i     <- tab_i[!is_total_i, , drop = FALSE]
    total_i    <- tab_i[ is_total_i, , drop = FALSE]

    if (!nrow(body_i)) next

    # asegurar que estén todas las cats_global en el mismo orden
    body_i <- body_i |>
      dplyr::filter(.data$Opciones %in% cats_global)

    pct_vec   <- rep(0, length(cats_global))
    match_idx <- match(cats_global, body_i$Opciones)
    pct_vec[!is.na(match_idx)] <- body_i$pct[match_idx]

    df_row <- tibble::tibble(
      categoria = as.character(val_i),
      n        = if (nrow(total_i)) as.numeric(total_i$n[1]) else sum(body_i$n, na.rm = TRUE)
    )

    for (j in seq_along(cats_global)) {
      nm_col <- paste0("pct_", j)
      df_row[[nm_col]] <- pct_vec[j]
    }

    lista_filas[[i]] <- df_row
  }

  wide <- dplyr::bind_rows(lista_filas)
  if (!nrow(wide)) {
    stop("No hay datos válidos para los cruces de '", var, "' por '", var_cruce, "'.",
         call. = FALSE)
  }

  cols_pct  <- paste0("pct_", seq_along(cats_global))
  etiquetas <- cats_global
  names(etiquetas) <- cols_pct

  titulo_var_plot <- paste0(
    .obtener_label_var(var, instrumento, data = data),
    " según ",
    .obtener_label_var(var_cruce, instrumento, data = data)
  )

  nota_pie <- fuente %||% NULL

  nota_pie_derecha <- paste0(
    "Cruce por ",
    .obtener_label_var(var_cruce, instrumento, data = data)
  )

  # Colores por list_name, igual que en vista simple
  colores_grupos <- NULL
  if (!is.null(colores_apiladas_por_listname) &&
      !is.null(survey) &&
      all(c("name", "list_name") %in% names(survey)) &&
      var %in% survey$name) {

    ln <- survey$list_name[survey$name == var][1]
    if (!is.na(ln) && ln %in% names(colores_apiladas_por_listname)) {
      pal <- colores_apiladas_por_listname[[ln]]
      if (!is.null(pal) && length(pal)) {
        n_etq <- length(etiquetas)
        if (length(pal) < n_etq) {
          pal <- rep(pal, length.out = n_etq)
        } else if (length(pal) > n_etq) {
          pal <- pal[seq_len(n_etq)]
        }
        names(pal) <- etiquetas
        colores_grupos <- pal
      }
    }
  }

  graficar_barras_apiladas(
    data                 = wide,
    var_categoria        = "categoria",
    var_n                = "n",
    cols_porcentaje      = cols_pct,
    etiquetas_grupos     = etiquetas,
    escala_valor         = "proporcion_1",
    colores_grupos       = colores_grupos,
    mostrar_valores      = TRUE,
    decimales            = 1,
    umbral_etiqueta      = 0.03,
    umbral_etiqueta_peq  = 0.015,
    mostrar_barra_extra  = TRUE,
    barra_extra_preset   = "totales",
    prefijo_barra_extra  = NULL,
    titulo_barra_extra   = NULL,
    barra_extra_vjust    = NULL,
    titulo               = titulo_var_plot,
    subtitulo            = NULL,
    nota_pie             = nota_pie,
    nota_pie_derecha     = nota_pie_derecha,
    pos_titulo           = "izquierda",
    pos_nota_pie         = "derecha",
    centro_cowplot       = NA_real_,
    color_titulo         = "#000000",
    size_titulo          = 11,
    color_subtitulo      = "#000000",
    size_subtitulo       = 9,
    color_nota_pie       = "#000000",
    size_nota_pie        = 8,
    color_leyenda        = "#000000",
    size_leyenda         = 8,
    color_texto_barras   = "white",
    size_texto_barras    = 3,
    size_texto_barras_peq = 2.5,
    color_barra_extra    = "#000000",
    size_barra_extra     = 3,
    color_ejes           = "#000000",
    size_ejes            = 9,
    color_fondo          = NA,
    grosor_barras        = 0.7,
    extra_derecha_rel    = 0.12,
    espacio_izquierda_rel = 0,
    ancho_max_eje_y      = 40,
    mostrar_leyenda      = TRUE,
    usar_leyenda_cowplot = FALSE,
    invertir_leyenda     = FALSE,
    invertir_barras      = FALSE,
    invertir_segmentos   = FALSE,
    textos_negrita       = c("titulo", "porcentajes", "barra_extra"),
    exportar             = "rplot"
  )
}

# -----------------------------------------------------------------------------
# Conversión ggplot → plotly con transición animada
# -----------------------------------------------------------------------------

#' Convertir un ggplot a plotly con transición suave
#'
#' Envuelve `plotly::ggplotly()` y añade una configuración básica
#' de `transition` para suavizar los cambios cuando se actualiza el
#' gráfico en la app.
#'
#' @param p Objeto `ggplot`.
#'
#' @return Objeto `plotly` (o `p` si no es `ggplot`).
#' @keywords internal
#' @noRd
gg_to_plotly_interactivo <- function(p) {
  if (!inherits(p, "gg")) return(p)

  pl <- plotly::ggplotly(p)

  pl <- pl |>
    plotly::layout(
      transition = list(
        duration = 400,
        easing   = "cubic-in-out"
      )
    )

  pl
}

# -----------------------------------------------------------------------------
# App principal: reporte_interactivo()
# -----------------------------------------------------------------------------

#' Explorador interactivo de resultados tipo OPS/ACNUR
#'
#' Genera una aplicación \pkg{shiny} que permite:
#' \itemize{
#'   \item Aplicar filtros dinámicos sobre variables seleccionadas.
#'   \item Elegir hasta cuatro variables para graficar simultáneamente.
#'   \item Alternar entre vista simple (distribución total) y vista de
#'         cruces (una barra por estrato).
#'   \item Visualizar los gráficos como barras horizontales apiladas con
#'         etiquetas de porcentaje y N total.
#' }
#'
#' El layout de la sección de configuración de gráficos se mantiene
#' fijo (cuatro selectores para variables), y el panel principal
#' ajusta dinámicamente el número y disposición de gráficos:
#' \itemize{
#'   \item 1 gráfico: ocupa todo el ancho.
#'   \item 2 gráficos: mitad – mitad.
#'   \item 3–4 gráficos: cuadrícula 2 × 2.
#' }
#'
#' @param data Base de reporte (idealmente salida de `reporte_data()`).
#' @param instrumento Objeto devuelto por `reporte_instrumento()`, que debe
#'   contener al menos el componente `survey` y, opcionalmente,
#'   `orders_list`.
#' @param secciones Lista nombrada de vectores de variables por sección,
#'   usada para poblar los selectores de variables de los gráficos.
#' @param fuente Texto de fuente a mostrar en los pies de los gráficos.
#' @param titulo Título principal de la app.
#' @param colores_apiladas_por_listname Lista nombrada de paletas de
#'   colores por `list_name` del instrumento (para las barras apiladas).
#' @param codigos_perdidos Vector de códigos que deben excluirse de las
#'   variables categóricas (por ejemplo, 96, 97, 98, 99).
#' @param facet_vars Vector de nombres de variables que se usarán como
#'   candidatas a filtro y a variable de cruce (estratos).
#'
#' @return Un objeto \code{shiny.appobj} que puede ejecutarse con
#'   `print()` o `shiny::runApp()`.
#'
#' @importFrom dplyr filter mutate all_of
#' @importFrom tidyr pivot_longer
#' @importFrom tibble tibble
#' @importFrom stats setNames
#' @export
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
      !requireNamespace("plotly", quietly = TRUE)) {
    stop("Se requieren los paquetes 'shiny' y 'plotly' para `reporte_interactivo()`.",
         call. = FALSE)
  }

  survey <- instrumento$survey
  if (is.null(survey) || !"name" %in% names(survey)) {
    stop("El `instrumento` debe contener un `survey` válido.", call. = FALSE)
  }

  # ---------------------------------------------------------------------------
  # Labels amigables para desplegar variables
  # ---------------------------------------------------------------------------
  survey_labels <- survey |>
    dplyr::filter(!is.na(.data$name), .data$name != "") |>
    dplyr::mutate(
      label = ifelse(is.na(.data$label) | .data$label == "",
                     .data$name,
                     as.character(.data$label))
    )

  label_var <- function(v) {
    if (!v %in% survey_labels$name) return(v)
    survey_labels$label[survey_labels$name == v][1]
  }

  # ---------------------------------------------------------------------------
  # Variables candidatas a filtro y a cruce
  # ---------------------------------------------------------------------------
  facet_vars <- facet_vars %||% character(0)
  facet_vars <- facet_vars[facet_vars %in% names(data)]

  filtro_vars_default <- facet_vars
  filtro_choices <- stats::setNames(
    filtro_vars_default,
    vapply(filtro_vars_default, label_var, character(1))
  )

  cruce_choices <- stats::setNames(
    facet_vars,
    vapply(facet_vars, label_var, character(1))
  )

  # ---------------------------------------------------------------------------
  # Variables para los 4 gráficos (secciones)
  # ---------------------------------------------------------------------------
  vars_disponibles <- unique(unlist(secciones, use.names = FALSE))
  vars_disponibles <- vars_disponibles[vars_disponibles %in% names(data)]

  var_choices <- stats::setNames(
    vars_disponibles,
    vapply(vars_disponibles, label_var, character(1))
  )

  default_vars <- head(vars_disponibles, 4)
  default_vars <- c(default_vars, rep(NA_character_, 4 - length(default_vars)))
  names(default_vars) <- paste0("var_", 1:4)

  # ---------------------------------------------------------------------------
  # UI
  # ---------------------------------------------------------------------------
  ui <- shiny::fluidPage(
    shiny::titlePanel(title = titulo),

    shiny::sidebarLayout(
      shiny::sidebarPanel(
        width = 3,

        # ------------------ CONFIGURACIÓN DE GRÁFICOS (ARRIBA) ------------------
        shiny::h3("Configuración de gráficos"),

        shiny::p("Seleccione hasta cuatro variables para graficar."),

        shiny::selectInput(
          inputId = "var_1",
          label   = "Gráfico 1",
          choices = c("Ninguno" = "", var_choices),
          selected = default_vars[1] %||% ""
        ),
        shiny::selectInput(
          inputId = "var_2",
          label   = "Gráfico 2",
          choices = c("Ninguno" = "", var_choices),
          selected = default_vars[2] %||% ""
        ),
        shiny::selectInput(
          inputId = "var_3",
          label   = "Gráfico 3",
          choices = c("Ninguno" = "", var_choices),
          selected = default_vars[3] %||% ""
        ),
        shiny::selectInput(
          inputId = "var_4",
          label   = "Gráfico 4",
          choices = c("Ninguno" = "", var_choices),
          selected = default_vars[4] %||% ""
        ),

        shiny::hr(),

        # ------------------ FILTROS (ABAJO) ------------------
        shiny::h3("Filtros"),

        shiny::selectizeInput(
          inputId = "filtro_vars",
          label   = "Variables de filtro",
          choices = filtro_choices,
          multiple = TRUE,
          options = list(
            placeholder = "Seleccione una o más variables de filtro"
          )
        ),

        shiny::uiOutput("filtros_dinamicos"),

        shiny::hr(),

        # ------------------ CRUCES (ABAJO) ------------------
        shiny::h3("Cruces"),

        shiny::checkboxInput(
          inputId = "activar_cruce",
          label   = "Mostrar cruces por estrato",
          value   = FALSE
        ),

        shiny::selectInput(
          inputId = "var_cruce",
          label   = "Variable de cruce",
          choices = cruce_choices,
          selected = if (length(facet_vars)) facet_vars[1] else NULL
        ),

        shiny::helpText(
          "Cuando se activa, cada gráfico muestra una barra por estrato ",
          "de la variable de cruce seleccionada."
        )
      ),

      shiny::mainPanel(
        width = 9,
        shiny::uiOutput("plots_panel")
      )
    )
  )

  # ---------------------------------------------------------------------------
  # Server
  # ---------------------------------------------------------------------------
  server <- function(input, output, session) {

    # --------- filtros dinámicos UI ---------
    output$filtros_dinamicos <- shiny::renderUI({
      req(input$filtro_vars)

      purrr::map(input$filtro_vars, function(v) {
        vals <- sort(unique(data[[v]]))
        vals <- vals[!is.na(vals)]

        shiny::selectInput(
          inputId = paste0("filtro_", v),
          label   = label_var(v),
          choices = vals,
          selected = vals,
          multiple = TRUE
        )
      })
    })

    # --------- data filtrada ---------
    data_filtrada <- shiny::reactive({
      df <- data

      if (!is.null(input$filtro_vars) && length(input$filtro_vars) > 0) {
        for (v in input$filtro_vars) {
          id_f <- paste0("filtro_", v)
          vals_sel <- input[[id_f]]
          if (!is.null(vals_sel) && length(vals_sel)) {
            df <- df[df[[v]] %in% vals_sel, , drop = FALSE]
          }
        }
      }

      df
    })

    # --------- indicador de modo cruce ---------
    modo_cruce <- shiny::reactive({
      isTRUE(input$activar_cruce) &&
        !is.null(input$var_cruce) &&
        nzchar(input$var_cruce)
    })

    # --------- helper para un gráfico dado ---------
    build_plot_for_var <- function(var_input) {
      req(var_input)
      if (!nzchar(var_input)) {
        return(NULL)
      }

      df <- data_filtrada()

      if (modo_cruce()) {
        p <- build_gg_for_var_cruce(
          data        = df,
          instrumento = instrumento,
          var         = var_input,
          var_cruce   = input$var_cruce,
          fuente      = fuente,
          colores_apiladas_por_listname = colores_apiladas_por_listname,
          codigos_perdidos = codigos_perdidos
        )
      } else {
        p <- build_gg_for_var(
          data        = df,
          instrumento = instrumento,
          var         = var_input,
          fuente      = fuente,
          colores_apiladas_por_listname = colores_apiladas_por_listname,
          codigos_perdidos = codigos_perdidos
        )
      }

      gg_to_plotly_interactivo(p)
    }

    # --------- render de cada plot (1–4) ---------
    output$plot_1 <- plotly::renderPlotly({
      req(input$var_1)
      if (!nzchar(input$var_1)) return(NULL)
      build_plot_for_var(input$var_1)
    })

    output$plot_2 <- plotly::renderPlotly({
      req(input$var_2)
      if (!nzchar(input$var_2)) return(NULL)
      build_plot_for_var(input$var_2)
    })

    output$plot_3 <- plotly::renderPlotly({
      req(input$var_3)
      if (!nzchar(input$var_3)) return(NULL)
      build_plot_for_var(input$var_3)
    })

    output$plot_4 <- plotly::renderPlotly({
      req(input$var_4)
      if (!nzchar(input$var_4)) return(NULL)
      build_plot_for_var(input$var_4)
    })

    # --------- layout dinámico 1–4 gráficos ---------
    output$plots_panel <- shiny::renderUI({
      vars <- list(input$var_1, input$var_2, input$var_3, input$var_4)
      activos_idx <- which(vapply(vars, function(x) nzchar(x %||% ""), logical(1)))
      n_activos   <- length(activos_idx)

      if (n_activos == 0) {
        return(
          shiny::div(
            shiny::br(),
            shiny::p("Seleccione al menos una variable en la configuración de gráficos.")
          )
        )
      }

      ids <- paste0("plot_", activos_idx)

      if (n_activos == 1) {
        plotly::plotlyOutput(ids[1], height = "520px")

      } else if (n_activos == 2) {
        shiny::fluidRow(
          shiny::column(6, plotly::plotlyOutput(ids[1], height = "480px")),
          shiny::column(6, plotly::plotlyOutput(ids[2], height = "480px"))
        )

      } else {
        # 3 o 4 → cuadrícula 2x2 (referencia fija a plot_1:plot_4)
        shiny::fluidRow(
          shiny::column(6, plotly::plotlyOutput("plot_1", height = "420px")),
          shiny::column(6, plotly::plotlyOutput("plot_2", height = "420px")),
          shiny::column(6, plotly::plotlyOutput("plot_3", height = "420px")),
          shiny::column(6, plotly::plotlyOutput("plot_4", height = "420px"))
        )
      }
    })
  }

  shiny::shinyApp(ui = ui, server = server)
}
