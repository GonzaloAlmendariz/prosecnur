# =============================================================================
# reporte_ppt()
# -----------------------------------------------------------------------------
# Generar PPT con gráficos a partir de tablas de frecuencias (SO / SM)
# usando los graficadores ya existentes del paquete.
# =============================================================================

#' Generar un reporte PPT a partir de frecuencias (SO/SM) con graficadores propios
#'
#' `reporte_ppt()` toma una base de reporte (idealmente la misma que se usa en
#' `reporte_frecuencias()`), el instrumento y una definición de secciones, y
#' genera:
#' \itemize{
#'   \item Una lista de gráficos \code{ggplot} (uno por variable válida).
#'   \item Opcionalmente, un archivo PPT donde cada gráfico ocupa una diapositiva.
#' }
#'
#' El tipo de gráfico para cada variable se decide en tres niveles:
#' \enumerate{
#'   \item Sobrescritura explícita por variable:
#'     \itemize{
#'       \item \code{vars_dico}: variables que deben usarse como dicotómicas
#'             con \code{graficar_dico()}.
#'       \item \code{vars_barras_apiladas}: variables para \code{graficar_barras_apiladas()}.
#'       \item \code{vars_barras_agrupadas}: variables para \code{graficar_barras_agrupadas()}.
#'       \item \code{vars_radar}: reservado para futuros usos (por ahora no se
#'             construye tabla radar automáticamente).
#'     }
#'   \item Decisión por \code{list_name} de la pregunta en el instrumento:
#'     \itemize{
#'       \item \code{listnames_dico}: todas las variables cuyo \code{list_name}
#'             esté aquí tienden a graficarse como dicotómicas.
#'       \item \code{listnames_apiladas}: todas las variables cuyo \code{list_name}
#'             esté aquí tienden a graficarse como barras apiladas.
#'     }
#'   \item Defaults por tipo de pregunta:
#'     \itemize{
#'       \item \code{default_so}: tipo de gráfico por defecto para \code{select_one}.
#'       \item \code{default_sm}: tipo de gráfico por defecto para \code{select_multiple}.
#'     }
#' }
#'
#' Para las variables dicotómicas, \strong{no se asume ninguna pareja de categorías
#' por defecto}. En su lugar, se requiere que este par se defina explícitamente en:
#' \itemize{
#'   \item \code{dico_labels_por_var[[var]]} (prioridad más alta), o
#'   \item \code{dico_labels_por_listname[[list_name]]}.
#' }
#' Si no se encuentra una pareja de etiquetas válida, la variable se grafica con
#' el tipo de barras por defecto (típicamente \code{"barras_agrupadas"}), dejando
#' constancia en \code{log_decisiones}.
#'
#' Internamente, las tablas de frecuencias se construyen con
#' \code{freq_table_spss()}, por lo que solo se grafican variables que:
#' \itemize{
#'   \item Existen como columna en \code{data}, o
#'   \item Tienen dummies asociadas (\code{var/cod} o \code{var.cod}) para
#'         el caso de \code{select_multiple}.
#' }
#'
#' @param data Data frame o tibble con la base de reporte (idealmente el
#'   resultado de \code{reporte_data()}), que contiene los atributos \code{label}
#'   y \code{labels} en sus variables.
#' @param instrumento Objeto devuelto por \code{reporte_instrumento()}. Si es
#'   \code{NULL}, se intentará recuperar desde \code{attr(data, "instrumento_reporte")}.
#'   Debe contener al menos las tibbles \code{survey} y \code{choices}, y,
#'   opcionalmente, \code{orders_list}.
#' @param secciones Lista nombrada que define qué variables se incluyen en cada
#'   sección, de forma análoga a \code{reporte_frecuencias()}. Si es \code{NULL},
#'   se intentará inferir desde una columna \code{section} o \code{seccion} del
#'   \code{survey}.
#' @param path_ppt Ruta del archivo PPTX a generar cuando \code{solo_lista = FALSE}.
#' @param fuente Texto de fuente que se mostrará en el bloque de texto inferior
#'   izquierdo de cada diapositiva de gráficos (por ejemplo,
#'   `"Fuente: Pulso PUCP 2025"`). No se usa como caption interno del gráfico.
#' @param sm_vars_force Vector opcional de nombres de variables que deben tratarse
#'   como \code{select_multiple} aunque el instrumento no las marque como tales.
#' @param mostrar_todo Lógico; se pasa a \code{freq_table_spss()} como
#'   \code{mostrar_todo}. Controla si las tablas de frecuencias incluyen
#'   categorías con frecuencia 0 definidas en el instrumento. En todos los casos,
#'   el generador de PPT excluye categorías con \code{n == 0} al graficar.
#' @param solo_lista Lógico; si \code{TRUE}, no se genera PPT y solo se devuelve
#'   una lista con \code{plots} y \code{log_decisiones}. Si \code{FALSE}, además
#'   se construye \code{path_ppt}.
#' @param incluir_titulo_var Lógico; si \code{TRUE}, el label de la variable se
#'   usa como título de la diapositiva de gráfico (placeholder de título del
#'   layout de PowerPoint). Los gráficos en sí se generan sin título.
#' @param mensajes_progreso Lógico; si \code{TRUE}, muestra mensajes por sección
#'   y por variable indicando qué tipo de gráfico se está usando y por qué.
#'
#' @param vars_dico Vector de nombres de variables que se deben graficar como
#'   dicotómicas mediante \code{graficar_dico()}, siempre que exista una pareja
#'   de etiquetas en \code{dico_labels_por_var} o \code{dico_labels_por_listname}.
#' @param vars_barras_apiladas Vector de nombres de variables que se deben
#'   graficar como barras horizontales apiladas mediante
#'   \code{graficar_barras_apiladas()}.
#' @param vars_barras_agrupadas Vector de nombres de variables que se deben
#'   graficar como barras agrupadas mediante \code{graficar_barras_agrupadas()}.
#' @param vars_radar Vector de nombres de variables que se desean reservar para
#'   un tratamiento tipo radar. Actualmente se registra en el log, pero no se
#'   construye automáticamente una tabla para \code{graficar_radar()}.
#'
#' @param listnames_dico Vector de \code{list_name} (del \code{survey}) para los
#'   que, si una variable no aparece en \code{vars_dico}, se intentará por
#'   defecto tratarla como dicotómica.
#' @param listnames_apiladas Vector de \code{list_name} para los que, si una
#'   variable no aparece en \code{vars_barras_apiladas} ni en \code{vars_dico},
#'   se intentará por defecto graficar como barras apiladas.
#'
#' @param dico_labels_por_var Lista nombrada donde cada elemento es un vector
#'   de dos etiquetas \code{c("Positiva","Negativa")} para la variable
#'   correspondiente.
#' @param dico_labels_por_listname Lista nombrada por \code{list_name} donde cada
#'   elemento es un vector de dos etiquetas \code{c("Positiva","Negativa")} que
#'   se usan cuando la variable no tiene entrada propia en \code{dico_labels_por_var}.
#'
#' @param colores_apiladas_por_listname Lista nombrada por \code{list_name} donde
#'   cada elemento es un vector de colores HEX con nombre para los segmentos de
#'   las barras apiladas. Los nombres de estos colores deben coincidir con los
#'   labels de las categorías (\code{Opciones}) que se graficarán.
#'
#' @param default_so Tipo de gráfico por defecto para variables \code{select_one}
#'   cuando no se encuentren en ninguna lista de overrides. Puede ser
#'   \code{"barras_agrupadas"} o \code{"barras_apiladas"}.
#' @param default_sm Tipo de gráfico por defecto para variables \code{select_multiple}
#'   cuando no se encuentren en ninguna lista de overrides. Igual que
#'   \code{default_so}, puede ser \code{"barras_agrupadas"} o
#'   \code{"barras_apiladas"}.
#'
#' @param barra_extra Controla si se agrega o no una barra final con el N total
#'   (o algún agregado) en las barras apiladas/agrupadas. Puede ser:
#'   \itemize{
#'     \item \code{"ninguna"}: no se agrega barra extra.
#'     \item \code{"total_n"}: se agrega barra extra con el N total (\code{"N = ..."}).
#'   }
#'
#' @param estilos_barras_agrupadas Lista de parámetros de estilo que se pasan
#'   directamente a \code{graficar_barras_agrupadas()}.
#' @param estilos_barras_apiladas Lista de parámetros de estilo que se pasan
#'   directamente a \code{graficar_barras_apiladas()}.
#' @param estilos_dico Lista de parámetros de estilo que se pasan directamente a
#'   \code{graficar_dico()}.
#'
#' @param template_pptx Ruta a una plantilla PPTX (por ejemplo, en formato 16:9).
#'   Si es \code{NULL}, se intentará usar una plantilla interna del paquete
#'   llamada \code{"plantillas/plantilla_16_9.pptx"}; si tampoco existe, se
#'   usará la plantilla por defecto de PowerPoint a través de \code{officer::read_pptx()}.
#'
#' @param titulo_portada Título que se colocará en la diapositiva de portada
#'   (layout `"Title Slide"`) cuando exista dicho layout en la plantilla.
#' @param subtitulo_portada Subtítulo para la diapositiva de portada.
#' @param fecha_portada Texto de fecha que se colocará en el placeholder de
#'   fecha (`type = "dt"`) de la portada y, si existe, de la contraportada.
#' @param mostrar_resumen_n Lógico; si \code{TRUE}, en cada diapositiva de
#'   gráficos se escribe en el bloque de texto inferior derecho un resumen del
#'   tipo `"N = X | Ratio de respuestas: Y%"`.
#'
#' En plantillas que tengan un layout llamado \code{"Graficos"} con un
#' marcador de imagen (\code{type = "pic"}) y dos placeholders de texto
#' \code{type = "body"} (índices 2 y 3), los gráficos se insertan en dicho
#' marcador, el texto de \code{fuente} va al bloque izquierdo y el resumen de N
#' va al bloque derecho. En caso contrario, se usa una diapositiva en blanco a
#' pantalla completa y no se insertan esos textos.
#'
#' Si la plantilla incluye un layout \code{"Contraportada"}, se agrega una
#' diapositiva final usando ese layout.
#'
#' @return Una lista con dos elementos:
#' \describe{
#'   \item{plots}{Lista de objetos \code{ggplot} generados (uno por variable).}
#'   \item{log_decisiones}{Tibble con información sobre cada variable procesada:
#'         sección, nombre de variable, tipo de pregunta, \code{list_name},
#'         origen de la decisión (\code{override}) y tipo de gráfico final.}
#' }
#'
#' @export
reporte_ppt <- function(
    data,
    instrumento      = NULL,
    secciones        = NULL,
    path_ppt         = "reporte_ppt.pptx",
    fuente           = NULL,
    sm_vars_force    = NULL,
    mostrar_todo     = FALSE,
    solo_lista       = FALSE,
    incluir_titulo_var = TRUE,
    mensajes_progreso  = TRUE,

    # Sobrescritura por variable
    vars_dico             = NULL,
    vars_barras_apiladas  = NULL,
    vars_barras_agrupadas = NULL,
    vars_radar            = NULL,

    # Overrides por list_name
    listnames_dico        = NULL,
    listnames_apiladas    = NULL,

    # Etiquetas explícitas para dicotómicas
    dico_labels_por_var      = list(),
    dico_labels_por_listname = list(),

    # Colores para barras apiladas por list_name
    colores_apiladas_por_listname = list(),

    # Defaults por tipo de pregunta
    default_so = c("barras_agrupadas", "barras_apiladas"),
    default_sm = c("barras_agrupadas", "barras_apiladas"),

    # Barra extra en barras apiladas/agrupadas
    barra_extra = c("ninguna", "total_n"),

    # Estilos por tipo de gráfico
    estilos_barras_agrupadas = list(),
    estilos_barras_apiladas  = list(),
    estilos_dico             = list(),

    # Plantilla PPT
    template_pptx = NULL,

    # Texto de portada
    titulo_portada    = NULL,
    subtitulo_portada = NULL,
    fecha_portada     = NULL,

    # Resumen de N en bloque derecho
    mostrar_resumen_n = TRUE
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y

  default_so  <- match.arg(default_so)
  default_sm  <- match.arg(default_sm)
  barra_extra <- match.arg(barra_extra)

  vars_dico             <- vars_dico             %||% character(0)
  vars_barras_apiladas  <- vars_barras_apiladas  %||% character(0)
  vars_barras_agrupadas <- vars_barras_agrupadas %||% character(0)
  vars_radar            <- vars_radar            %||% character(0)

  listnames_dico     <- listnames_dico     %||% character(0)
  listnames_apiladas <- listnames_apiladas %||% character(0)

  if (!is.data.frame(data)) {
    stop("`data` debe ser un data.frame o tibble.", call. = FALSE)
  }

  # ---------------------------------------------------------------------------
  # 1. Instrumento y survey / choices / orders_list
  # ---------------------------------------------------------------------------
  if (is.null(instrumento)) {
    instrumento <- attr(data, "instrumento_reporte", exact = TRUE)
    if (is.null(instrumento)) {
      stop("No se proporcionó `instrumento` y `data` no tiene atributo ",
           "`instrumento_reporte`.", call. = FALSE)
    }
  }

  survey  <- instrumento$survey
  choices <- instrumento$choices
  if (is.null(survey) || !all(c("name", "label") %in% names(survey))) {
    stop("El `instrumento` no contiene un `survey` con columnas `name` y `label`.",
         call. = FALSE)
  }

  orders_list <- instrumento$orders_list %||% NULL

  dic_vars <- survey |>
    dplyr::filter(!is.na(.data$name), .data$name != "") |>
    dplyr::select(name, label) |>
    dplyr::mutate(label = trimws(as.character(.data$label))) |>
    dplyr::distinct(name, .keep_all = TRUE)

  # ---------------------------------------------------------------------------
  # 2. Inferir secciones si no se pasan
  # ---------------------------------------------------------------------------
  if (is.null(secciones)) {
    seccion_col <- NULL
    if ("section" %in% names(survey)) {
      seccion_col <- "section"
    } else if ("seccion" %in% names(survey)) {
      seccion_col <- "seccion"
    }

    if (is.null(seccion_col)) {
      stop("No se especificaron `secciones` y el `survey` no tiene columna ",
           "`section` ni `seccion`.", call. = FALSE)
    }

    secciones_df <- survey |>
      dplyr::filter(
        !is.na(.data[[seccion_col]]),
        !is.na(.data$name),
        .data$name %in% names(data)
      ) |>
      dplyr::select(seccion = !!rlang::sym(seccion_col), name)

    if (nrow(secciones_df) == 0) {
      stop("No se pudieron inferir secciones desde `survey$",
           seccion_col, "`.", call. = FALSE)
    }

    secciones <- split(secciones_df$name, secciones_df$seccion)
  }

  SECCIONES <- lapply(secciones, function(vars) {
    vars[vapply(
      vars,
      function(v) .has_var_or_dummies(data, v),
      logical(1)
    )]
  })
  SECCIONES <- SECCIONES[vapply(SECCIONES, length, integer(1)) > 0L]

  if (length(SECCIONES) == 0L) {
    stop("Después de filtrar por presencia en `data` (variable o dummies), ",
         "ninguna sección tiene variables válidas.", call. = FALSE)
  }

  # ---------------------------------------------------------------------------
  # 3. Helpers internos
  # ---------------------------------------------------------------------------

  .titulo_var_safe <- function(var) {
    titulo_var(
      var,
      dic_vars        = dic_vars,
      labels_override = NULL,
      orders_list     = orders_list,
      df              = data
    )
  }

  .tab_freq_var <- function(var) {
    tab <- freq_table_spss(
      data,
      var,
      survey        = survey,
      sm_vars_force = sm_vars_force,
      orders_list   = orders_list,
      mostrar_todo  = mostrar_todo
    )

    if (!nrow(tab)) return(tab)

    tab |>
      dplyr::filter(.data$Opciones != "Total") |>
      dplyr::filter(is.na(.data$n) == FALSE & .data$n > 0)
  }

  .build_tab_barras_agrupadas <- function(tab_freq, var_label) {
    if (!nrow(tab_freq)) return(NULL)

    n_total <- sum(tab_freq$n, na.rm = TRUE)

    tibble::tibble(
      categoria = tab_freq$Opciones,
      n_base    = n_total,
      pct       = tab_freq$pct
    )
  }

  .build_tab_barras_apiladas <- function(tab_freq, var_label) {
    if (!nrow(tab_freq)) return(NULL)

    n_total <- sum(tab_freq$n, na.rm = TRUE)
    n_cat   <- nrow(tab_freq)

    cols_pct <- paste0("pct_", seq_len(n_cat))
    df_wide  <- tibble::tibble(
      categoria = var_label %||% "",
      n_base    = n_total
    )

    for (i in seq_len(n_cat)) {
      df_wide[[cols_pct[i]]] <- tab_freq$pct[i]
    }

    etiquetas_grupos <- stats::setNames(as.character(tab_freq$Opciones), cols_pct)

    list(
      data             = df_wide,
      cols_porcentaje  = cols_pct,
      etiquetas_grupos = etiquetas_grupos
    )
  }

  .build_tab_dico <- function(tab_freq, var, var_label, labels_dico) {
    if (length(labels_dico) < 2) return(NULL)

    pos_lab <- labels_dico[1]
    neg_lab <- labels_dico[2]

    sub <- tab_freq |>
      dplyr::filter(.data$Opciones %in% c(pos_lab, neg_lab))

    if (nrow(sub) < 2) {
      warning(
        "En la variable '", var, "' no se encontraron ambas categorías ",
        "indicadas en `labels_dico`. Se omitirá este tratamiento dicotómico.",
        call. = FALSE
      )
      return(NULL)
    }

    n_pos <- sub$n[sub$Opciones == pos_lab][1]
    n_neg <- sub$n[sub$Opciones == neg_lab][1]
    denom <- n_pos + n_neg

    if (!is.finite(denom) || denom <= 0) return(NULL)

    pct_si <- n_pos / denom * 100

    indicador_val <- if (incluir_titulo_var) "" else (var_label %||% var)

    tibble::tibble(
      indicador = indicador_val,
      pct_si    = pct_si,
      n_total   = denom
    )
  }

  # ---------------------------------------------------------------------------
  # 4. Recorrido por secciones y variables
  # ---------------------------------------------------------------------------
  plots_list      <- list()
  titulos_list    <- list()
  resumenN_list   <- list()
  log_list        <- list()

  total_casos <- nrow(data)

  for (sec in names(SECCIONES)) {
    vars_sec <- SECCIONES[[sec]]

    if (mensajes_progreso) {
      message("Procesando sección: ", sec)
    }

    for (v in vars_sec) {

      tipo_v <- tipo_pregunta_spss(v, survey, sm_vars_force)
      if (tipo_v == "so_or_open") tipo_v <- "so"

      list_name_v <- NA_character_
      if ("list_name" %in% names(survey)) {
        tmp <- survey$list_name[survey$name == v]
        if (length(tmp)) list_name_v <- tmp[1]
      } else if ("list_norm" %in% names(survey)) {
        tmp <- survey$list_norm[survey$name == v]
        if (length(tmp)) list_name_v <- tmp[1]
      }

      override     <- NA_character_
      tipo_grafico <- NULL

      if (v %in% vars_dico) {
        override     <- "vars_dico"
        tipo_grafico <- "dico"
      } else if (v %in% vars_barras_apiladas) {
        override     <- "vars_barras_apiladas"
        tipo_grafico <- "barras_apiladas"
      } else if (v %in% vars_barras_agrupadas) {
        override     <- "vars_barras_agrupadas"
        tipo_grafico <- "barras_agrupadas"
      } else if (v %in% vars_radar) {
        override     <- "vars_radar"
        tipo_grafico <- "radar"
      } else if (!is.na(list_name_v) && list_name_v %in% listnames_dico) {
        override     <- paste0("listnames_dico=", list_name_v)
        tipo_grafico <- "dico"
      } else if (!is.na(list_name_v) && list_name_v %in% listnames_apiladas) {
        override     <- paste0("listnames_apiladas=", list_name_v)
        tipo_grafico <- "barras_apiladas"
      } else {
        if (tipo_v == "sm") {
          override     <- paste0("default_sm=", default_sm)
          tipo_grafico <- default_sm
        } else {
          override     <- paste0("default_so=", default_so)
          tipo_grafico <- default_so
        }
      }

      tab_freq <- .tab_freq_var(v)
      if (!nrow(tab_freq)) {
        log_list[[length(log_list) + 1]] <- tibble::tibble(
          seccion      = sec,
          var          = v,
          tipo_var     = tipo_v,
          list_name    = list_name_v,
          override     = override,
          tipo_grafico = NA_character_
        )
        next
      }

      # --- resumen N y ratio de respuestas para esta variable ---
      n_var <- sum(tab_freq$n, na.rm = TRUE)
      if (is.finite(n_var) && n_var >= 0 && total_casos > 0) {
        ratio <- n_var / total_casos * 100
        resumen_n_txt <- sprintf("N = %s | Ratio de respuestas: %.1f%%",
                                 format(n_var, big.mark = ",", scientific = FALSE),
                                 ratio)
      } else if (is.finite(n_var)) {
        resumen_n_txt <- sprintf("N = %s",
                                 format(n_var, big.mark = ",", scientific = FALSE))
      } else {
        resumen_n_txt <- NULL
      }

      var_label    <- .titulo_var_safe(v)
      titulo_plot  <- NULL  # el gráfico va sin título
      titulo_slide <- if (incluir_titulo_var) var_label else NULL

      # No usamos caption interno del plot; la fuente se escribirá en el
      # placeholder de texto de la diapositiva.
      nota_pie_plot <- NULL

      if (mensajes_progreso) {
        message("   - ", v, " → ", tipo_grafico,
                " (list_name = ", list_name_v, ") [", override, "]")
      }

      p <- NULL
      tipo_grafico_final <- tipo_grafico

      if (tipo_grafico %in% c("barras_agrupadas", "barras_apiladas", "dico")) {

        if (tipo_grafico == "barras_agrupadas") {

          tab_agr <- .build_tab_barras_agrupadas(tab_freq, var_label)
          if (is.null(tab_agr) || !nrow(tab_agr)) next

          cols_porcentaje  <- "pct"
          etiquetas_series <- c(pct = "Porcentaje")

          colores_series <- estilos_barras_agrupadas$colores_series %||%
            c("Porcentaje" = "#004B8D")

          args_barras <- c(
            list(
              data             = tab_agr,
              var_categoria    = "categoria",
              var_n            = "n_base",
              cols_porcentaje  = cols_porcentaje,
              etiquetas_series = etiquetas_series,
              escala_valor     = "proporcion_1",
              colores_series   = colores_series,
              mostrar_valores  = TRUE,
              titulo           = titulo_plot,
              subtitulo        = NULL,
              nota_pie         = nota_pie_plot,
              mostrar_barra_extra = barra_extra == "total_n",
              prefijo_barra_extra = if (barra_extra == "total_n") "N = " else "N = ",
              titulo_barra_extra  = if (barra_extra == "total_n") "Total" else NULL,
              exportar            = "rplot"
            ),
            estilos_barras_agrupadas
          )

          p <- do.call(graficar_barras_agrupadas, args_barras)
        }

        if (tipo_grafico == "barras_apiladas") {

          tab_apil <- .build_tab_barras_apiladas(tab_freq, var_label)
          if (is.null(tab_apil)) next

          colores_grupos <- NULL
          if (!is.na(list_name_v) &&
              !is.null(colores_apiladas_por_listname[[list_name_v]])) {
            colores_grupos <- colores_apiladas_por_listname[[list_name_v]]
          }

          args_apiladas <- c(
            list(
              data             = tab_apil$data,
              var_categoria    = "categoria",
              var_n            = "n_base",
              cols_porcentaje  = tab_apil$cols_porcentaje,
              etiquetas_grupos = tab_apil$etiquetas_grupos,
              escala_valor     = "proporcion_1",
              colores_grupos   = colores_grupos,
              mostrar_valores  = TRUE,
              titulo           = titulo_plot,
              subtitulo        = NULL,
              nota_pie         = nota_pie_plot,
              mostrar_barra_extra = barra_extra == "total_n",
              prefijo_barra_extra = if (barra_extra == "total_n") "N = " else "N = ",
              titulo_barra_extra  = NULL,
              exportar            = "rplot"
            ),
            estilos_barras_apiladas
          )

          p <- do.call(graficar_barras_apiladas, args_apiladas)

          p <- p +
            ggplot2::theme(
              axis.text.y  = ggplot2::element_blank(),
              axis.title.y = ggplot2::element_blank()
            )
        }

        if (tipo_grafico == "dico") {

          labels_dico <- NULL
          if (!is.null(dico_labels_por_var[[v]])) {
            labels_dico <- dico_labels_por_var[[v]]
          } else if (!is.na(list_name_v) &&
                     !is.null(dico_labels_por_listname[[list_name_v]])) {
            labels_dico <- dico_labels_por_listname[[list_name_v]]
          }

          if (is.null(labels_dico) || length(labels_dico) < 2) {
            warning(
              "En la variable '", v, "' no se encontraron etiquetas dicotómicas ",
              "en `dico_labels_por_var` ni `dico_labels_por_listname`. ",
              "Se usará barras agrupadas.",
              call. = FALSE
            )
            tipo_grafico_final <- "barras_agrupadas"

            tab_agr <- .build_tab_barras_agrupadas(tab_freq, var_label)
            if (is.null(tab_agr) || !nrow(tab_agr)) next

            cols_porcentaje  <- "pct"
            etiquetas_series <- c(pct = "Porcentaje")

            colores_series <- estilos_barras_agrupadas$colores_series %||%
              c("Porcentaje" = "#004B8D")

            args_barras <- c(
              list(
                data             = tab_agr,
                var_categoria    = "categoria",
                var_n            = "n_base",
                cols_porcentaje  = cols_porcentaje,
                etiquetas_series = etiquetas_series,
                escala_valor     = "proporcion_1",
                colores_series   = colores_series,
                mostrar_valores  = TRUE,
                titulo           = titulo_plot,
                subtitulo        = NULL,
                nota_pie         = nota_pie_plot,
                mostrar_barra_extra = barra_extra == "total_n",
                prefijo_barra_extra = if (barra_extra == "total_n") "N = " else "N = ",
                titulo_barra_extra  = if (barra_extra == "total_n") "Total" else NULL,
                exportar            = "rplot"
              ),
              estilos_barras_agrupadas
            )

            p <- do.call(graficar_barras_agrupadas, args_barras)

          } else {

            tab_dico <- .build_tab_dico(tab_freq, v, var_label, labels_dico)
            if (is.null(tab_dico) || !nrow(tab_dico)) next

            args_dico <- c(
              list(
                data              = tab_dico,
                var_indicador     = "indicador",
                var_porcentaje_si = "pct_si",
                var_n             = "n_total",
                escala_valor      = "proporcion_100",
                etiqueta_si       = labels_dico[1],
                etiqueta_no       = labels_dico[2],
                titulo            = titulo_plot,
                subtitulo         = NULL,
                nota_pie          = nota_pie_plot,
                incluir_n_en_titulo = FALSE,
                exportar          = "rplot"
              ),
              estilos_dico
            )

            p <- do.call(graficar_dico, args_dico)
          }
        }

      } else if (tipo_grafico == "radar") {

        warning(
          "Tipo de gráfico 'radar' señalado para la variable '", v,
          "', pero el constructor genérico aún no está implementado. ",
          "La variable se omitirá en este reporte.",
          call. = FALSE
        )
        next

      } else {
        warning(
          "Tipo de gráfico '", tipo_grafico, "' no reconocido para la variable '",
          v, "'. Se omitirá en este reporte.",
          call. = FALSE
        )
        next
      }

      log_list[[length(log_list) + 1]] <- tibble::tibble(
        seccion      = sec,
        var          = v,
        tipo_var     = tipo_v,
        list_name    = list_name_v,
        override     = override,
        tipo_grafico = tipo_grafico_final
      )

      plots_list[[length(plots_list) + 1]]       <- p
      titulos_list[[length(titulos_list) + 1]]   <- titulo_slide
      resumenN_list[[length(resumenN_list) + 1]] <- resumen_n_txt
    }
  }

  log_decisiones <- dplyr::bind_rows(log_list)
  # ---------------------------------------------------------------------------
  # 6. PPT
  # ---------------------------------------------------------------------------
  if (!solo_lista) {
    if (!requireNamespace("officer", quietly = TRUE) ||
        !requireNamespace("rvg", quietly = TRUE)) {
      stop(
        "Para exportar a PPT se requieren los paquetes 'officer' y 'rvg'.",
        call. = FALSE
      )
    }

    # 6.1. Leer plantilla
    if (is.null(template_pptx)) {
      template_interno <- system.file("plantillas/plantilla_16_9.pptx",
                                      package = "prosecnur")
      if (nzchar(template_interno) && file.exists(template_interno)) {
        if (mensajes_progreso) {
          message("Usando plantilla interna 16:9: ", template_interno)
        }
        doc <- officer::read_pptx(path = template_interno)
      } else {
        if (mensajes_progreso) {
          message(
            "No se encontró 'plantilla_16_9.pptx' dentro del paquete. ",
            "Se usará la plantilla por defecto de PowerPoint."
          )
        }
        doc <- officer::read_pptx()
      }
    } else {
      if (!file.exists(template_pptx)) {
        stop(
          "No se encontró el archivo de plantilla especificado en `template_pptx`: ",
          template_pptx,
          call. = FALSE
        )
      }
      if (mensajes_progreso) {
        message("Usando plantilla definida por el usuario: ", template_pptx)
      }
      doc <- officer::read_pptx(path = template_pptx)
    }

    # 6.2. Info de layouts
    layout_info <- tryCatch(
      officer::layout_summary(doc),
      error = function(e) NULL
    )

    tiene_layout_graficos      <- FALSE
    layout_graficos            <- "Blank"
    usar_pic_placeholder       <- FALSE
    tiene_layout_title_slide   <- FALSE
    tiene_layout_contraportada <- FALSE

    if (!is.null(layout_info)) {

      # Prioridad: Graficos2 > Graficos > Blank
      if ("Graficos2" %in% layout_info$layout) {
        tiene_layout_graficos <- TRUE
        layout_graficos       <- "Graficos2"
        usar_pic_placeholder  <- TRUE
      } else if ("Graficos" %in% layout_info$layout) {
        tiene_layout_graficos <- TRUE
        layout_graficos       <- "Graficos"
        usar_pic_placeholder  <- TRUE
      }

      if ("Title Slide" %in% layout_info$layout) {
        tiene_layout_title_slide <- TRUE
      }
      if ("Contraportada" %in% layout_info$layout) {
        tiene_layout_contraportada <- TRUE
      }
    }

    if (mensajes_progreso) {
      if (tiene_layout_graficos) {
        message(
          "Las diapositivas de gráficos usarán el layout '",
          layout_graficos, "'."
        )
      } else {
        message("No se encontró un layout 'Graficos' ni 'Graficos2'; se usará 'Blank' a pantalla completa.")
      }
    }

    # 6.3. Portada (Title Slide), si corresponde
    if (tiene_layout_title_slide &&
        ( !is.null(titulo_portada)    && nzchar(titulo_portada)    ||
          !is.null(subtitulo_portada) && nzchar(subtitulo_portada) ||
          !is.null(fecha_portada)     && nzchar(fecha_portada) )) {

      if (mensajes_progreso) {
        message("Agregando diapositiva de portada (Title Slide).")
      }

      doc <- officer::add_slide(
        doc,
        layout = "Title Slide",
        master = "Office Theme"
      )

      # Título: ctrTitle o title
      if (!is.null(titulo_portada) && nzchar(titulo_portada)) {
        loc_title <- tryCatch(
          officer::ph_location_type(type = "ctrTitle"),
          error = function(e) officer::ph_location_type(type = "title")
        )
        doc <- tryCatch(
          officer::ph_with(doc, titulo_portada, location = loc_title),
          error = function(e) {
            if (mensajes_progreso) {
              message("No se pudo escribir el título en la portada: ", e$message)
            }
            doc
          }
        )
      }

      # Subtítulo: subTitle
      if (!is.null(subtitulo_portada) && nzchar(subtitulo_portada)) {
        doc <- tryCatch(
          officer::ph_with(
            doc,
            subtitulo_portada,
            location = officer::ph_location_type(type = "subTitle")
          ),
          error = function(e) {
            if (mensajes_progreso) {
              message("No se pudo escribir el subtítulo en la portada: ", e$message)
            }
            doc
          }
        )
      }

      # Fecha: dt
      if (!is.null(fecha_portada) && nzchar(fecha_portada)) {
        doc <- tryCatch(
          officer::ph_with(
            doc,
            fecha_portada,
            location = officer::ph_location_type(type = "dt")
          ),
          error = function(e) {
            if (mensajes_progreso) {
              message("No se pudo escribir la fecha en la portada: ", e$message)
            }
            doc
          }
        )
      }
    } else if (!is.null(fecha_portada) || !is.null(titulo_portada) || !is.null(subtitulo_portada)) {
      if (mensajes_progreso && !tiene_layout_title_slide) {
        message("Se solicitaron textos de portada, pero la plantilla no tiene layout 'Title Slide'.")
      }
    }

    # 6.4. Diapositivas de gráficos
    if (length(plots_list)) {
      for (i in seq_along(plots_list)) {

        p        <- plots_list[[i]]
        st       <- titulos_list[[i]]   %||% NULL
        resumenN <- resumenN_list[[i]]  %||% NULL

        doc <- officer::add_slide(
          doc,
          layout = layout_graficos,
          master = "Office Theme"
        )

        # Escribir título de la diapositiva si hay y existe placeholder
        if (!is.null(st) && nzchar(st)) {
          loc_gtitle <- tryCatch(
            officer::ph_location_type(type = "title"),
            error = function(e) {
              tryCatch(
                officer::ph_location_type(type = "ctrTitle"),
                error = function(e2) NULL
              )
            }
          )
          if (!is.null(loc_gtitle)) {
            doc <- tryCatch(
              officer::ph_with(doc, st, location = loc_gtitle),
              error = function(e) {
                if (mensajes_progreso) {
                  message("No se pudo escribir el título en la diapositiva de gráficos: ", e$message)
                }
                doc
              }
            )
          }
        }

        # Insertar gráfico en placeholder de imagen o a pantalla completa
        if (usar_pic_placeholder) {
          loc_pic <- officer::ph_location_type(type = "pic")
        } else {
          loc_pic <- officer::ph_location_fullsize()
        }

        doc <- officer::ph_with(
          doc,
          rvg::dml(ggobj = p, bg = "transparent"),
          location = loc_pic
        )

        # Escribir fuente en bloque de texto izquierdo (body id = 2) si corresponde
        if (tiene_layout_graficos &&
            !is.null(fuente) && nzchar(fuente)) {

          loc_fuente <- tryCatch(
            officer::ph_location_type(type = "body", id = 2),
            error = function(e) NULL
          )

          if (!is.null(loc_fuente)) {
            doc <- tryCatch(
              officer::ph_with(
                doc,
                fuente,
                location = loc_fuente
              ),
              error = function(e) {
                if (mensajes_progreso) {
                  message("No se pudo escribir la fuente en el bloque izquierdo: ", e$message)
                }
                doc
              }
            )
          }
        }

        # Escribir resumen N en bloque de texto derecho (body id = 3) si corresponde
        if (tiene_layout_graficos &&
            mostrar_resumen_n &&
            !is.null(resumenN) && nzchar(resumenN)) {

          loc_resumen <- tryCatch(
            officer::ph_location_type(type = "body", id = 3),
            error = function(e) NULL
          )

          if (!is.null(loc_resumen)) {
            doc <- tryCatch(
              officer::ph_with(
                doc,
                resumenN,
                location = loc_resumen
              ),
              error = function(e) {
                if (mensajes_progreso) {
                  message("No se pudo escribir el resumen de N en el bloque derecho: ", e$message)
                }
                doc
              }
            )
          }
        }
      }
    }

    # 6.5. Contraportada (si existe)
    if (tiene_layout_contraportada) {
      if (mensajes_progreso) {
        message("Agregando diapositiva de contraportada.")
      }

      doc <- officer::add_slide(
        doc,
        layout = "Contraportada",
        master = "Office Theme"
      )

      # Si hay fecha definida, intentar ponerla en el placeholder de fecha (dt)
      if (!is.null(fecha_portada) && nzchar(fecha_portada)) {
        doc <- tryCatch(
          officer::ph_with(
            doc,
            fecha_portada,
            location = officer::ph_location_type(type = "dt")
          ),
          error = function(e) {
            if (mensajes_progreso) {
              message("No se pudo escribir la fecha en la contraportada: ", e$message)
            }
            doc
          }
        )
      }
    }

    # 6.6. Guardar
    print(doc, target = path_ppt)
    if (mensajes_progreso) {
      message("PPT generado en: ", normalizePath(path_ppt, winslash = "/"))
    }
  }

  invisible(list(
    plots          = plots_list,
    log_decisiones = log_decisiones
  ))
}
