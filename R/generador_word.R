# =============================================================================
# reporte_word()
# -----------------------------------------------------------------------------
# Generar Word con gráficos a partir de tablas de frecuencias (SO / SM)
# usando los graficadores ya existentes del paquete.
# =============================================================================

#' Generar un reporte Word a partir de frecuencias (SO/SM) con graficadores propios
#'
#' `reporte_word()` toma una base de reporte (idealmente la misma que se usa en
#' `reporte_frecuencias()`), el instrumento y una definición de secciones, y
#' genera:
#' \itemize{
#'   \item Una lista de gráficos \code{ggplot} (uno por variable válida).
#'   \item Opcionalmente, un archivo Word donde cada gráfico ocupa un bloque:
#'         título externo + gráfico centrado + pie gris de fuente.
#' }
#'
#' El tipo de gráfico para cada variable se decide igual que en `reporte_ppt()`:
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
#' Para las variables dicotómicas, no se asume ninguna pareja de categorías por
#' defecto: se requiere definirla en \code{dico_labels_por_var} o
#' \code{dico_labels_por_listname}. Si no se encuentra pareja válida, se usa
#' barras agrupadas por defecto.
#'
#' En el documento Word:
#' \itemize{
#'   \item El título del gráfico se escribe \emph{fuera} del gráfico, como
#'         \code{"Gráfico X: "} (en negrita) seguido del label de la variable
#'         (en cursiva), todo en color azul Pulso.
#'   \item El gráfico se inserta centrado, sin título interno.
#'   \item Debajo del gráfico se escribe un pie de fuente en gris (argumento
#'         \code{fuente}).
#'   \item Dentro del gráfico, en el caption, se escribe \code{"N = ..."} a la
#'         derecha, usando los argumentos \code{nota_pie} / \code{nota_pie_derecha}
#'         de los graficadores.
#' }
#'
#' @inheritParams reporte_ppt
#' @param path_word Ruta del archivo DOCX a generar cuando \code{solo_lista = FALSE}.
#' @param fuente Texto de fuente que se mostrará en gris debajo de cada gráfico
#'   en el documento Word (por ejemplo, `"Fuente: Pulso PUCP 2025"`).
#'
#' @return Una lista con:
#' \describe{
#'   \item{plots}{Lista de objetos \code{ggplot} generados (uno por variable).}
#'   \item{log_decisiones}{Tibble con información sobre cada variable procesada.}
#' }
#'
#' @export
reporte_word <- function(
    data,
    instrumento      = NULL,
    secciones        = NULL,
    path_word        = "reporte_word.docx",
    fuente           = NULL,
    sm_vars_force    = NULL,
    mostrar_todo     = FALSE,
    solo_lista       = FALSE,
    incluir_titulo_var = TRUE,
    mensajes_progreso  = TRUE,
    template_docx = NULL,

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
    estilos_dico             = list()
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
  # 3. Helpers internos (mismos que reporte_ppt, con pequeño ajuste en N)
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
  # 4. Recorrido por secciones y variables (igual que en PPT)
  # ---------------------------------------------------------------------------
  plots_list       <- list()
  titulos_list     <- list()
  log_list         <- list()
  seccion_por_plot <- character(0)
  N_texto_list     <- list()

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

      # --- N de la pregunta (para N = ... dentro del gráfico) ---
      if ("N" %in% names(tab_freq)) {
        n_var <- unique(tab_freq$N)[1]
      } else {
        n_var <- sum(tab_freq$n, na.rm = TRUE)
      }

      N_texto <- NULL
      if (is.finite(n_var) && n_var >= 0) {
        N_texto <- sprintf("N = %s", format(n_var, big.mark = ",", scientific = FALSE))
      }

      var_label    <- .titulo_var_safe(v)
      titulo_plot  <- NULL   # siempre sin título interno
      titulo_slide <- if (incluir_titulo_var) var_label else v

      # pie interno del gráfico: queremos que sea el N
      nota_pie_plot      <- NULL
      nota_pie_derecha   <- N_texto
      pos_nota_pie_plot  <- "derecha"

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
              data               = tab_agr,
              var_categoria      = "categoria",
              var_n              = "n_base",
              cols_porcentaje    = cols_porcentaje,
              etiquetas_series   = etiquetas_series,
              escala_valor       = "proporcion_1",
              colores_series     = colores_series,
              mostrar_valores    = TRUE,
              titulo             = titulo_plot,
              subtitulo          = NULL,
              nota_pie           = nota_pie_plot,
              nota_pie_derecha   = nota_pie_derecha,
              pos_nota_pie       = pos_nota_pie_plot,
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
          preset_extra   <- NULL

          if (!is.na(list_name_v) &&
              !is.null(colores_apiladas_por_listname[[list_name_v]])) {

            obj_col <- colores_apiladas_por_listname[[list_name_v]]

            if (is.list(obj_col)) {
              if (!is.null(obj_col$colores)) {
                colores_grupos <- obj_col$colores
              }
              if (!is.null(obj_col$preset_barra_extra)) {
                preset_extra <- obj_col$preset_barra_extra
              }
            } else {
              colores_grupos <- obj_col
            }
          }

          # inversión de segmentos / leyenda según estilos_barras_apiladas
          ln_inv_seg <- estilos_barras_apiladas$listnames_invertir_segmentos
          ln_inv_seg <- if (is.null(ln_inv_seg)) character(0) else ln_inv_seg

          ln_inv_ley <- estilos_barras_apiladas$listnames_invertir_leyenda
          ln_inv_ley <- if (is.null(ln_inv_ley)) character(0) else ln_inv_ley

          vars_inv_seg <- estilos_barras_apiladas$vars_invertir_segmentos
          vars_inv_seg <- if (is.null(vars_inv_seg)) character(0) else vars_inv_seg

          vars_inv_ley <- estilos_barras_apiladas$vars_invertir_leyenda
          vars_inv_ley <- if (is.null(vars_inv_ley)) character(0) else vars_inv_ley

          invertir_segmentos_var <- (
            v %in% vars_inv_seg ||
              (!is.na(list_name_v) && list_name_v %in% ln_inv_seg)
          )

          invertir_leyenda_var <- (
            v %in% vars_inv_ley ||
              (!is.na(list_name_v) && list_name_v %in% ln_inv_ley)
          )

          estilos_apiladas_clean <- estilos_barras_apiladas
          estilos_apiladas_clean$listnames_invertir_segmentos <- NULL
          estilos_apiladas_clean$listnames_invertir_leyenda   <- NULL
          estilos_apiladas_clean$vars_invertir_segmentos      <- NULL
          estilos_apiladas_clean$vars_invertir_leyenda        <- NULL

          args_apiladas <- c(
            list(
              data                = tab_apil$data,
              var_categoria       = "categoria",
              var_n               = "n_base",
              cols_porcentaje     = tab_apil$cols_porcentaje,
              etiquetas_grupos    = tab_apil$etiquetas_grupos,
              escala_valor        = "proporcion_1",
              colores_grupos      = colores_grupos,
              mostrar_valores     = TRUE,
              titulo              = titulo_plot,
              subtitulo           = NULL,
              nota_pie            = nota_pie_plot,
              nota_pie_derecha    = nota_pie_derecha,
              pos_nota_pie        = pos_nota_pie_plot,
              mostrar_barra_extra = if (!is.null(preset_extra)) TRUE else (barra_extra == "total_n"),
              barra_extra_preset  = preset_extra,
              prefijo_barra_extra = if (barra_extra == "total_n") "N = " else "N = ",
              titulo_barra_extra  = if (!is.null(preset_extra)) NULL else if (barra_extra == "total_n") "Total" else NULL,
              invertir_segmentos  = invertir_segmentos_var,
              invertir_leyenda    = invertir_leyenda_var,
              exportar            = "rplot"
            ),
            estilos_apiladas_clean
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
                data               = tab_agr,
                var_categoria      = "categoria",
                var_n              = "n_base",
                cols_porcentaje    = cols_porcentaje,
                etiquetas_series   = etiquetas_series,
                escala_valor       = "proporcion_1",
                colores_series     = colores_series,
                mostrar_valores    = TRUE,
                titulo             = titulo_plot,
                subtitulo          = NULL,
                nota_pie           = nota_pie_plot,
                nota_pie_derecha   = nota_pie_derecha,
                pos_nota_pie       = pos_nota_pie_plot,
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
                nota_pie          = N_texto,  # N = ... como caption del pie
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

      idx <- length(plots_list) + 1L
      plots_list[[idx]]       <- p
      titulos_list[[idx]]     <- titulo_slide
      N_texto_list[[idx]]     <- N_texto
      seccion_por_plot[idx]   <- sec
    }
  }

  log_decisiones <- dplyr::bind_rows(log_list)

  # ---------------------------------------------------------------------------
  # 5. Construir DOCX
  # ---------------------------------------------------------------------------
  if (!solo_lista) {
    if (!requireNamespace("officer", quietly = TRUE)) {
      stop(
        "Para exportar a Word se requiere el paquete 'officer'.",
        call. = FALSE
      )
    }

    pulso_azul <- "#004B8D"

    # -------------------------------------------------------------------------
    # 5.1. Leer plantilla Word (análogo al PPT, con logos en encabezado)
    # -------------------------------------------------------------------------
    if (is.null(template_docx)) {
      # Buscar plantilla interna del paquete
      template_interno <- system.file(
        "plantillas/plantilla_pulso.docx",
        package = "prosecnur"
      )

      if (nzchar(template_interno) && file.exists(template_interno)) {
        if (mensajes_progreso) {
          message("Usando plantilla Word interna: ", template_interno)
        }
        doc <- officer::read_docx(path = template_interno)
      } else {
        if (mensajes_progreso) {
          message(
            "No se encontró 'plantillas/plantilla_pulso.docx' dentro del paquete. ",
            "Se usará la plantilla por defecto de Word."
          )
        }
        doc <- officer::read_docx()
      }

    } else {
      # Plantilla pasada explícitamente por el usuario
      if (!file.exists(template_docx)) {
        stop(
          "No se encontró el archivo de plantilla especificado en `template_docx`: ",
          template_docx,
          call. = FALSE
        )
      }
      if (mensajes_progreso) {
        message("Usando plantilla Word definida por el usuario: ", template_docx)
      }
      doc <- officer::read_docx(path = template_docx)
    }

    # -------------------------------------------------------------------------
    # 5.2. Primera página: índice de gráficas por sección
    # -------------------------------------------------------------------------
    if (length(plots_list)) {

      # Título del índice
      indice_titulo <- officer::fpar(
        officer::ftext(
          "ÍNDICE DE GRÁFICAS",
          prop = officer::fp_text(
            bold        = TRUE,
            italic      = FALSE,
            color       = pulso_azul,
            font.size   = 12,
            font.family = "Arial"
          )
        )
      )
      doc <- officer::body_add_fpar(doc, indice_titulo, style = "Normal")
      doc <- officer::body_add_par(doc, "", style = "Normal")

      # Agrupar por sección (si existe `seccion_por_plot`)
      if (exists("seccion_por_plot") &&
          length(seccion_por_plot) == length(plots_list)) {

        secciones_unicas <- unique(seccion_por_plot)

        for (sec in secciones_unicas) {
          if (is.na(sec) || !nzchar(sec)) next

          # Nombre de sección
          sec_fpar <- officer::fpar(
            officer::ftext(
              sec,
              prop = officer::fp_text(
                bold        = TRUE,
                italic      = FALSE,
                color       = "black",
                font.size   = 11,
                font.family = "Arial"
              )
            )
          )
          doc <- officer::body_add_fpar(doc, sec_fpar, style = "Normal")

          idx_sec <- which(seccion_por_plot == sec)

          for (i in idx_sec) {
            titulo_i <- titulos_list[[i]] %||% ""

            # Entrada índice: "Gráfica Nro. i. " (azul) + título (negro, cursiva)
            entrada_indice <- officer::fpar(
              officer::ftext(
                sprintf("Gráfica Nro. %d. ", i),
                prop = officer::fp_text(
                  bold        = TRUE,
                  italic      = FALSE,
                  color       = pulso_azul,
                  font.size   = 10,
                  font.family = "Arial"
                )
              ),
              officer::ftext(
                titulo_i,
                prop = officer::fp_text(
                  bold        = FALSE,
                  italic      = TRUE,
                  color       = "black",
                  font.size   = 10,
                  font.family = "Arial"
                )
              )
              # Aquí podrías luego agregar tabs y número de página manualmente
            )

            doc <- officer::body_add_fpar(doc, entrada_indice, style = "Normal")
          }

          # Línea en blanco entre secciones
          doc <- officer::body_add_par(doc, "", style = "Normal")
        }

      } else {
        # Fallback: índice lineal sin secciones
        for (i in seq_along(plots_list)) {
          titulo_i <- titulos_list[[i]] %||% ""

          entrada_indice <- officer::fpar(
            officer::ftext(
              sprintf("Gráfica Nro. %d. ", i),
              prop = officer::fp_text(
                bold        = TRUE,
                italic      = FALSE,
                color       = pulso_azul,
                font.size   = 10,
                font.family = "Arial"
              )
            ),
            officer::ftext(
              titulo_i,
              prop = officer::fp_text(
                bold        = FALSE,
                italic      = TRUE,
                color       = "black",
                font.size   = 10,
                font.family = "Arial"
              )
            )
          )

          doc <- officer::body_add_fpar(doc, entrada_indice, style = "Normal")
        }
      }

      # Salto de página después del índice
      doc <- officer::body_add_break(doc)

      # -----------------------------------------------------------------------
      # 5.3. Gráficas: título externo + gráfico + pie de página
      # -----------------------------------------------------------------------
      for (i in seq_along(plots_list)) {

        p        <- plots_list[[i]]
        titulo_i <- titulos_list[[i]] %||% ""
        N_i      <- N_texto_list[[i]] %||% NULL  # reservado si luego lo usas

        # Título externo:
        # "Gráfica Nro. i: " (negrita, azul) + título (negro, cursiva)
        titulo_word <- officer::fpar(
          officer::ftext(
            sprintf("Gráfica Nro. %d: ", i),
            prop = officer::fp_text(
              bold        = TRUE,
              italic      = FALSE,
              color       = pulso_azul,
              font.size   = 10,
              font.family = "Arial"
            )
          ),
          officer::ftext(
            titulo_i,
            prop = officer::fp_text(
              bold        = FALSE,
              italic      = TRUE,
              color       = "black",
              font.size   = 10,
              font.family = "Arial"
            )
          )
        )

        doc <- officer::body_add_fpar(doc, titulo_word, style = "Normal")

        # Gráfico (sin estilo personalizado; usa Normal de la plantilla)
        doc <- officer::body_add_gg(
          doc,
          value  = p,
          width  = 6.1,
          height = 3.5
        )

        # Pie gris debajo del gráfico (fuente) a la izquierda
        if (!is.null(fuente) && nzchar(fuente)) {
          pie_word <- officer::fpar(
            officer::ftext(
              fuente,
              prop = officer::fp_text(
                color       = "#7F7F7F",
                font.size   = 9,
                font.family = "Arial"
              )
            )
          )
          doc <- officer::body_add_fpar(doc, pie_word, style = "Normal")
        }

        # Espacio entre gráficas
        doc <- officer::body_add_par(doc, "", style = "Normal")
      }
    }

    print(doc, target = path_word)
    if (mensajes_progreso) {
      message("Word generado en: ", normalizePath(path_word, winslash = "/"))
    }
  }

  invisible(list(
    plots          = plots_list,
    log_decisiones = log_decisiones
  ))
}
