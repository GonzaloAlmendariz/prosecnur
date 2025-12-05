# =============================================================================
# reporte_word_cruces()
# -----------------------------------------------------------------------------
# Generar DOCX con gráficos a partir de cruces (var x estrato) por secciones,
# usando los graficadores propios del paquete (barras agrupadas, apiladas, dico).
# Lógica de cruces análoga a `reporte_ppt_cruces()`, pero exportando a Word.
# =============================================================================

#' Generar un reporte Word a partir de cruces (SO/SM vs estrato) con graficadores propios
#'
#' `reporte_word_cruces()` toma una base de datos, el instrumento, una definición
#' de secciones y una variable de cruce (\code{var_cruce}) y genera:
#' \itemize{
#'   \item Una lista de gráficos \code{ggplot} (uno por variable de las secciones).
#'   \item Opcionalmente, un archivo DOCX donde cada gráfico ocupa un bloque en
#'         páginas sucesivas, precedidos por una página de índice.
#' }
#'
#' La lógica de construcción de cruces es análoga a \code{reporte_ppt_cruces()};
#' véase esa función para más detalles conceptuales.
#'
#' @inheritParams reporte_ppt_cruces
#' @param path_word Ruta del archivo DOCX a generar cuando \code{solo_lista = FALSE}.
#' @param template_docx Ruta a una plantilla DOCX. Si es \code{NULL}, se intentará
#'   usar una plantilla interna del paquete llamada
#'   \code{"plantillas/plantilla_pulso.docx"}; si tampoco existe, se usará la
#'   plantilla por defecto de Word a través de \code{officer::read_docx()}.
#'
#' @return Una lista con dos elementos:
#' \describe{
#'   \item{plots}{Lista de objetos \code{ggplot} generados (uno por variable).}
#'   \item{log_decisiones}{Tibble con información sobre cada variable procesada:
#'         sección, nombre de variable, estrato, tipo de pregunta, \code{list_name},
#'         origen de la decisión (\code{override}) y tipo de gráfico final.}
#' }
#'
#' @export
reporte_word_cruces <- function(
    data,
    instrumento      = NULL,
    secciones,
    var_cruce,
    path_word        = "reporte_word_cruces.docx",
    fuente           = NULL,
    sm_vars_force    = NULL,
    weight_col       = "peso",
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
    # SO → apiladas por defecto, SM → agrupadas
    default_so = c("barras_apiladas", "barras_agrupadas"),
    default_sm = c("barras_agrupadas", "barras_apiladas"),

    # Barra extra en barras apiladas/agrupadas
    barra_extra = c("ninguna", "total_n"),

    # Estilos por tipo de gráfico
    estilos_barras_agrupadas = list(),
    estilos_barras_apiladas  = list(),
    estilos_dico             = list(),

    # Colores por estrato en barras agrupadas
    colores_estratos = NULL,

    # Opciones de la variable a excluir
    opciones_excluir = NULL,

    # Plantilla Word
    template_docx = NULL,

    # Resumen de N (texto tipo "N=... | Ratio ...") bajo cada gráfico
    mostrar_resumen_n = TRUE,

    # Override de labels de variables
    labels_override = NULL
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

  if (missing(secciones) || !length(secciones)) {
    stop("Debe especificar `secciones` (lista nombrada de variables).", call. = FALSE)
  }

  if (missing(var_cruce)) {
    stop("Debe especificar `var_cruce`.", call. = FALSE)
  }

  if (!(var_cruce %in% names(data))) {
    stop("La variable `var_cruce` no existe en `data`.", call. = FALSE)
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
  # 2. Filtrar SECCIONES a variables presentes (como en exportar_cruces_multi)
  # ---------------------------------------------------------------------------
  SECCIONES <- lapply(secciones, function(v) {
    v[vapply(v, function(x) .has_var_or_dummies(data, x), logical(1))]
  })
  SECCIONES <- SECCIONES[vapply(SECCIONES, length, integer(1)) > 0L]

  if (!length(SECCIONES)) {
    stop("Después de filtrar por presencia en `data` (variable o dummies), ",
         "ninguna sección tiene variables válidas.", call. = FALSE)
  }

  # ---------------------------------------------------------------------------
  # 3. Helpers internos (títulos y cruces)
  # ---------------------------------------------------------------------------

  .titulo_var_safe <- function(var_name) {
    label_variable(
      var             = var_name,
      dic_vars        = dic_vars,
      labels_override = labels_override,
      data            = data
    )
  }

  .list_name_safe <- function(var_name) {
    ln <- get_list_name(var_name, survey = survey)
    if (!is.na(ln)) return(ln)
    if ("list_norm" %in% names(survey)) {
      ln2 <- unique(na.omit(as.character(survey$list_norm[survey$name == var_name])))
      if (length(ln2)) return(ln2[1])
    }
    NA_character_
  }

  # Helper: dado un vector de frecuencias n, devuelve porcentajes ENTEROS
  # que suman exactamente 100 (pensado para barras apiladas / dico).
  .pct_enteros_100 <- function(n) {
    n <- as.numeric(n)
    if (!length(n) || all(is.na(n))) {
      return(numeric(0))
    }

    n[is.na(n)] <- 0
    total <- sum(n)

    if (!is.finite(total) || total <= 0) {
      return(rep(0L, length(n)))
    }

    # Porcentajes crudos en 0–100
    raw_pct   <- n / total * 100
    floor_pct <- floor(raw_pct)

    # Cuánto falta o sobra para llegar a 100
    resid <- as.integer(round(100 - sum(floor_pct)))
    frac  <- raw_pct - floor_pct

    if (resid > 0) {
      # Asignar +1 a los mayores restos decimales
      ord <- order(frac, decreasing = TRUE, na.last = TRUE)
      idx <- head(ord, resid)
      floor_pct[idx] <- floor_pct[idx] + 1L
    } else if (resid < 0) {
      # Quitar 1 a los menores restos decimales
      resid_neg <- abs(resid)
      ord <- order(frac, decreasing = FALSE, na.last = TRUE)
      idx <- head(ord, resid_neg)
      floor_pct[idx] <- pmax(0L, floor_pct[idx] - 1L)
    }

    floor_pct
  }

  .cruce_counts <- function(var_name, estrato_name) {

    tp <- tipo_pregunta(
      var           = var_name,
      survey        = survey,
      sm_vars_force = sm_vars_force,
      data          = data
    )

    cats_var <- get_categorias(
      var              = var_name,
      data             = data,
      survey           = survey,
      orders_list      = orders_list,
      opciones_excluir = opciones_excluir
    )
    categorias <- cats_var$labels
    codes_row  <- cats_var$codes

    if (!length(categorias)) return(NULL)

    cats_est <- get_categorias(
      var              = estrato_name,
      data             = data,
      survey           = survey,
      orders_list      = orders_list,
      opciones_excluir = NULL
    )
    estr_codes  <- cats_est$codes
    estr_labels <- cats_est$labels

    if (!length(estr_labels)) {
      v_estr <- as.character(data[[estrato_name]])
      estr_labels <- sort(unique(na.omit(v_estr)))
      estr_codes  <- estr_labels
    }

    v_estr <- as.character(data[[estrato_name]])

    usa_codes  <- any(v_estr %in% estr_codes)
    usa_labels <- any(v_estr %in% estr_labels)
    keys_vec   <- if (usa_codes || !usa_labels) estr_codes else estr_labels

    n_cat <- length(categorias)
    n_est <- length(keys_vec)

    n_mat     <- matrix(0, nrow = n_cat, ncol = n_est)
    N_estrato <- numeric(n_est)

    for (j in seq_along(keys_vec)) {
      key_j  <- keys_vec[j]
      mask_j <- !is.na(v_estr) & v_estr == key_j

      N_j <- denominador_validos(
        data       = data,
        var        = var_name,
        codes      = codes_row,
        tp         = tp,
        mask       = mask_j,
        weight_col = weight_col
      )

      N_estrato[j] <- N_j

      if (N_j == 0) next

      n_vec <- contar_por_opcion(
        data       = data,
        var        = var_name,
        codes      = codes_row,
        tp         = tp,
        mask       = mask_j,
        weight_col = weight_col
      )

      n_mat[, j] <- as.numeric(n_vec)
    }

    keep_est <- which(N_estrato > 0)
    if (!length(keep_est)) return(NULL)

    n_mat       <- n_mat[, keep_est, drop = FALSE]
    N_estrato   <- N_estrato[keep_est]
    estr_labels <- estr_labels[keep_est]
    estr_codes  <- estr_codes[keep_est]

    suma_cat <- rowSums(n_mat, na.rm = TRUE)
    keep_cat <- which(suma_cat > 0)
    if (!length(keep_cat)) return(NULL)

    n_mat      <- n_mat[keep_cat, , drop = FALSE]
    categorias <- categorias[keep_cat]
    codes_row  <- codes_row[keep_cat]

    list(
      tp           = tp,
      categorias   = categorias,
      codes_row    = codes_row,
      estr_labels  = estr_labels,
      estr_codes   = estr_codes,
      n_mat        = n_mat,
      N_estrato    = N_estrato
    )
  }

  .build_tab_barras_apiladas_cruce <- function(crc, var_label, list_name_v) {

    n_cat <- length(crc$categorias)
    n_est <- length(crc$estr_labels)
    if (n_cat == 0L || n_est == 0L) return(NULL)

    # Matriz de porcentajes ENTEROS por estrato (columnas) que suman 100
    pct_int_mat <- matrix(NA_integer_, nrow = n_cat, ncol = n_est)
    for (j in seq_len(n_est)) {
      Nj <- crc$N_estrato[j]
      if (Nj > 0) {
        pct_int_mat[, j] <- .pct_enteros_100(crc$n_mat[, j])
      } else {
        pct_int_mat[, j] <- 0L
      }
    }

    df <- tibble::tibble(
      categoria = crc$estr_labels,
      n_base    = as.numeric(crc$N_estrato)
    )

    cols_pct <- paste0("pct_", seq_len(n_cat))
    for (k in seq_len(n_cat)) {
      # Guardar como proporciones 0–1 (para escala_valor = "proporcion_1")
      df[[cols_pct[k]]] <- pct_int_mat[k, ] / 100
    }

    etiquetas_grupos <- stats::setNames(
      as.character(crc$categorias),
      cols_pct
    )

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

    list(
      data             = df,
      cols_porcentaje  = cols_pct,
      etiquetas_grupos = etiquetas_grupos,
      colores_grupos   = colores_grupos,
      preset_extra     = preset_extra,
      info_dim         = list(
        n_cat      = n_cat,
        n_estratos = n_est,
        N_estrato  = crc$N_estrato,
        N_total    = sum(crc$N_estrato, na.rm = TRUE)
      )
    )
  }

  # ---------------------------------------------------------------------------
  # barras agrupadas con eje = estrato y series = categorías (SO general)
  # ---------------------------------------------------------------------------
  .build_tab_barras_agrupadas_cruce <- function(crc, var_label) {

    n_cat <- length(crc$categorias)   # categorías de v (series)
    n_est <- length(crc$estr_labels)  # estratos (filas)
    if (n_cat == 0L || n_est == 0L) return(NULL)

    # matriz de porcentajes categoría (filas) x estrato (columnas)
    pct_mat <- matrix(NA_real_, nrow = n_cat, ncol = n_est)
    for (j in seq_len(n_est)) {
      Nj <- crc$N_estrato[j]
      if (Nj > 0) pct_mat[, j] <- crc$n_mat[, j] / Nj
    }

    # 1) Eliminar categorías cuya suma es 0 (todas 0/NA)
    keep_rows <- rowSums(pct_mat, na.rm = TRUE) > 0
    # 1-bis) Eliminar categorías con label NA
    keep_rows <- keep_rows & !is.na(crc$categorias)

    if (!any(keep_rows)) return(NULL)

    pct_mat         <- pct_mat[keep_rows, , drop = FALSE]
    categorias_keep <- crc$categorias[keep_rows]
    n_cat           <- length(categorias_keep)

    # 2) Eliminar estratos cuya suma es 0 (todas 0/NA)
    keep_cols <- colSums(pct_mat, na.rm = TRUE) > 0
    if (!any(keep_cols)) return(NULL)

    pct_mat          <- pct_mat[, keep_cols, drop = FALSE]
    estr_labels_keep <- crc$estr_labels[keep_cols]
    N_estrato_keep   <- crc$N_estrato[keep_cols]
    n_est            <- length(estr_labels_keep)

    # 3) Construir data.frame: filas = estratos, series = categorías
    df <- tibble::tibble(
      categoria = estr_labels_keep,
      n_base    = as.numeric(N_estrato_keep)
    )

    cols_pct <- paste0("pct_", seq_len(n_cat))
    for (i in seq_len(n_cat)) {
      v <- as.numeric(pct_mat[i, ])
      # 4) 0% → NA para no dibujar barra ni generar huecos visuales
      v[!is.na(v) & abs(v) < 1e-12] <- NA_real_
      df[[cols_pct[i]]] <- v
    }

    # series = categorías de la variable
    etiquetas_series <- stats::setNames(
      as.character(categorias_keep),
      cols_pct
    )

    list(
      data             = df,
      cols_porcentaje  = cols_pct,
      etiquetas_series = etiquetas_series,
      info_dim         = list(
        n_cat      = n_cat,            # Nº de series (categorías de v)
        n_estratos = n_est,            # Nº de filas (estratos)
        N_estrato  = N_estrato_keep,
        N_total    = sum(N_estrato_keep, na.rm = TRUE)
      )
    )
  }

  # ---------------------------------------------------------------------------
  # Helper: barras agrupadas tipo SM para UN estrato (igual que reporte_ppt)
  # ---------------------------------------------------------------------------
  .build_tab_barras_agrupadas_sm_estrato <- function(crc, j) {

    N_j <- crc$N_estrato[j]
    if (!is.finite(N_j) || N_j <= 0) return(NULL)

    n_vec <- crc$n_mat[, j]
    cats  <- crc$categorias

    # Eliminar categorías sin frecuencia
    keep <- !is.na(n_vec) & n_vec > 0
    if (!any(keep)) return(NULL)

    n_vec <- n_vec[keep]
    cats  <- cats[keep]

    # Porcentajes 0–100 para ese estrato
    pct_0_100 <- (n_vec / N_j) * 100
    pct_int   <- round(pct_0_100)     # enteros
    pct_prop  <- pct_int / 100        # escala 0–1 para el graficador

    tibble::tibble(
      categoria = cats,
      n_base    = N_j,
      pct       = pct_prop
    )
  }

  .build_tab_dico_cruce <- function(crc, var_name, var_label, labels_dico) {

    if (length(labels_dico) < 2L) return(NULL)

    pos_lab <- labels_dico[1]
    neg_lab <- labels_dico[2]

    idx_pos <- which(crc$categorias == pos_lab)
    idx_neg <- which(crc$categorias == neg_lab)
    if (!length(idx_pos) || !length(idx_neg)) {
      warning(
        "En la variable '", var_name, "' no se encontraron ambas categorías ",
        "indicadas en `labels_dico` dentro del cruce. Se omitirá tratamiento ",
        "dicotómico en cruces.",
        call. = FALSE
      )
      return(NULL)
    }

    idx_pos <- idx_pos[1]
    idx_neg <- idx_neg[1]

    n_pos <- crc$n_mat[idx_pos, ]
    n_neg <- crc$n_mat[idx_neg, ]
    denom <- n_pos + n_neg

    # Porcentaje enteros que suman 100
    pct_si_int <- rep(NA_real_, length(denom))
    for (j in seq_along(denom)) {
      if (is.finite(denom[j]) && denom[j] > 0) {
        n_pair   <- c(n_pos[j], n_neg[j])
        pct_pair <- .pct_enteros_100(n_pair)   # c(%Sí, %No) enteros que suman 100
        pct_si_int[j] <- pct_pair[1]           # Sí en 0–100
      } else {
        pct_si_int[j] <- NA_real_
      }
    }

    tibble::tibble(
      indicador  = as.character(crc$estr_labels),
      pct_si     = pct_si_int,      # escala 0–100
      n_total    = as.numeric(denom),
      pct_si_int = pct_si_int       # se deja por si quieres inspeccionarlo
    )
  }

  # ---------------------------------------------------------------------------
  # 4. Loop por secciones y variables
  # ---------------------------------------------------------------------------
  plots_list       <- list()
  titulos_list     <- list()
  resumenN_list    <- list()
  log_list         <- list()
  seccion_por_plot <- character(0)

  total_casos <- nrow(data)
  idx_plot    <- 0L

  for (sec in names(SECCIONES)) {
    vars_sec <- SECCIONES[[sec]]
    if (!length(vars_sec)) next

    if (mensajes_progreso) {
      message("Procesando sección: ", sec)
    }

    for (v in vars_sec) {

      tipo_v <- tipo_pregunta(
        var           = v,
        survey        = survey,
        sm_vars_force = sm_vars_force,
        data          = data
      )
      if (tipo_v == "so_or_open") tipo_v <- "so"

      list_name_v <- .list_name_safe(v)

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

      crc <- .cruce_counts(v, var_cruce)
      if (is.null(crc)) {
        log_list[[length(log_list) + 1]] <- tibble::tibble(
          seccion      = sec,
          var          = v,
          estrato      = var_cruce,
          tipo_var     = tipo_v,
          list_name    = list_name_v,
          override     = override,
          tipo_grafico = NA_character_
        )
        next
      }

      N_total_cruce <- sum(crc$N_estrato, na.rm = TRUE)
      if (is.finite(N_total_cruce) && N_total_cruce >= 0 && total_casos > 0) {
        ratio <- N_total_cruce / total_casos * 100
        resumen_n_txt <- sprintf(
          "N = %s | Ratio de respuestas: %.1f%%",
          format(N_total_cruce, big.mark = ",", scientific = FALSE),
          ratio
        )
      } else if (is.finite(N_total_cruce)) {
        resumen_n_txt <- sprintf(
          "N = %s",
          format(N_total_cruce, big.mark = ",", scientific = FALSE)
        )
      } else {
        resumen_n_txt <- NULL
      }

      var_label     <- .titulo_var_safe(v)
      titulo_plot   <- NULL
      titulo_slide  <- if (incluir_titulo_var) var_label else NULL
      nota_pie_plot <- NULL

      if (mensajes_progreso) {
        message("   - ", v, " x ", var_cruce, " → ", tipo_grafico,
                " (list_name = ", list_name_v, ") [", override, "]")
      }

      p <- NULL
      tipo_grafico_final <- tipo_grafico

      if (tipo_grafico %in% c("barras_agrupadas", "barras_apiladas", "dico")) {

        if (tipo_grafico == "barras_agrupadas") {

          tab_agr <- .build_tab_barras_agrupadas_cruce(crc, var_label)
          if (!is.null(tab_agr)) {

            cols_porcentaje  <- tab_agr$cols_porcentaje
            etiquetas_series <- tab_agr$etiquetas_series

            # -----------------------------------------------------------------
            # CASO ESPECIAL: SM + barras agrupadas → UNA PÁGINA POR ESTRATO
            # (mismo comportamiento conceptual que en reporte_ppt_cruces)
            # -----------------------------------------------------------------
            if (tipo_v == "sm") {

              n_estratos <- length(crc$estr_labels)
              N_estrato  <- crc$N_estrato

              for (j in seq_len(n_estratos)) {

                tab_j <- .build_tab_barras_agrupadas_sm_estrato(crc, j)
                if (is.null(tab_j) || !nrow(tab_j)) next

                cols_porcentaje_j  <- "pct"
                etiquetas_series_j <- c(pct = "Porcentaje")

                args_extra_j <- list()

                # Color por defecto: azul Pulso sólido
                if (is.null(estilos_barras_agrupadas$colores_series)) {

                  pulso_azul <- "#1B679D"

                  colores_series <- stats::setNames(
                    rep(pulso_azul, length(etiquetas_series_j)),
                    etiquetas_series_j
                  )

                  args_extra_j$colores_series <- colores_series
                }

                # Tamaño automático según Nº de categorías en ese estrato
                if (is.null(estilos_barras_agrupadas$size_texto_barras)) {
                  n_cat_plot <- nrow(tab_j)
                  n_series   <- 1L

                  size_auto <- if (n_series <= 3) {
                    if (n_cat_plot <= 4) 4.5 else if (n_cat_plot <= 8) 4.0 else 3.5
                  } else if (n_series <= 5) {
                    if (n_cat_plot <= 4) 4.0 else 3.5
                  } else {
                    3.0
                  }
                  args_extra_j$size_texto_barras <- size_auto
                }

                args_barras_j <- c(
                  list(
                    data             = tab_j,
                    var_categoria    = "categoria",
                    var_n            = "n_base",
                    cols_porcentaje  = cols_porcentaje_j,
                    etiquetas_series = etiquetas_series_j,
                    escala_valor     = "proporcion_1",
                    mostrar_valores  = TRUE,
                    titulo           = NULL,
                    subtitulo        = NULL,
                    nota_pie         = nota_pie_plot,
                    mostrar_barra_extra = FALSE,
                    exportar            = "rplot"
                  ),
                  args_extra_j,
                  estilos_barras_agrupadas
                )

                p_j <- do.call(graficar_barras_agrupadas, args_barras_j)

                # Para ajustar altura en Word según nº de categorías
                attr(p_j, "n_series_cruce") <- nrow(tab_j)

                # Registrar en el log
                log_list[[length(log_list) + 1]] <- tibble::tibble(
                  seccion      = sec,
                  var          = v,
                  estrato      = paste0(var_cruce, " = ", crc$estr_labels[j]),
                  tipo_var     = tipo_v,
                  list_name    = list_name_v,
                  override     = override,
                  tipo_grafico = "barras_agrupadas"
                )

                idx_plot <- idx_plot + 1L
                plots_list[[idx_plot]] <- p_j

                # Título en Word: "Pregunta - Estrato"
                titulos_list[[idx_plot]] <- if (incluir_titulo_var) {
                  paste0(var_label, " - ", crc$estr_labels[j])
                } else {
                  NULL
                }

                # Resumen N específico del estrato
                Nj <- N_estrato[j]
                if (is.finite(Nj) && Nj >= 0 && total_casos > 0) {
                  ratio_j <- Nj / total_casos * 100
                  resumenN_list[[idx_plot]] <- sprintf(
                    "N = %s | Ratio de respuestas: %.1f%%",
                    format(Nj, big.mark = ",", scientific = FALSE),
                    ratio_j
                  )
                } else if (is.finite(Nj)) {
                  resumenN_list[[idx_plot]] <- sprintf(
                    "N = %s",
                    format(Nj, big.mark = ",", scientific = FALSE)
                  )
                } else {
                  resumenN_list[[idx_plot]] <- NULL
                }

                seccion_por_plot[idx_plot] <- sec
              }

              # Ya se agregaron todos los plots de esta variable (uno por estrato)
              next
            }

            # -----------------------------------------------------------------
            # Resto de casos (SO u otros): un solo gráfico con todos los estratos
            # -----------------------------------------------------------------
            args_extra <- list()

            # COLORES por defecto (si el usuario no los definió):
            # una sola paleta azul para todas las series
            if (is.null(estilos_barras_agrupadas$colores_series)) {

              pulso_azul <- "#004B8D"

              colores_series <- stats::setNames(
                rep(pulso_azul, length(etiquetas_series)),
                etiquetas_series
              )

              args_extra$colores_series <- colores_series
            }

            # Tamaño automático de texto
            if (is.null(estilos_barras_agrupadas$size_texto_barras)) {
              n_series   <- tab_agr$info_dim$n_cat
              n_cat_plot <- tab_agr$info_dim$n_estratos

              size_auto <- if (n_series <= 3) {
                if (n_cat_plot <= 4) 4.5 else if (n_cat_plot <= 8) 4.0 else 3.5
              } else if (n_series <= 5) {
                if (n_cat_plot <= 4) 4.0 else 3.5
              } else {
                3.0
              }
              args_extra$size_texto_barras <- size_auto
            }

            args_barras <- c(
              list(
                data             = tab_agr$data,
                var_categoria    = "categoria",
                var_n            = "n_base",
                cols_porcentaje  = cols_porcentaje,
                etiquetas_series = etiquetas_series,
                escala_valor     = "proporcion_1",
                mostrar_valores  = TRUE,
                titulo           = NULL,
                subtitulo        = NULL,
                nota_pie         = nota_pie_plot,
                mostrar_barra_extra = FALSE,
                exportar            = "rplot"
              ),
              args_extra,
              estilos_barras_agrupadas
            )

            p <- do.call(graficar_barras_agrupadas, args_barras)
          }
        }

        if (tipo_grafico == "barras_apiladas") {

          tab_apil <- .build_tab_barras_apiladas_cruce(crc, var_label, list_name_v)
          if (!is.null(tab_apil)) {

            colores_grupos <- tab_apil$colores_grupos
            preset_extra   <- tab_apil$preset_extra

            args_extra <- list()
            if (!is.null(colores_grupos)) args_extra$colores_grupos <- colores_grupos

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
                mostrar_valores     = TRUE,
                titulo              = titulo_plot,
                subtitulo           = NULL,
                nota_pie            = nota_pie_plot,

                # ------------------------------
                # BARRA EXTRA: lógica elegante
                # ------------------------------
                mostrar_barra_extra = if (!is.null(preset_extra)) TRUE else (barra_extra == "total_n"),
                barra_extra_preset  = preset_extra,

                prefijo_barra_extra = if (!is.null(preset_extra)) {
                  ""
                } else if (barra_extra == "total_n") {
                  "N = "
                } else {
                  ""
                },

                # Sin título excepto cuando tú lo definas
                titulo_barra_extra = NULL,

                # Color:
                # - Si hay preset → dejar NULL (para que el Top2Box se pinte VERDE)
                # - Si NO hay preset → N= en azul (#092147)
                color_barra_extra = if (!is.null(preset_extra)) {
                  NULL
                } else {
                  "#092147"
                },

                # ------------------------------
                exportar           = "rplot",
                invertir_segmentos = invertir_segmentos_var,
                invertir_leyenda   = invertir_leyenda_var
              ),
              args_extra,
              estilos_apiladas_clean
            )

            p <- do.call(graficar_barras_apiladas, args_apiladas)
          }
        }

        if (tipo_grafico == "dico") {

          labels_dico <- NULL
          if (!is.null(dico_labels_por_var[[v]])) {
            labels_dico <- dico_labels_por_var[[v]]
          } else if (!is.na(list_name_v) &&
                     !is.null(dico_labels_por_listname[[list_name_v]])) {
            labels_dico <- dico_labels_por_listname[[list_name_v]]
          }

          if (is.null(labels_dico) || length(labels_dico) < 2L) {
            warning(
              "En la variable '", v, "' no se encontraron etiquetas dicotómicas ",
              "en `dico_labels_por_var` ni `dico_labels_por_listname`. ",
              "Se usa barras agrupadas para el cruce.",
              call. = FALSE
            )
            tipo_grafico_final <- "barras_agrupadas"

            tab_agr <- .build_tab_barras_agrupadas_cruce(crc, var_label)
            if (!is.null(tab_agr)) {

              cols_porcentaje  <- tab_agr$cols_porcentaje
              etiquetas_series <- tab_agr$etiquetas_series

              args_extra <- list()

              if (is.null(estilos_barras_agrupadas$colores_series)) {

                colores_series <- NULL

                # 1) Intentar colores por list_name (como en apiladas)
                if (!is.na(list_name_v) &&
                    !is.null(colores_apiladas_por_listname[[list_name_v]])) {

                  obj_col <- colores_apiladas_por_listname[[list_name_v]]

                  col_vec <- NULL
                  if (is.list(obj_col)) {
                    if (!is.null(obj_col$colores)) col_vec <- obj_col$colores
                  } else {
                    col_vec <- obj_col
                  }

                  if (!is.null(col_vec)) {
                    if (is.null(names(col_vec))) {
                      pal <- rep(
                        as.character(col_vec),
                        length.out = length(etiquetas_series)
                      )
                      colores_series <- stats::setNames(pal, etiquetas_series)
                    } else {
                      cs <- col_vec[etiquetas_series]
                      if (all(!is.na(cs))) {
                        colores_series <- stats::setNames(
                          as.character(cs),
                          etiquetas_series
                        )
                      }
                    }
                  }
                }

                # 2) Respaldo: colores_estratos si matchean nombres
                if (is.null(colores_series) &&
                    !is.null(colores_estratos) &&
                    all(etiquetas_series %in% names(colores_estratos))) {

                  cs <- colores_estratos[etiquetas_series]
                  colores_series <- stats::setNames(
                    as.character(cs),
                    etiquetas_series
                  )
                }

                # 3) Paleta por defecto multi-color
                if (is.null(colores_series)) {
                  pal <- c(
                    "#004B8D", "#F26C4F", "#8CC63E",
                    "#FFC20E", "#A54399", "#00A3E0", "#7F7F7F"
                  )
                  pal <- rep(pal, length.out = length(etiquetas_series))
                  colores_series <- stats::setNames(pal, etiquetas_series)
                }

                args_extra$colores_series <- colores_series
              }

              if (is.null(estilos_barras_agrupadas$size_texto_barras)) {
                n_series   <- tab_agr$info_dim$n_cat
                n_cat_plot <- tab_agr$info_dim$n_estratos

                size_auto <- if (n_series <= 3) {
                  if (n_cat_plot <= 4) 4.5 else if (n_cat_plot <= 8) 4.0 else 3.5
                } else if (n_series <= 5) {
                  if (n_cat_plot <= 4) 4.0 else 3.5
                } else {
                  3.0
                }
                args_extra$size_texto_barras <- size_auto
              }

              args_barras <- c(
                list(
                  data             = tab_agr$data,
                  var_categoria    = "categoria",
                  var_n            = "n_base",
                  cols_porcentaje  = cols_porcentaje,
                  etiquetas_series = etiquetas_series,
                  escala_valor     = "proporcion_1",
                  mostrar_valores  = TRUE,
                  titulo           = titulo_plot,
                  subtitulo        = NULL,
                  nota_pie         = nota_pie_plot,
                  mostrar_barra_extra = FALSE,
                  exportar            = "rplot"
                ),
                args_extra,
                estilos_barras_agrupadas
              )

              p <- do.call(graficar_barras_agrupadas, args_barras)
            }

          } else {

            tab_dico <- .build_tab_dico_cruce(crc, v, var_label, labels_dico)
            if (!is.null(tab_dico) && nrow(tab_dico) > 0) {

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
        }

      } else if (tipo_grafico == "radar") {

        warning(
          "Tipo de gráfico 'radar' señalado para la variable '", v,
          "', pero el constructor genérico de cruces aún no está implementado. ",
          "La variable se omitirá en este reporte.",
          call. = FALSE
        )

      } else {
        warning(
          "Tipo de gráfico '", tipo_grafico, "' no reconocido para la variable '",
          v, "'. Se omitirá en este reporte.",
          call. = FALSE
        )
      }

      if (!is.null(p)) {
        log_list[[length(log_list) + 1]] <- tibble::tibble(
          seccion      = sec,
          var          = v,
          estrato      = var_cruce,
          tipo_var     = tipo_v,
          list_name    = list_name_v,
          override     = override,
          tipo_grafico = tipo_grafico_final
        )

        idx_plot <- idx_plot + 1L
        plots_list[[idx_plot]]       <- p
        titulos_list[[idx_plot]]     <- titulo_slide
        resumenN_list[[idx_plot]]    <- resumen_n_txt
        seccion_por_plot[idx_plot]   <- sec
      } else {
        log_list[[length(log_list) + 1]] <- tibble::tibble(
          seccion      = sec,
          var          = v,
          estrato      = var_cruce,
          tipo_var     = tipo_v,
          list_name    = list_name_v,
          override     = override,
          tipo_grafico = NA_character_
        )
      }
    }
  }

  log_decisiones <- dplyr::bind_rows(log_list)

  # ---------------------------------------------------------------------------
  # 5. Construir DOCX (índice + gráficos)
  # ---------------------------------------------------------------------------
  if (!solo_lista && length(plots_list)) {

    if (!requireNamespace("officer", quietly = TRUE)) {
      stop(
        "Para exportar a Word se requiere el paquete 'officer'.",
        call. = FALSE
      )
    }

    pulso_azul <- "#004B8D"

    # 5.1 Leer plantilla (análogo a la lógica de PPT interna)
    if (is.null(template_docx)) {
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

    # 5.2 Primera página: índice de gráficas por sección
    if (length(plots_list)) {

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

      if (length(seccion_por_plot) == length(plots_list)) {

        secciones_unicas <- unique(seccion_por_plot)

        for (sec in secciones_unicas) {
          if (is.na(sec) || !nzchar(sec)) next

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

          doc <- officer::body_add_par(doc, "", style = "Normal")
        }

      } else {

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

      # 5.3 Gráficas: título externo + gráfico + pie (fuente / resumen N)
      for (i in seq_along(plots_list)) {

        p_i       <- plots_list[[i]]
        titulo_i  <- titulos_list[[i]]  %||% ""
        resumen_i <- resumenN_list[[i]] %||% NULL

        # TÍTULO EN WORD (arriba de la gráfica)
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

        # Ancho fijo ~ 15.5 cm en pulgadas
        width_in <- 15.5 / 2.54

        # Altura base sugerida por el graficador (si existe)
        alto_sugerido <- attr(p_i, "alto_word_sugerido", exact = TRUE)

        doc <- officer::body_add_gg(
          doc,
          value  = p_i,
          width  = width_in,
          height = alto_sugerido,
          style  = "Normal"
        )

        # Pie: fuente + (opcional) resumen N, en gris
        pie_line <- NULL

        if (!is.null(fuente) && nzchar(fuente)) {
          pie_line <- fuente
        }

        if (mostrar_resumen_n &&
            !is.null(resumen_i) && nzchar(resumen_i)) {

          if (is.null(pie_line)) {
            pie_line <- resumen_i
          } else {
            pie_line <- paste0(pie_line, "   |   ", resumen_i)
          }
        }

        if (!is.null(pie_line) && nzchar(pie_line)) {
          pie_word <- officer::fpar(
            officer::ftext(
              pie_line,
              prop = officer::fp_text(
                color       = "#7F7F7F",
                font.size   = 9,
                font.family = "Arial"
              )
            )
          )
          doc <- officer::body_add_fpar(doc, pie_word, style = "Normal")
        }

        doc <- officer::body_add_par(doc, "", style = "Normal")
      }
    }

    print(doc, target = path_word)
    if (mensajes_progreso) {
      message("Word de cruces generado en: ", normalizePath(path_word, winslash = "/"))
    }
  }

  invisible(list(
    plots          = plots_list,
    log_decisiones = log_decisiones
  ))
}
