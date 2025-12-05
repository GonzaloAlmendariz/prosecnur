# =============================================================================
# reporte_ppt_cruces()
# -----------------------------------------------------------------------------
# Generar PPT con gráficos a partir de cruces (var x estrato) por secciones,
# usando los graficadores propios del paquete (barras agrupadas, apiladas, dico).
# =============================================================================

#' Generar un reporte PPT a partir de cruces (SO/SM vs estrato) con graficadores propios
#'
#' `reporte_ppt_cruces()` toma una base de datos, el instrumento, una definición
#' de secciones y una variable de cruce (\code{var_cruce}) y genera:
#' \itemize{
#'   \item Una lista de gráficos \code{ggplot} (uno por variable de las secciones).
#'   \item Opcionalmente, un archivo PPT donde cada gráfico ocupa una diapositiva.
#' }
#'
#' A diferencia de \code{reporte_ppt()}, que trabaja con frecuencias simples,
#' esta función construye las tablas de cruce usando los mismos helpers que
#' \code{reporte_cruces()}:
#' \itemize{
#'   \item \code{tipo_pregunta()}
#'   \item \code{get_categorias()}
#'   \item \code{contar_por_opcion()}
#'   \item \code{denominador_validos()}
#'   \item \code{.has_var_or_dummies()}
#'   \item \code{label_variable()}
#' }
#'
#' Para cada variable \code{v} en \code{secciones} se construye internamente
#' una tabla de cruce \code{v x var_cruce} con:
#' \itemize{
#'   \item \code{categorias}: categorías de respuesta de \code{v}.
#'   \item \code{estratos}: categorías de \code{var_cruce}.
#'   \item \code{n_mat}: matriz de conteos (filas = categorías de \code{v},
#'         columnas = estratos).
#'   \item \code{N_estrato}: denominadores por estrato (vector nombrado).
#' }
#'
#' A partir de esa tabla, se generan los insumos para los graficadores:
#' \itemize{
#'   \item Barras apiladas:
#'     \item Cada fila es un estrato (\code{categoria} = etiqueta del estrato),
#'           \code{n_base} = \code{N_estrato[j]}.
#'     \item Los segmentos son las categorías de la variable \code{v}.
#'     \item Los colores de los segmentos se obtienen de
#'           \code{colores_apiladas_por_listname[[list_name_v]]}, igual que en
#'           \code{reporte_ppt()}, con soporte de objetos tipo
#'           \code{list(colores = ..., preset_barra_extra = ...)}.
#'   \item Barras agrupadas (cruces):
#'     \item Cada fila es un estrato (\code{categoria} = etiqueta del estrato).
#'     \item Las series (barras dentro de cada estrato) son las categorías de
#'           la variable \code{v}.
#'     \item Los colores de las series se intentan obtener primero de
#'           \code{colores_apiladas_por_listname[[list_name_v]]} (por
#'           \code{list_name}); en su defecto de
#'           \code{estilos_barras_agrupadas$colores_series} o de una paleta
#'           por defecto. \code{colores_estratos} se mantiene como respaldo si
#'           sus nombres coinciden con las etiquetas de serie.
#'   \item Dicotómicas:
#'     \item Cada estrato es un "indicador" distinto.
#'     \item Se calcula el porcentaje de la categoría "positiva" sobre el total
#'           (positiva + negativa) en ese estrato.
#'     \item Los colores se mantienen configurados por \code{estilos_dico}, igual
#'           que en \code{reporte_ppt()}.
#' }
#'
#' En todos los casos se eliminan:
#' \itemize{
#'   \item Estratos con \code{N_estrato == 0} (no se grafican).
#'   \item Categorías con suma de conteos 0 en todos los estratos.
#' }
#'
#' La lógica de exportación a PPT (plantilla, layouts, portada, contraportada,
#' bloques de fuente y resumen de N) es idéntica a la de \code{reporte_ppt()}.
#'
#' @param data Data frame o tibble con la base de datos a cruzar. Debe contener
#'   las variables de \code{secciones} y la variable \code{var_cruce}.
#' @param instrumento Objeto devuelto por \code{reporte_instrumento()}. Si es
#'   \code{NULL}, se intentará recuperar desde \code{attr(data, "instrumento_reporte")}.
#'   Debe contener al menos las tibbles \code{survey} y \code{choices} y,
#'   opcionalmente, \code{orders_list}.
#' @param secciones Lista nombrada de secciones; cada elemento es un vector de
#'   nombres de variables que se cruzarán con \code{var_cruce}.
#' @param var_cruce Nombre de la variable de estrato (columna de \code{data})
#'   con la que se cruzarán todas las variables de \code{secciones}.
#' @param path_ppt Ruta del archivo PPTX a generar cuando \code{solo_lista = FALSE}.
#' @param fuente Texto de fuente que se mostrará en el bloque de texto inferior
#'   izquierdo de cada diapositiva de gráficos.
#' @param sm_vars_force Vector opcional de nombres de variables que deben tratarse
#'   como \code{select_multiple} aunque el instrumento no las marque como tales.
#' @param weight_col Nombre de la variable de pesos (por defecto `"peso"`).
#' @param mostrar_todo Reservado para coherencia; no se usa explícitamente en
#'   esta función (el comportamiento depende de \code{get_categorias()}).
#' @param solo_lista Lógico; si \code{TRUE}, no se genera PPT y solo se devuelve
#'   una lista con \code{plots} y \code{log_decisiones}. Si \code{FALSE}, además
#'   se construye \code{path_ppt}.
#' @param incluir_titulo_var Lógico; si \code{TRUE}, el label de la variable se
#'   usa como título de la diapositiva (placeholder de título). Los gráficos se
#'   generan sin título interno.
#' @param mensajes_progreso Lógico; si \code{TRUE}, muestra mensajes por consola
#'   indicando el tipo de gráfico usado para cada variable.
#'
#' @param vars_dico Vector de nombres de variables que se deben graficar como
#'   dicotómicas mediante \code{graficar_dico()}, siempre que exista una pareja
#'   de etiquetas en \code{dico_labels_por_var} o \code{dico_labels_por_listname}.
#' @param vars_barras_apiladas Vector de nombres de variables que se deben
#'   graficar como barras horizontales apiladas mediante
#'   \code{graficar_barras_apiladas()}.
#' @param vars_barras_agrupadas Vector de nombres de variables que se deben
#'   graficar como barras agrupadas mediante \code{graficar_barras_agrupadas()}.
#' @param vars_radar Reservado; actualmente no se implementa salida tipo radar.
#'
#' @param listnames_dico Vector de \code{list_name} para los que, si una variable
#'   no aparece en \code{vars_dico}, se intentará por defecto tratarla como
#'   dicotómica.
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
#'   cada elemento puede ser:
#'   \itemize{
#'     \item Un vector nombrado de colores HEX (forma clásica).
#'     \item Una lista con elementos \code{colores} y
#'           \code{preset_barra_extra}, para controlar tanto colores como la
#'           barra extra en \code{graficar_barras_apiladas()}.
#'   }
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
#'   en barras apiladas/agrupadas (cuando no se usa un preset específico).
#'   Puede ser:
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
#' @param colores_estratos Vector nombrado opcional de colores HEX para los
#'   estratos en barras agrupadas (se mantiene como respaldo si sus nombres
#'   coinciden con las etiquetas de serie).
#'
#' @param opciones_excluir Vector de labels de opciones a excluir de la
#'   variable de interés (por ejemplo, categorías de no respuesta).
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
#' @param labels_override Lista nombrada opcional para sobrescribir etiquetas
#'   de preguntas (se usa en los títulos de las diapositivas).
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
reporte_ppt_cruces <- function(
    data,
    instrumento      = NULL,
    secciones,
    var_cruce,
    path_ppt         = "reporte_ppt_cruces.pptx",
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

    # Colores por estrato en barras agrupadas (respaldo)
    colores_estratos = NULL,

    # Opciones de la variable a excluir
    opciones_excluir = NULL,

    # Plantilla PPT
    template_pptx = NULL,

    # Texto de portada
    titulo_portada    = NULL,
    subtitulo_portada = NULL,
    fecha_portada     = NULL,

    # Resumen de N en bloque derecho
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

  # Helper: porcentajes ENTEROS que suman exactamente 100 (para apiladas / dico)
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

    raw_pct   <- n / total * 100
    floor_pct <- floor(raw_pct)

    resid <- as.integer(round(100 - sum(floor_pct)))
    frac  <- raw_pct - floor_pct

    if (resid > 0) {
      ord <- order(frac, decreasing = TRUE, na.last = TRUE)
      idx <- head(ord, resid)
      floor_pct[idx] <- floor_pct[idx] + 1L
    } else if (resid < 0) {
      resid_neg <- abs(resid)
      ord <- order(frac, decreasing = FALSE, na.last = TRUE)
      idx <- head(ord, resid_neg)
      floor_pct[idx] <- pmax(0L, floor_pct[idx] - 1L)
    }

    floor_pct
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

  .build_tab_barras_agrupadas_cruce <- function(crc, var_label) {

    n_cat <- length(crc$categorias)
    n_est <- length(crc$estr_labels)
    if (n_cat == 0L || n_est == 0L) return(NULL)

    # Proporciones por opción (fila) y estrato (columna)
    pct_mat <- matrix(NA_real_, nrow = n_cat, ncol = n_est)
    for (j in seq_len(n_est)) {
      Nj <- crc$N_estrato[j]
      if (Nj > 0) pct_mat[, j] <- crc$n_mat[, j] / Nj
    }

    # 1) Filtrar opciones sin información
    keep_cat <- rowSums(pct_mat, na.rm = TRUE) > 0
    keep_cat <- keep_cat & !is.na(crc$categorias)

    if (!any(keep_cat)) return(NULL)

    pct_mat   <- pct_mat[keep_cat, , drop = FALSE]
    cats_keep <- crc$categorias[keep_cat]
    n_cat     <- length(cats_keep)

    # 2) Filtrar estratos sin información
    keep_est <- colSums(pct_mat, na.rm = TRUE) > 0
    if (!any(keep_est)) return(NULL)

    pct_mat        <- pct_mat[, keep_est, drop = FALSE]
    estr_labels    <- crc$estr_labels[keep_est]
    N_estrato_keep <- crc$N_estrato[keep_est]
    n_est          <- length(estr_labels)

    # Convertir proporciones a enteros (no imponemos suma 100, puede >100)
    pct_mat_int <- matrix(NA_real_, nrow = n_cat, ncol = n_est)
    for (j in seq_len(n_est)) {
      col_raw <- pct_mat[, j] * 100
      pct_mat_int[, j] <- round(col_raw) / 100
    }

    # 3) Eje de categorías = ESTRATOS
    df <- tibble::tibble(
      categoria = estr_labels,
      n_base    = as.numeric(N_estrato_keep)
    )

    # 4) Series = OPCIONES de la variable
    cols_pct <- paste0("pct_", seq_len(n_cat))
    for (k in seq_len(n_cat)) {
      v <- as.numeric(pct_mat_int[k, ])
      v[!is.na(v) & abs(v) < 1e-12] <- NA_real_
      df[[cols_pct[k]]] <- v
    }

    etiquetas_series <- stats::setNames(
      as.character(cats_keep),
      cols_pct
    )

    list(
      data             = df,
      cols_porcentaje  = cols_pct,
      etiquetas_series = etiquetas_series,
      info_dim         = list(
        n_cat      = n_cat,
        n_estratos = n_est,
        N_estrato  = N_estrato_keep,
        N_total    = sum(N_estrato_keep, na.rm = TRUE)
      )
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

    # Para cada estrato: enteros que sumen 100 entre Sí / No
    pct_si <- rep(NA_real_, length(denom))
    for (j in seq_along(denom)) {
      if (denom[j] > 0) {
        p_pair   <- .pct_enteros_100(c(n_pos[j], n_neg[j])) # c(%Sí, %No)
        pct_si[j] <- p_pair[1]                              # 0–100 entero
      } else {
        pct_si[j] <- NA_real_
      }
    }

    tibble::tibble(
      indicador = as.character(crc$estr_labels),
      pct_si    = as.numeric(pct_si),   # escala 0–100, entero
      n_total   = as.numeric(denom)
    )
  }

  # ---------------------------------------------------------------------------
  # barras agrupadas con eje = estrato y series = categorías
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

    pct_si <- ifelse(denom > 0, n_pos / denom * 100, NA_real_)

    tibble::tibble(
      indicador = as.character(crc$estr_labels),
      pct_si    = as.numeric(pct_si),
      n_total   = as.numeric(denom)
    )
  }


  # Helper para SM "clásico": misma lógica que en reporte_ppt()
  .build_tab_barras_agrupadas_simple <- function(tab_freq) {
    if (!nrow(tab_freq)) return(NULL)

    n_total <- sum(tab_freq$n, na.rm = TRUE)

    pct_raw <- tab_freq$pct
    if (all(is.na(pct_raw))) return(NULL)

    # Detectar escala de pct (0–1 o 0–100)
    max_pct <- max(pct_raw, na.rm = TRUE)
    if (is.finite(max_pct) && max_pct <= 1 + 1e-8) {
      pct_0_100 <- pct_raw * 100
    } else {
      pct_0_100 <- pct_raw
    }

    # Convertir a enteros (no imponemos suma 100, en SM puede ser > 100)
    pct_int  <- round(pct_0_100)
    pct_prop <- pct_int / 100

    tibble::tibble(
      categoria = tab_freq$Opciones,
      n_base    = n_total,
      pct       = pct_prop
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
        message(
          "   - ", v, " x ", var_cruce, " → ", tipo_grafico,
          " (list_name = ", list_name_v, ") [", override, "]"
        )
      }

      p <- NULL
      tipo_grafico_final <- tipo_grafico

      if (tipo_grafico %in% c("barras_agrupadas", "barras_apiladas", "dico")) {

        if (tipo_grafico == "barras_agrupadas") {

          tab_agr <- .build_tab_barras_agrupadas_cruce(crc, var_label)
          if (!is.null(tab_agr)) {

            cols_porcentaje  <- tab_agr$cols_porcentaje
            etiquetas_series <- tab_agr$etiquetas_series

            args_extra <- list()

            # COLORES por defecto (si el usuario no los definió)
            if (is.null(estilos_barras_agrupadas$colores_series)) {

              pulso_azul <- "#004B8D"

              colores_series <- stats::setNames(
                rep(pulso_azul, length(etiquetas_series)),
                etiquetas_series
              )

              args_extra$colores_series <- colores_series
            }

            # Tamaño automático de texto para el caso general (SO / multi-series)
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

            # -----------------------------------------------------------------
            # CASO ESPECIAL: SM + barras agrupadas → UNA DIAPOSITIVA POR ESTRATO
            # -----------------------------------------------------------------
            if (tipo_v == "sm") {

              n_estratos <- length(crc$estr_labels)
              N_estrato  <- crc$N_estrato

              for (j in seq_len(n_estratos)) {

                tab_j <- .build_tab_barras_agrupadas_sm_estrato(crc, j)
                if (is.null(tab_j) || !nrow(tab_j)) next

                # Una sola serie: "Porcentaje"
                cols_porcentaje  <- "pct"
                etiquetas_series <- c(pct = "Porcentaje")

                args_extra <- list()

                # COLORES por defecto (si el usuario no los definió)
                if (is.null(estilos_barras_agrupadas$colores_series)) {

                  pulso_azul <- "#1B679D"

                  colores_series <- stats::setNames(
                    rep(pulso_azul, length(etiquetas_series)),
                    etiquetas_series
                  )

                  args_extra$colores_series <- colores_series
                }

                # Tamaño automático de texto, ahora según Nº de categorías
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
                  args_extra$size_texto_barras <- size_auto
                }

                args_barras_j <- c(
                  list(
                    data             = tab_j,
                    var_categoria    = "categoria",  # opciones SM en eje Y (via coord_flip)
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

                p_j <- do.call(graficar_barras_agrupadas, args_barras_j)

                # OJO: aquí ya NO quitamos labels del eje X, porque ahora
                # el eje tiene las OPCIONES, que sí queremos ver.
                # (Si hay coord_flip en el graficador, esto será el eje Y visual.)

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

                # Título de diapositiva: "Título - Estrato"
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
            args_barras <- c(
              list(
                data             = tab_agr$data,
                var_categoria    = "categoria",
                var_n            = "n_base",
                cols_porcentaje  = cols_porcentaje,
                etiquetas_series = etiquetas_series,
                escala_valor     = "proporcion_1",
                mostrar_valores  = TRUE,
                titulo           = NULL,       # sin título interno
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

            # --------------------------------------------------
            # Inversión por list_name desde estilos_barras_apiladas
            # --------------------------------------------------
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
                # - Si hay preset → dejar NULL (Top2Box verde)
                # - Si NO hay preset → N= en azul (#092147)
                color_barra_extra = if (!is.null(preset_extra)) {
                  NULL
                } else {
                  "#092147"
                },

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

                if (is.null(colores_series) &&
                    !is.null(colores_estratos) &&
                    all(etiquetas_series %in% names(colores_estratos))) {

                  cs <- colores_estratos[etiquetas_series]
                  colores_series <- stats::setNames(
                    as.character(cs),
                    etiquetas_series
                  )
                }

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
                  var_categoria    = "categoria",  # estrato
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
                  data                = tab_dico,
                  var_indicador       = "indicador",
                  var_porcentaje_si   = "pct_si",
                  var_n               = "n_total",
                  escala_valor        = "proporcion_100",
                  etiqueta_si         = labels_dico[1],
                  etiqueta_no         = labels_dico[2],
                  titulo              = titulo_plot,
                  subtitulo           = NULL,
                  nota_pie            = nota_pie_plot,
                  incluir_n_en_titulo = FALSE,
                  exportar            = "rplot"
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
        plots_list[[idx_plot]]     <- p
        titulos_list[[idx_plot]]   <- titulo_slide
        resumenN_list[[idx_plot]]  <- resumen_n_txt
        seccion_por_plot[idx_plot] <- sec
      } else if (tipo_grafico %in% c("barras_apiladas", "dico")) {
        # Solo registramos aquí como omitida si falló apiladas/dico;
        # en barras_agrupadas ya se registró antes (cuando tab_agr no es NULL).
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
  # 5. PPT (portada + secciones + diapositivas de gráficos de cruces)
  # ---------------------------------------------------------------------------
  if (!solo_lista && length(plots_list)) {
    if (!requireNamespace("officer", quietly = TRUE) ||
        !requireNamespace("rvg", quietly = TRUE)) {
      stop(
        "Para exportar a PPT se requieren los paquetes 'officer' y 'rvg'.",
        call. = FALSE
      )
    }

    # 5.1. Leer plantilla
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

    # 5.2. Info de layouts
    layout_info <- tryCatch(
      officer::layout_summary(doc),
      error = function(e) NULL
    )

    tiene_layout_graficos        <- FALSE
    layout_graficos              <- "Blank"
    usar_pic_placeholder         <- FALSE
    tiene_layout_title_slide     <- FALSE
    tiene_layout_contraportada   <- FALSE
    tiene_layout_section_header  <- FALSE

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
      if ("Section Header" %in% layout_info$layout) {
        tiene_layout_section_header <- TRUE
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

    # 5.3. Portada (Title Slide), si corresponde
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

    # 5.4. Diapositivas de gráficos (con Section Header si existe)
    if (length(plots_list)) {
      for (i in seq_along(plots_list)) {

        p        <- plots_list[[i]]
        st       <- titulos_list[[i]]   %||% NULL
        resumenN <- resumenN_list[[i]]  %||% NULL

        # Diapositiva de sección si cambia la sección (y hay layout Section Header)
        if (tiene_layout_section_header &&
            length(seccion_por_plot) >= i) {

          sec_i <- seccion_por_plot[i]

          if (!is.null(sec_i) && nzchar(sec_i) &&
              (i == 1L || !identical(sec_i, seccion_por_plot[i - 1L]))) {

            doc <- officer::add_slide(
              doc,
              layout = "Section Header",
              master = "Office Theme"
            )

            loc_sec_title <- tryCatch(
              officer::ph_location_type(type = "title"),
              error = function(e) {
                tryCatch(
                  officer::ph_location_type(type = "ctrTitle"),
                  error = function(e2) NULL
                )
              }
            )

            if (!is.null(loc_sec_title)) {
              doc <- tryCatch(
                officer::ph_with(doc, sec_i, location = loc_sec_title),
                error = function(e) doc
              )
            }
          }
        }

        # Diapositiva de gráfico
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

    # 5.5. Contraportada (si existe)
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

    # 5.6. Guardar
    print(doc, target = path_ppt)
    if (mensajes_progreso) {
      message("PPT de cruces generado en: ", normalizePath(path_ppt, winslash = "/"))
    }
  }

  invisible(list(
    plots          = plots_list,
    log_decisiones = log_decisiones
  ))
}
