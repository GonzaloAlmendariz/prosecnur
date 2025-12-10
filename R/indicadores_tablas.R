#' Generar tablas de indicadores a partir de un YAML
#'
#' Esta función:
#' 1. Lee un archivo de configuración YAML de indicadores.
#' 2. Calcula las tablas para cada indicador según su \code{tipo}
#'    (\code{tabla_conceptos}, \code{tabla_compuesta}, \code{freq_multiple}).
#' 3. Opcionalmente, exporta todas las tablas a un archivo Excel con formato.
#'
#' @param rp_data Data frame (o tibble) con la base ya limpia/recodificada.
#' @param config Ruta al archivo YAML de indicadores (character) o lista ya
#'   leída con \code{yaml::read_yaml()}.
#' @param output_xlsx Ruta del archivo Excel a generar. Si es \code{NULL}
#'   (por defecto), no se exporta a Excel y la función solo devuelve las tablas
#'   en una lista.
#' @param hoja Nombre de la hoja donde se escribirán las tablas. Por defecto
#'   \code{"Indicadores"}.
#' @param estilos Opcional: lista de estilos de \pkg{openxlsx}. Puede ser la
#'   salida de una función tipo \code{mk_styles_spss()} o similar. Si es
#'   \code{NULL}, se usan estilos simples por defecto.
#' @param fuente Texto de la fuente a mostrar debajo de cada tabla (por ejemplo,
#'   "Pulso PUCP 2025"). Si es \code{NULL}, se escribe una fila vacía
#'   pero con la línea de cierre de tabla.
#'
#' @return Una lista nombrada, donde cada elemento corresponde a un indicador y
#'   contiene:
#'   \itemize{
#'     \item \code{id}, \code{titulo}, \code{tipo}, \code{grafico}, \code{notas}
#'     \item \code{tabla}: data.frame con los resultados.
#'     \item \code{tablas_por_variable} (opcional, en tipo freq_multiple con
#'           modo = "por_variable"): lista de tablas de frecuencias simples.
#'   }
#'
#' Si \code{output_xlsx} no es \code{NULL}, también genera un archivo Excel.
#'
#' @export
indicadores_tablas <- function(rp_data,
                               config,
                               output_xlsx = NULL,
                               hoja = "Indicadores",
                               estilos = NULL,
                               fuente = NULL) {

  if (!requireNamespace("yaml", quietly = TRUE)) {
    stop("Se requiere el paquete 'yaml' para leer la configuración.", call. = FALSE)
  }

  # ---------------------------------------------------------------------------
  # 1. Leer configuración YAML
  # ---------------------------------------------------------------------------
  cfg <- if (is.character(config)) {
    yaml::read_yaml(config)
  } else {
    config
  }

  if (is.null(cfg$indicadores)) {
    stop("El YAML no contiene la clave 'indicadores'.", call. = FALSE)
  }

  # Variable de pesos (opcional)
  peso_var <- cfg$peso
  pesos <- .get_pesos(rp_data, peso_var)

  # ---------------------------------------------------------------------------
  # 2. Calcular cada indicador
  # ---------------------------------------------------------------------------
  resultados <- list()

  message("Calculando tablas para ", length(cfg$indicadores), " indicadores...")

  for (ind in cfg$indicadores) {
    tipo <- ind$tipo

    res_ind <- switch(
      tipo,
      "tabla_conceptos" = .calc_tabla_conceptos(rp_data, ind, pesos = pesos),
      "tabla_compuesta" = .calc_tabla_compuesta(rp_data, ind, pesos = pesos),
      "freq_multiple"   = .calc_freq_multiple(rp_data, ind, pesos = pesos),
      {
        warning("Tipo de indicador desconocido: ", tipo)
        NULL
      }
    )

    if (!is.null(res_ind)) {
      res_ind$id      <- ind$id
      res_ind$titulo  <- ind$titulo
      res_ind$tipo    <- tipo
      res_ind$grafico <- ind$grafico %||% "ninguno"
      res_ind$notas   <- ind$notas %||% character(0)
      res_ind$modo    <- ind$modo %||% NULL
      resultados[[ind$id]] <- res_ind
    }
  }

  message("Se han calculado ", length(resultados), " indicadores con éxito.")

  # ---------------------------------------------------------------------------
  # 3. Exportar a Excel (opcional)
  # ---------------------------------------------------------------------------
  if (!is.null(output_xlsx)) {
    if (!requireNamespace("openxlsx", quietly = TRUE)) {
      stop("Se requiere el paquete 'openxlsx' para exportar a Excel.", call. = FALSE)
    }

    wb <- openxlsx::createWorkbook()
    openxlsx::addWorksheet(wb, hoja)

    if (is.null(estilos)) {
      estilos <- .estilos_sencillos()
    }

    estilos <- .normalizar_estilos_excel(estilos)

    current_row <- 1L
    for (id_ind in names(resultados)) {
      bloque <- resultados[[id_ind]]

      message("Escribiendo indicador '", id_ind, "' (tipo: ", bloque$tipo, ") en Excel...")

      # Caso especial: freq_multiple con modo = "por_variable"
      if (identical(bloque$tipo, "freq_multiple") &&
          identical(bloque$modo, "por_variable") &&
          !is.null(bloque$tablas_por_variable)) {

        for (sub in bloque$tablas_por_variable) {
          sub_ind <- list(
            tabla  = sub$tabla,
            titulo = sub$titulo,
            tipo   = "freq_multiple"
          )
          escrito <- .escribir_indicador_excel(
            wb          = wb,
            sheet       = hoja,
            indicador   = sub_ind,
            row_start   = current_row,
            col_start   = 1L,
            estilos     = estilos,
            fuente      = fuente
          )
          current_row <- escrito$next_row
        }

      } else {
        escrito <- .escribir_indicador_excel(
          wb          = wb,
          sheet       = hoja,
          indicador   = bloque,
          row_start   = current_row,
          col_start   = 1L,
          estilos     = estilos,
          fuente      = fuente
        )
        current_row <- escrito$next_row
      }
    }

    openxlsx::saveWorkbook(wb, file = output_xlsx, overwrite = TRUE)
    message("Archivo de indicadores guardado en: ",
            normalizePath(output_xlsx, winslash = "/"))
  }

  invisible(resultados)
}

# ===========================================================================
# UTILIDADES BÁSICAS
# ===========================================================================

`%||%` <- function(x, y) {
  if (is.null(x)) y else x
}

.get_pesos <- function(data, peso_var) {
  if (is.null(peso_var) || !nzchar(peso_var) || !peso_var %in% names(data)) {
    rep(1, nrow(data))
  } else {
    w <- data[[peso_var]]
    w[is.na(w)] <- 0
    w
  }
}

# ===========================================================================
# HELPERS PARA CONDICIONES
# ===========================================================================

.valido_helper <- function(..., .n) {
  args <- list(...)
  if (length(args) == 0L) {
    return(rep(TRUE, .n))
  }
  masks <- lapply(args, function(v) !is.na(v))
  Reduce("&", masks)
}

.todo_prefijo_helper <- function(data, prefijo, valor = 1) {
  vars_pref <- grep(paste0("^", prefijo), names(data), value = TRUE)
  if (length(vars_pref) == 0L) {
    return(rep(FALSE, nrow(data)))
  }
  sub <- data[, vars_pref, drop = FALSE]
  apply(sub, 1L, function(row) {
    if (all(is.na(row))) {
      FALSE
    } else {
      all(!is.na(row) & row == valor)
    }
  })
}

.valido_prefijo_helper <- function(data, prefijo) {
  vars_pref <- grep(paste0("^", prefijo), names(data), value = TRUE)
  if (length(vars_pref) == 0L) {
    return(rep(FALSE, nrow(data)))
  }
  sub <- data[, vars_pref, drop = FALSE]
  apply(!is.na(sub), 1L, any)
}

.eval_condicion_fila <- function(data, condicion) {
  if (is.null(condicion) || !nzchar(condicion)) {
    return(rep(TRUE, nrow(data)))
  }

  env <- list2env(as.list(data), parent = parent.frame())

  env$valido <- function(...) .valido_helper(..., .n = nrow(data))
  env$valid  <- env$valido

  env$todo_prefijo <- function(prefijo, valor = 1) {
    .todo_prefijo_helper(data, prefijo = prefijo, valor = valor)
  }

  env$valido_prefijo <- function(prefijo) {
    .valido_prefijo_helper(data, prefijo = prefijo)
  }

  res <- try(
    eval(parse(text = condicion), envir = env),
    silent = TRUE
  )

  if (inherits(res, "try-error")) {
    warning("No se pudo evaluar la condición '", condicion,
            "'. Se usará todo NA para esa fila.\nDetalle: ",
            conditionMessage(attr(res, "condition")))
    return(rep(NA, nrow(data)))
  }

  if (!is.logical(res) || length(res) != nrow(data)) {
    stop("La condición '", condicion,
         "' no devolvió un vector lógico de longitud nrow(data).")
  }
  res
}

.eval_condicion_vector <- function(x, condicion) {
  if (is.null(condicion) || !nzchar(condicion)) {
    return(!is.na(x))
  }
  if (condicion == "no_es_na") {
    return(!is.na(x))
  }

  expr <- if (grepl("\\bx\\b", condicion)) {
    condicion
  } else {
    paste0("x ", condicion)
  }

  res <- eval(parse(text = expr), envir = list(x = x))
  if (!is.logical(res) || length(res) != length(x)) {
    stop("La condición '", condicion,
         "' no devolvió un vector lógico de longitud length(x).")
  }
  res
}

# ===========================================================================
# CALCULAR tabla_compuesta
# ===========================================================================
.calc_tabla_compuesta <- function(data, ind_cfg, pesos) {

  filas_cfg <- ind_cfg$filas
  cols_cfg  <- ind_cfg$columnas

  n_por_fila <- numeric(length(filas_cfg))
  names(n_por_fila) <- vapply(filas_cfg, function(f) f$id, FUN.VALUE = character(1))

  for (i in seq_along(filas_cfg)) {
    f <- filas_cfg[[i]]
    cond <- .eval_condicion_fila(data, f$condicion)
    cond_log <- cond
    cond_log[is.na(cond_log)] <- FALSE
    n_por_fila[i] <- sum(pesos[cond_log], na.rm = TRUE)
  }

  res_mat <- matrix(
    NA_real_,
    nrow = length(filas_cfg),
    ncol = length(cols_cfg),
    dimnames = list(
      vapply(filas_cfg, function(f) f$label, FUN.VALUE = character(1)),
      vapply(cols_cfg,  function(c) c$label, FUN.VALUE = character(1))
    )
  )

  ids_filas <- vapply(filas_cfg, function(f) f$id, FUN.VALUE = character(1))

  for (j in seq_along(cols_cfg)) {
    col_def <- cols_cfg[[j]]
    tipo    <- col_def$tipo

    if (tipo == "conteo_cond_fila") {

      res_mat[, j] <- n_por_fila

    } else if (tipo == "proporcion_sobre_total") {

      ref_id <- col_def$referencia_total_fila
      if (is.null(ref_id)) {
        stop("En 'proporcion_sobre_total' se requiere 'referencia_total_fila'.")
      }
      idx_total <- match(ref_id, ids_filas)
      if (is.na(idx_total)) {
        stop("No se encontró la fila de referencia_total_fila = '", ref_id, "'.")
      }
      denom <- n_por_fila[idx_total]
      res_mat[, j] <- ifelse(denom > 0, n_por_fila / denom, NA_real_)  # 0–1

    } else {

      warning("Tipo de columna no soportado en tabla_compuesta: ", tipo)
    }
  }

  tabla <- as.data.frame(res_mat, check.names = FALSE)
  tabla <- cbind(
    Fila = vapply(filas_cfg, function(f) f$label, FUN.VALUE = character(1)),
    tabla
  )
  list(tabla = tabla)
}

# -----------------------------------------------------------------------------
# TABLA DE FRECUENCIA SIMPLE (para modo = "por_variable")
# -----------------------------------------------------------------------------
.tabla_frecuencia_simple <- function(x,
                                     pesos,
                                     usar_labels   = TRUE,
                                     incluir_total = TRUE,
                                     total_label   = "Total") {

  n <- length(x)
  if (length(pesos) != n) {
    pesos <- rep(1, n)
  }

  lbls <- attr(x, "labels", exact = TRUE)
  x_chr <- as.character(x)

  usar_labels_ef <- usar_labels && !is.null(lbls) && length(lbls) > 0

  if (usar_labels_ef) {
    # Detectar automáticamente dónde están los códigos
    lbl_vals_chr  <- as.character(unname(lbls))
    lbl_names_chr <- as.character(names(lbls))

    ux <- unique(x_chr[!is.na(x_chr) & nzchar(x_chr)])

    overlap_vals  <- length(intersect(lbl_vals_chr,  ux))
    overlap_names <- length(intersect(lbl_names_chr, ux))

    if (overlap_vals == 0 && overlap_names == 0) {
      usar_labels_ef <- FALSE
    } else {
      if (overlap_names >= overlap_vals) {
        # Códigos en NAMES, etiquetas en valores (tu caso)
        cod_chr   <- lbl_names_chr
        etiquetas <- lbl_vals_chr
      } else {
        # Códigos en valores, etiquetas en NAMES (haven clásico)
        cod_chr   <- lbl_vals_chr
        etiquetas <- if (all(nzchar(lbl_names_chr))) lbl_names_chr else lbl_vals_chr
      }

      mask_valid <- !is.na(x_chr) & x_chr %in% cod_chr

      if (!any(mask_valid)) {
        usar_labels_ef <- FALSE
      } else {
        n_vec <- vapply(
          seq_along(cod_chr),
          function(k) {
            sum(pesos[!is.na(x_chr) & x_chr == cod_chr[k]], na.rm = TRUE)
          },
          numeric(1)
        )

        total_w <- sum(pesos[mask_valid], na.rm = TRUE)
      }
    }
  }

  if (!usar_labels_ef) {
    # Modo genérico: usar los valores tal cual
    mask_valid <- !is.na(x_chr) & nzchar(x_chr)

    if (!any(mask_valid)) {
      return(data.frame(
        Opcion = character(0),
        n      = numeric(0),
        `%`    = numeric(0),
        check.names = FALSE,
        stringsAsFactors = FALSE
      ))
    }

    cod_chr   <- sort(unique(x_chr[mask_valid]))
    etiquetas <- cod_chr

    n_vec <- vapply(
      cod_chr,
      function(v) sum(pesos[mask_valid & x_chr == v], na.rm = TRUE),
      numeric(1)
    )

    total_w <- sum(pesos[mask_valid], na.rm = TRUE)
  }

  pct_vec <- if (total_w > 0) n_vec / total_w else NA_real_

  tabla <- data.frame(
    Opcion = etiquetas,
    n      = n_vec,
    `%`    = pct_vec,
    check.names = FALSE,
    stringsAsFactors = FALSE
  )

  if (incluir_total) {
    fila_total <- data.frame(
      Opcion = total_label,
      n      = total_w,
      `%`    = if (total_w > 0) 1 else NA_real_,
      check.names = FALSE,
      stringsAsFactors = FALSE
    )
    tabla <- rbind(tabla, fila_total)
  }

  tabla
}

# -----------------------------------------------------------------------------
# CALCULAR freq_multiple
# -----------------------------------------------------------------------------
.calc_freq_multiple <- function(data, ind_cfg, pesos) {

  # Modo: "dummy" (multi-respuesta clásica) o "por_variable"
  modo <- ind_cfg$modo %||% "dummy"

  # --------------------------------------------------------
  # A. MODO "POR_VARIABLE": una tabla de frecuencias por var
  # --------------------------------------------------------
  if (identical(modo, "por_variable")) {

    vars_cfg <- ind_cfg$vars
    incluir_total <- isTRUE(ind_cfg$opciones$incluir_total)
    total_label   <- ind_cfg$opciones$total_label %||% "Total"

    tablas_var <- list()

    for (v_def in vars_cfg) {
      vname <- v_def$var
      vid   <- v_def$id   %||% vname
      vlab  <- v_def$label %||% vname

      if (!vname %in% names(data)) {
        warning("Variable no encontrada en rp_data: ", vname)
        next
      }

      x <- data[[vname]]

      tabla_v <- .tabla_frecuencia_simple(
        x             = x,
        pesos         = pesos,
        usar_labels   = TRUE,
        incluir_total = incluir_total,
        total_label   = total_label
      )

      tablas_var[[vid]] <- list(
        titulo = vlab,
        tabla  = tabla_v
      )
    }

    return(list(
      tablas_por_variable = tablas_var,
      modo                = "por_variable"
    ))
  }

  # --------------------------------------------------------
  # B. MODO "DUMMY": filas = variables (multi-respuesta)
  # --------------------------------------------------------
  vars_cfg <- ind_cfg$vars
  usar_labels   <- isTRUE(ind_cfg$usar_labels)
  valor_si      <- ind_cfg$valor_si %||% 1
  incluir_total <- isTRUE(ind_cfg$incluir_total)
  total_label   <- ind_cfg$total_label %||% "Total"

  # Caso A: vars es vector de nombres de variables (p.ej. "p106.1", "p106.2"...)
  if (is.character(vars_cfg)) {
    vars_list <- lapply(vars_cfg, function(v) list(id = v, label = NULL, var = v))
  } else {
    # Caso B: lista con id/label/var
    vars_list <- lapply(vars_cfg, function(e) {
      list(
        id    = e$id    %||% e$var,
        label = e$label %||% e$var,
        var   = e$var
      )
    })
  }

  filas <- list()
  n_total_global <- 0

  for (v_def in vars_list) {
    vname <- v_def$var
    if (!vname %in% names(data)) {
      warning("Variable no encontrada en rp_data: ", vname)
      next
    }
    x <- data[[vname]]

    mask_valid <- !is.na(x)
    mask_si    <- mask_valid & (x == valor_si)
    n_si       <- sum(pesos[mask_si], na.rm = TRUE)
    n_total    <- sum(pesos[mask_valid], na.rm = TRUE)

    n_total_global <- max(n_total_global, n_total)
    pct_si   <- if (n_total > 0) n_si / n_total else NA_real_

    lbl <- v_def$label
    if (usar_labels && is.null(lbl)) {
      lbl <- attr(x, "label") %||% vname
    }
    if (is.null(lbl)) lbl <- vname

    filas[[length(filas) + 1L]] <- data.frame(
      Opcion = lbl,
      n      = n_si,
      `%`    = pct_si,
      stringsAsFactors = FALSE,
      check.names = FALSE
    )
  }

  if (length(filas) == 0) {
    tabla <- data.frame(
      Opcion = character(0),
      n      = numeric(0),
      `%`    = numeric(0),
      check.names = FALSE,
      stringsAsFactors = FALSE
    )
  } else {
    tabla <- do.call(rbind, filas)
  }

  if (incluir_total && nrow(tabla) > 0) {
    fila_total <- data.frame(
      Opcion = total_label,
      n      = n_total_global,
      `%`    = 1,
      stringsAsFactors = FALSE,
      check.names = FALSE
    )
    tabla <- rbind(tabla, fila_total)
  }

  list(
    tabla = tabla,
    modo  = "dummy"
  )
}

# ===========================================================================
# CALCULAR tabla_conceptos
# ===========================================================================
.calc_tabla_conceptos <- function(data, ind_cfg, pesos) {

  filas_cfg <- ind_cfg$filas
  cols_cfg  <- ind_cfg$columnas

  filas_labels <- vapply(filas_cfg, function(f) f$label, FUN.VALUE = character(1))

  res_mat <- matrix(
    NA_real_,
    nrow = length(filas_cfg),
    ncol = length(cols_cfg),
    dimnames = list(
      filas_labels,
      vapply(cols_cfg, function(c) c$label, FUN.VALUE = character(1))
    )
  )

  for (i in seq_along(filas_cfg)) {
    fila_cfg <- filas_cfg[[i]]

    grupos <- list()
    if (!is.null(fila_cfg$grupos)) {
      for (g_name in names(fila_cfg$grupos)) {
        g_cfg  <- fila_cfg$grupos[[g_name]]
        vars_g <- g_cfg$vars

        if (!all(vars_g %in% names(data))) {
          faltan <- vars_g[!vars_g %in% names(data)]
          warning(
            "En la fila '", fila_cfg$id, "' (grupo '", g_name,
            "') faltan variables en rp_data: ",
            paste(faltan, collapse = ", "),
            ". Se llenará con NA."
          )
          v <- rep(NA_real_, nrow(data))
        } else if (length(vars_g) == 1L) {
          v <- data[[vars_g]]
        } else {
          v <- rowSums(data[, vars_g, drop = FALSE], na.rm = TRUE)
        }

        grupos[[g_name]] <- v
      }
    }

    vals_fila <- list()

    for (j in seq_along(cols_cfg)) {
      col_def <- cols_cfg[[j]]
      tipo    <- col_def$tipo
      valor   <- NA_real_

      if (tipo %in% c("suma", "conteo_cond", "proporcion_cond",
                      "media", "mediana", "minimo", "maximo")) {

        var_ref <- col_def$var
        if (is.null(var_ref)) {
          stop("En columnas de tipo ", tipo, " se requiere 'var'.")
        }
        if (!startsWith(var_ref, "@")) {
          stop("Por ahora solo se soportan referencias a grupos (@nombre).")
        }
        g_name <- substring(var_ref, 2L)
        if (is.null(grupos[[g_name]])) {
          warning("No se encontró el grupo '", g_name,
                  "' en la fila ", fila_cfg$id, ". Se usa NA.")
          x <- rep(NA_real_, nrow(data))
        } else {
          x <- grupos[[g_name]]
        }

        w <- pesos
        mask_base <- !is.na(x)

        condicion <- col_def$condicion
        if (!is.null(condicion)) {
          cond_vec <- .eval_condicion_vector(x, condicion)
          mask_base <- mask_base & cond_vec
        }

        if (tipo == "suma") {

          valor <- sum(x[mask_base] * w[mask_base], na.rm = TRUE)

        } else if (tipo == "conteo_cond") {

          valor <- sum(w[mask_base], na.rm = TRUE)

        } else if (tipo == "proporcion_cond") {

          num   <- sum(w[mask_base], na.rm = TRUE)
          denom <- sum(w[!is.na(x)],  na.rm = TRUE)
          valor <- if (denom > 0) num / denom else NA_real_  # 0–1

        } else if (tipo == "media") {

          num   <- sum(x[mask_base] * w[mask_base], na.rm = TRUE)
          denom <- sum(w[mask_base], na.rm = TRUE)
          valor <- if (denom > 0) num / denom else NA_real_

        } else if (tipo == "mediana") {

          valor <- stats::median(x[!is.na(x)], na.rm = TRUE)

        } else if (tipo == "minimo") {

          valor <- suppressWarnings(min(x, na.rm = TRUE))

        } else if (tipo == "maximo") {

          valor <- suppressWarnings(max(x, na.rm = TRUE))
        }

      } else if (tipo == "proporcion_rel") {

        num_id <- col_def$numerador
        den_id <- col_def$denominador
        if (is.null(vals_fila[[num_id]]) || is.null(vals_fila[[den_id]])) {
          stop("En 'proporcion_rel' los ids numerador/denominador deben haberse calculado antes.")
        }
        num   <- vals_fila[[num_id]]
        den   <- vals_fila[[den_id]]
        valor <- if (den > 0) num / den else NA_real_  # 0–1

      } else {

        warning("Tipo de columna no soportado en tabla_conceptos: ", tipo)
      }

      vals_fila[[col_def$id]] <- valor
      res_mat[i, j] <- valor
    }
  }

  tabla <- as.data.frame(res_mat, check.names = FALSE)
  tabla <- cbind(
    Fila = filas_labels,
    tabla
  )

  list(tabla = tabla)
}

# ===========================================================================
# ESTILOS POR DEFECTO Y NORMALIZACIÓN
# ===========================================================================
.estilos_sencillos <- function() {
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("Se requiere 'openxlsx' para los estilos.", call. = FALSE)
  }
  list(
    titulo = openxlsx::createStyle(
      fontSize       = 11,
      textDecoration = "italic",
      halign         = "left",
      valign         = "center",
      wrapText       = TRUE
    ),
    notas = openxlsx::createStyle(
      fontSize       = 9,
      textDecoration = "italic",
      halign         = "left",
      valign         = "top",
      wrapText       = TRUE
    ),
    header = openxlsx::createStyle(
      fontSize       = 10,
      textDecoration = "bold",
      border         = c("top", "bottom"),
      borderStyle    = "thin",
      halign         = "center",
      valign         = "center",
      wrapText       = TRUE
    ),
    cuerpo = openxlsx::createStyle(
      fontSize       = 10,
      halign         = "center",
      valign         = "center",
      wrapText       = TRUE
    )
  )
}

.normalizar_estilos_excel <- function(estilos) {

  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("Se requiere 'openxlsx' para los estilos.", call. = FALSE)
  }

  if (is.null(estilos$titulo) || !inherits(estilos$titulo, "Style")) {
    if (!is.null(estilos$q_title) && inherits(estilos$q_title, "Style")) {
      estilos$titulo <- estilos$q_title
    } else if (!is.null(estilos$sec_title) && inherits(estilos$sec_title, "Style")) {
      estilos$titulo <- estilos$sec_title
    } else {
      estilos$titulo <- openxlsx::createStyle(
        fontSize       = 11,
        textDecoration = "italic",
        halign         = "left",
        valign         = "center",
        wrapText       = TRUE
      )
    }
  }

  if (is.null(estilos$header) || !inherits(estilos$header, "Style")) {
    estilos$header <- openxlsx::createStyle(
      fontSize       = 10,
      textDecoration = "bold",
      border         = c("top", "bottom"),
      borderStyle    = "thin",
      halign         = "center",
      valign         = "center",
      wrapText       = TRUE
    )
  }

  if (is.null(estilos$cuerpo) || !inherits(estilos$cuerpo, "Style")) {
    estilos$cuerpo <- openxlsx::createStyle(
      fontSize       = 10,
      halign         = "center",
      valign         = "center",
      wrapText       = TRUE
    )
  }

  if (is.null(estilos$body_txt) || !inherits(estilos$body_txt, "Style")) {
    estilos$body_txt <- estilos$cuerpo
  }
  if (is.null(estilos$body_int) || !inherits(estilos$body_int, "Style")) {
    estilos$body_int <- estilos$cuerpo
  }
  if (is.null(estilos$body_pct) || !inherits(estilos$body_pct, "Style")) {
    estilos$body_pct <- openxlsx::createStyle(
      fontSize       = 10,
      numFmt         = "0.0%",
      halign         = "right",
      valign         = "center",
      wrapText       = TRUE
    )
  }
  if (is.null(estilos$body_num) || !inherits(estilos$body_num, "Style")) {
    estilos$body_num <- openxlsx::createStyle(
      fontSize       = 10,
      numFmt         = "0.0",
      halign         = "right",
      valign         = "center",
      wrapText       = TRUE
    )
  }

  if (is.null(estilos$notas) || !inherits(estilos$notas, "Style")) {
    estilos$notas <- estilos$body_txt
  }

  if (is.null(estilos$fuente) || !inherits(estilos$fuente, "Style")) {
    estilos$fuente <- openxlsx::createStyle(
      fontSize       = 9,
      halign         = "left",
      valign         = "center",
      fontColour     = "#808080",
      wrapText       = TRUE
    )
  }

  estilos$table_end <- openxlsx::createStyle(
    border       = c("top"),
    borderStyle  = "thin",
    borderColour = "#000000"
  )

  estilos
}

# ===========================================================================
# ESCRIBIR UN INDICADOR EN EXCEL
# ===========================================================================
.escribir_indicador_excel <- function(wb,
                                      sheet,
                                      indicador,
                                      row_start = 1L,
                                      col_start = 1L,
                                      estilos,
                                      fuente = NULL) {

  tabla   <- indicador$tabla
  titulo  <- indicador$titulo %||% indicador$id
  tipo    <- indicador$tipo  %||% NA_character_

  n_filas_tabla <- nrow(tabla)
  n_cols_tabla  <- ncol(tabla)

  r <- row_start

  last_col <- if (n_cols_tabla > 0) {
    col_start + n_cols_tabla - 1L
  } else {
    col_start
  }

  # ------------------------------------------------------------------------
  # 1. TÍTULO DEL INDICADOR
  # ------------------------------------------------------------------------
  openxlsx::writeData(wb, sheet, titulo, startRow = r, startCol = col_start)
  if (last_col > col_start) {
    openxlsx::mergeCells(
      wb, sheet,
      rows = r,
      cols = col_start:last_col
    )
  }
  openxlsx::addStyle(
    wb, sheet, estilos$titulo,
    rows = r, cols = col_start,
    gridExpand = TRUE, stack = TRUE
  )
  r <- r + 1L

  # ------------------------------------------------------------------------
  # 2. CUERPO DE LA TABLA (cabeceras, datos, estilos)
  # ------------------------------------------------------------------------
  if (n_filas_tabla <= 0 || n_cols_tabla <= 0) {

    # Sin datos: dejamos espacio y seguimos
    row_after_body <- r + 1L

  } else {

    col_names <- names(tabla)

    # ¿Es tabla_conceptos con columnas "Grupo - subcol"?
    hay_grupos <- identical(tipo, "tabla_conceptos") &&
      any(grepl(" - ", col_names[-1], fixed = TRUE))

    if (hay_grupos) {
      # --------------------------------------------------------------------
      # 2.a Cabeceras en dos niveles (grupos)
      # --------------------------------------------------------------------
      row_top  <- r
      row_sub  <- r + 1L
      row_data <- r + 2L

      # Primera columna (rótulo de fila) ocupa dos filas de cabecera
      openxlsx::writeData(
        wb, sheet,
        x = matrix("", nrow = 1),
        startRow = row_sub, startCol = col_start,
        colNames = FALSE
      )
      openxlsx::mergeCells(
        wb, sheet,
        rows = row_top:row_sub,
        cols = col_start
      )

      info <- lapply(col_names[-1], function(nm) {
        parts <- strsplit(nm, " - ", fixed = TRUE)[[1]]
        base  <- parts[1]
        suf   <- if (length(parts) >= 2) parts[2] else ""
        list(base = base, suf = suf)
      })
      bases <- vapply(info, `[[`, character(1), "base")

      used_bases <- character(0)
      for (idx in seq_along(bases)) {
        base <- bases[idx]
        if (base %in% used_bases) next
        used_bases <- c(used_bases, base)

        cols_idx <- which(bases == base) + 1L
        col_from <- col_start + min(cols_idx) - 1L
        col_to   <- col_start + max(cols_idx) - 1L

        # Nombre del grupo (fila superior)
        openxlsx::writeData(
          wb, sheet,
          x = matrix(base, nrow = 1),
          startRow = row_top, startCol = col_from,
          colNames = FALSE
        )
        if (col_from != col_to) {
          openxlsx::mergeCells(
            wb, sheet,
            rows = row_top,
            cols = col_from:col_to
          )
        }

        # Subtítulos (fila inferior de header)
        for (k in cols_idx) {
          suf <- info[[k - 1L]]$suf
          openxlsx::writeData(
            wb, sheet,
            x = matrix(suf, nrow = 1),
            startRow = row_sub,
            startCol = col_start + k - 1L,
            colNames = FALSE
          )
        }
      }

      openxlsx::addStyle(
        wb, sheet, estilos$header,
        rows = row_top:row_sub,
        cols = seq(col_start, length.out = n_cols_tabla),
        gridExpand = TRUE, stack = TRUE
      )

      # Datos
      openxlsx::writeData(
        wb, sheet,
        x = tabla,
        startRow = row_data,
        startCol = col_start,
        colNames = FALSE
      )

      openxlsx::setColWidths(wb, sheet, cols = col_start, widths = 35)
      if (n_cols_tabla > 1) {
        openxlsx::setColWidths(
          wb, sheet,
          cols   = (col_start + 1):(col_start + n_cols_tabla - 1L),
          widths = 10
        )
      }

      body_row_ini <- row_data
      body_row_fin <- row_data + n_filas_tabla - 1L

    } else {
      # --------------------------------------------------------------------
      # 2.b Cabecera simple (frecuencias, tabla_compuesta, etc.)
      # --------------------------------------------------------------------
      row_head <- r
      row_data <- r + 1L

      header_cells <- rep("", n_cols_tabla)

      if (identical(tipo, "freq_multiple") && n_cols_tabla >= 3) {

        header_cells[1] <- ""
        header_cells[2] <- "n"
        header_cells[3] <- "%"

      } else if (n_cols_tabla >= 2) {

        for (j in 2:n_cols_tabla) {
          nm <- col_names[j]
          if (grepl("%", nm, fixed = TRUE) ||
              grepl("pct", nm, ignore.case = TRUE)) {
            header_cells[j] <- "%"
          } else {
            header_cells[j] <- "n"
          }
        }
      }

      openxlsx::writeData(
        wb, sheet,
        x = matrix(header_cells, nrow = 1),
        startRow = row_head,
        startCol = col_start,
        colNames = FALSE
      )

      openxlsx::addStyle(
        wb, sheet, estilos$header,
        rows = row_head,
        cols = seq(col_start, length.out = n_cols_tabla),
        gridExpand = TRUE, stack = TRUE
      )

      # Datos
      openxlsx::writeData(
        wb, sheet,
        x = tabla,
        startRow = row_data,
        startCol = col_start,
        colNames = FALSE
      )

      openxlsx::setColWidths(wb, sheet, cols = col_start, widths = 45)
      if (n_cols_tabla > 1) {
        openxlsx::setColWidths(
          wb, sheet,
          cols   = (col_start + 1):(col_start + n_cols_tabla - 1L),
          widths = 10
        )
      }

      body_row_ini <- row_data
      body_row_fin <- row_data + n_filas_tabla - 1L
    }

    # ----------------------------------------------------------------------
    # 3. Estilos del cuerpo (texto vs numérico vs %)
    # ----------------------------------------------------------------------
    if (n_filas_tabla > 0 && n_cols_tabla > 0) {

      # Primera columna: texto
      openxlsx::addStyle(
        wb, sheet, estilos$body_txt,
        rows = body_row_ini:body_row_fin,
        cols = col_start,
        gridExpand = TRUE, stack = TRUE
      )

      # Resto de columnas: numéricas o porcentajes
      if (n_cols_tabla > 1) {
        for (j in 2:n_cols_tabla) {
          col_abs  <- col_start + j - 1L
          nm       <- col_names[j]
          col_data <- tabla[[j]]

          es_pct <- FALSE
          if (grepl("%", nm, fixed = TRUE) ||
              grepl("pct", nm, ignore.case = TRUE)) {
            es_pct <- TRUE
          }
          if (identical(tipo, "freq_multiple") && j == 3L) {
            es_pct <- TRUE
          }

          if (es_pct) {
            estilo_col <- estilos$body_pct
          } else {
            is_entero <- all(
              is.na(col_data) |
                (abs(col_data - round(col_data)) < 1e-8)
            )
            if (is_entero) {
              estilo_col <- estilos$body_int
            } else {
              estilo_col <- estilos$body_num
            }
          }

          openxlsx::addStyle(
            wb, sheet, estilo_col,
            rows = body_row_ini:body_row_fin,
            cols = col_abs,
            gridExpand = TRUE, stack = TRUE
          )
        }
      }
    }

    row_after_body <- body_row_fin + 1L
  }

  # ------------------------------------------------------------------------
  # 4. FUENTE + BORDE SUPERIOR DE CIERRE
  # ------------------------------------------------------------------------
  if (!is.null(fuente)) {

    row_fuente <- row_after_body

    openxlsx::writeData(
      wb, sheet,
      fuente,
      startRow = row_fuente,
      startCol = col_start,
      colNames = FALSE
    )
    if (last_col > col_start) {
      openxlsx::mergeCells(
        wb, sheet,
        rows = row_fuente,
        cols = col_start:last_col
      )
    }

    if (!is.null(estilos$fuente)) {
      openxlsx::addStyle(
        wb, sheet, estilos$fuente,
        rows = row_fuente,
        cols = col_start,
        gridExpand = TRUE,
        stack = TRUE
      )
    }

    # Línea de cierre justo debajo de la fuente
    if (!is.null(estilos$table_end)) {
      openxlsx::addStyle(
        wb, sheet, estilos$table_end,
        rows = row_fuente,
        cols = col_start:last_col,
        gridExpand = TRUE,
        stack = TRUE
      )
    }

    next_row <- row_fuente + 2L

  } else {

    # Sin texto de fuente: la línea de cierre va justo debajo del cuerpo
    row_cierre <- row_after_body

    if (!is.null(estilos$table_end)) {
      openxlsx::addStyle(
        wb, sheet, estilos$table_end,
        rows = row_cierre,
        cols = col_start:last_col,
        gridExpand = TRUE,
        stack = TRUE
      )
    }

    next_row <- row_cierre + 2L
  }

  invisible(list(next_row = next_row))
}
