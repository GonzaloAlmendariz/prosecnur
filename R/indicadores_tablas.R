#' Generar tablas de indicadores a partir de un YAML
#'
#' Esta función:
#' 1. Lee un archivo de configuración YAML de indicadores (como \code{OPS_EESS_indicadores}).
#' 2. Calcula las tablas para cada indicador según su \code{tipo}
#'    (\code{tabla_conceptos}, \code{tabla_compuesta}, \code{freq_multiple}).
#' 3. Opcionalmente, exporta todas las tablas a un archivo Excel con formato
#'    similar al usado en las tablas de frecuencias/cruces.
#'
#' El YAML debe tener la estructura general:
#' \preformatted{
#' version: 1
#' nombre: "OPS_EESS_indicadores"
#' peso: "peso"
#' indicadores:
#'   - id: IND1
#'     tipo: tabla_conceptos
#'     ...
#'   - id: IND2
#'     tipo: tabla_compuesta
#'     ...
#'   - id: IND3
#'     tipo: freq_multiple
#'     ...
#' }
#'
#' Tipos de indicador soportados:
#' - \code{tabla_conceptos}: filas definidas por \code{filas}/\code{grupos},
#'   columnas definidas por operaciones (\code{suma}, \code{conteo_cond}, \code{proporcion_cond},
#'   \code{media}, \code{mediana}, \code{minimo}, \code{maximo}, \code{proporcion_rel}).
#' - \code{tabla_compuesta}: cada fila es una condición sobre \code{rp_data} y las columnas suelen
#'   ser \code{conteo_cond_fila} y \code{proporcion_sobre_total}.
#' - \code{freq_multiple}: select_multiple ya dummificada en \code{rp_data}, con 0/1 y labels.
#'
#' @param rp_data Data frame (o tibble) con la base ya limpia/recodificada.
#' @param config Ruta al archivo YAML de indicadores (character) o lista ya leída
#'   con \code{yaml::read_yaml()}.
#' @param output_xlsx Ruta del archivo Excel a generar. Si es \code{NULL} (por defecto),
#'   no se exporta a Excel y la función solo devuelve las tablas en una lista.
#' @param hoja Nombre de la hoja donde se escribirán las tablas. Por defecto
#'   \code{\"Indicadores\"}.
#' @param estilos Opcional: lista de estilos de \pkg{openxlsx}. Puede ser la salida
#'   de una función tipo \code{mk_styles_cruces()}. Si es \code{NULL}, se usan
#'   estilos simples por defecto.
#'
#' @return Una lista nombrada, donde cada elemento corresponde a un indicador y
#'   contiene:
#'   \itemize{
#'     \item \code{id}, \code{titulo}, \code{tipo}, \code{grafico}, \code{notas}
#'     \item \code{tabla}: data.frame con los resultados.
#'   }
#'
#' Si \code{output_xlsx} no es \code{NULL}, también genera un archivo Excel.
#'
#' @export
#'
#' @examples
#' \dontrun{
#' # Calcular indicadores sin exportar a Excel
#' res <- indicadores_tablas(rp_data, "OPS_EESS_indicadores.yaml")
#'
#' # Calcular y exportar a Excel
#' indicadores_tablas(
#'   rp_data      = rp_data,
#'   config       = "OPS_EESS_indicadores.yaml",
#'   output_xlsx  = "indicadores_ops_eess.xlsx",
#'   hoja         = "Indicadores",
#'   estilos      = mk_styles_cruces()  # si ya tienes esta función
#' )
#' }
indicadores_tablas <- function(rp_data,
                               config,
                               output_xlsx = NULL,
                               hoja = "Indicadores",
                               estilos = NULL) {

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
      resultados[[ind$id]] <- res_ind
    }
  }

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

    current_row <- 1L
    for (id_ind in names(resultados)) {
      bloque <- resultados[[id_ind]]

      escrito <- .escribir_indicador_excel(
        wb          = wb,
        sheet       = hoja,
        indicador   = bloque,
        row_start   = current_row,
        col_start   = 1L,
        estilos     = estilos
      )

      current_row <- escrito$next_row
    }

    openxlsx::saveWorkbook(wb, file = output_xlsx, overwrite = TRUE)
  }

  invisible(resultados)
}

# Utilidad tipo %||%
`%||%` <- function(x, y) {
  if (is.null(x)) y else x
}

# -----------------------------------------------------------------------------
# PESOS
# -----------------------------------------------------------------------------
.get_pesos <- function(data, peso_var) {
  if (is.null(peso_var) || !nzchar(peso_var) || !peso_var %in% names(data)) {
    rep(1, nrow(data))
  } else {
    w <- data[[peso_var]]
    w[is.na(w)] <- 0
    w
  }
}

# -----------------------------------------------------------------------------
# EVALUAR CONDICIONES A NIVEL DE FILA (tabla_compuesta)
# -----------------------------------------------------------------------------
.eval_condicion_fila <- function(data, condicion) {
  if (is.null(condicion) || !nzchar(condicion)) {
    return(rep(TRUE, nrow(data)))
  }
  env <- list2env(as.list(data), parent = parent.frame())
  res <- eval(parse(text = condicion), envir = env)
  if (!is.logical(res) || length(res) != nrow(data)) {
    stop("La condición '", condicion, "' no devolvió un vector lógico de longitud nrow(data).")
  }
  res
}

# -----------------------------------------------------------------------------
# EVALUAR CONDICIÓN EN UN VECTOR (tabla_conceptos)
# -----------------------------------------------------------------------------
.eval_condicion_vector <- function(x, condicion) {
  # casos especiales
  if (is.null(condicion) || !nzchar(condicion)) {
    return(!is.na(x))
  }
  if (condicion == "no_es_na") {
    return(!is.na(x))
  }

  # si el texto ya contiene 'x', se evalúa tal cual; si no, se asume del tipo "== 1", "> 0", etc.
  expr <- if (grepl("\\bx\\b", condicion)) {
    condicion
  } else {
    paste0("x ", condicion)
  }
  res <- eval(parse(text = expr), envir = list(x = x))
  if (!is.logical(res) || length(res) != length(x)) {
    stop("La condición '", condicion, "' no devolvió un vector lógico de longitud length(x).")
  }
  res
}

# -----------------------------------------------------------------------------
# CALCULAR tabla_compuesta
# -----------------------------------------------------------------------------
.calc_tabla_compuesta <- function(data, ind_cfg, pesos) {

  filas_cfg <- ind_cfg$filas
  cols_cfg  <- ind_cfg$columnas

  # Primero calculamos n de cada fila si se usa conteo_cond_fila
  n_por_fila <- numeric(length(filas_cfg))
  names(n_por_fila) <- vapply(filas_cfg, function(f) f$id, FUN.VALUE = character(1))

  for (i in seq_along(filas_cfg)) {
    f <- filas_cfg[[i]]
    cond <- .eval_condicion_fila(data, f$condicion)
    n_por_fila[i] <- sum(pesos[cond], na.rm = TRUE)
  }

  # Luego construimos la tabla
  res_mat <- matrix(NA_real_,
                    nrow = length(filas_cfg),
                    ncol = length(cols_cfg),
                    dimnames = list(
                      vapply(filas_cfg, function(f) f$label, FUN.VALUE = character(1)),
                      vapply(cols_cfg,  function(c) c$label, FUN.VALUE = character(1))
                    ))

  # vector de ids de fila para referencia_total_fila
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
      res_mat[, j] <- ifelse(denom > 0, 100 * n_por_fila / denom, NA_real_)

    } else {
      warning("Tipo de columna no soportado en tabla_compuesta: ", tipo)
    }
  }

  tabla <- as.data.frame(res_mat, check.names = FALSE)
  tabla <- cbind(
    Fila = vapply(filas_cfg, function(f) f$label, FUN.VALUE = character(1)),
    tabla
  )
  tabla
  list(tabla = tabla)
}

# -----------------------------------------------------------------------------
# CALCULAR freq_multiple
# -----------------------------------------------------------------------------
.calc_freq_multiple <- function(data, ind_cfg, pesos) {

  vars_cfg <- ind_cfg$vars
  usar_labels   <- isTRUE(ind_cfg$usar_labels)
  valor_si      <- ind_cfg$valor_si %||% 1
  incluir_total <- isTRUE(ind_cfg$incluir_total)
  total_label   <- ind_cfg$total_label %||% "Total"

  # Caso A: vars es vector de nombres de variables (p.ej. "p106.1", "p106.2"...)
  if (is.character(vars_cfg)) {
    vars_list <- lapply(vars_cfg, function(v) list(id = v, label = NULL, var = v))
  } else {
    # Caso B: lista con id/label/var (como en IND10)
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

    # Numerador: casos con valor_si
    mask_valid <- !is.na(x)
    mask_si    <- mask_valid & (x == valor_si)
    n_si       <- sum(pesos[mask_si], na.rm = TRUE)
    n_total    <- sum(pesos[mask_valid], na.rm = TRUE)

    n_total_global <- max(n_total_global, n_total)
    pct_si   <- if (n_total > 0) 100 * n_si / n_total else NA_real_

    # Label de fila
    lbl <- v_def$label
    if (usar_labels && is.null(lbl)) {
      lbl <- attr(x, "label") %||% vname
    }
    if (is.null(lbl)) lbl <- vname

    filas[[length(filas) + 1L]] <- data.frame(
      Opcion = lbl,
      n      = n_si,
      `%`    = pct_si,
      stringsAsFactors = FALSE
    )
  }

  tabla <- do.call(rbind, filas)

  if (incluir_total && nrow(tabla) > 0) {
    fila_total <- data.frame(
      Opcion = total_label,
      n      = n_total_global,
      `%`    = 100,
      stringsAsFactors = FALSE
    )
    tabla <- rbind(tabla, fila_total)
  }

  list(tabla = tabla)
}

# -----------------------------------------------------------------------------
# CALCULAR tabla_conceptos
# -----------------------------------------------------------------------------
.calc_tabla_conceptos <- function(data, ind_cfg, pesos) {

  filas_cfg <- ind_cfg$filas
  cols_cfg  <- ind_cfg$columnas

  # Para cada fila, crear una lista de grupos con sus vectores
  filas_labels <- vapply(filas_cfg, function(f) f$label, FUN.VALUE = character(1))

  res_mat <- matrix(NA_real_,
                    nrow = length(filas_cfg),
                    ncol = length(cols_cfg),
                    dimnames = list(
                      filas_labels,
                      vapply(cols_cfg, function(c) c$label, FUN.VALUE = character(1))
                    ))

  for (i in seq_along(filas_cfg)) {
    fila_cfg <- filas_cfg[[i]]

    # Construir lista de grupos evaluados: nombre -> vector (longitud nrow(data))
    grupos <- list()
    if (!is.null(fila_cfg$grupos)) {
      for (g_name in names(fila_cfg$grupos)) {
        g_cfg  <- fila_cfg$grupos[[g_name]]
        vars_g <- g_cfg$vars
        if (length(vars_g) == 1) {
          v <- data[[vars_g]]
        } else {
          v <- rowSums(data[, vars_g, drop = FALSE], na.rm = TRUE)
        }
        grupos[[g_name]] <- v
      }
    }

    # Dentro de la fila, almacenar valores de columnas por id (para proporcion_rel)
    vals_fila <- list()

    for (j in seq_along(cols_cfg)) {
      col_def <- cols_cfg[[j]]
      tipo    <- col_def$tipo
      valor   <- NA_real_

      if (tipo %in% c("suma", "conteo_cond", "proporcion_cond",
                      "media", "mediana", "minimo", "maximo")) {

        # Obtener vector según var = "@grupo"
        var_ref <- col_def$var
        if (is.null(var_ref)) {
          stop("En columnas de tipo ", tipo, " se requiere 'var'.")
        }
        if (!startsWith(var_ref, "@")) {
          stop("Por ahora solo se soportan referencias a grupos (@nombre).")
        }
        g_name <- substring(var_ref, 2L)
        if (is.null(grupos[[g_name]])) {
          stop("No se encontró el grupo '", g_name, "' en la fila ", fila_cfg$id)
        }
        x <- grupos[[g_name]]

        w <- pesos
        mask_base <- !is.na(x)

        # Condición específica
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
          num <- sum(w[mask_base], na.rm = TRUE)
          denom <- sum(w[!is.na(x)], na.rm = TRUE)
          valor <- if (denom > 0) 100 * num / denom else NA_real_

        } else if (tipo == "media") {
          if (!is.null(condicion) && condicion != "no_es_na") {
            # ya está aplicado en mask_base
          }
          num <- sum(x[mask_base] * w[mask_base], na.rm = TRUE)
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
        num <- vals_fila[[num_id]]
        den <- vals_fila[[den_id]]
        valor <- if (den > 0) 100 * num / den else NA_real_

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

# -----------------------------------------------------------------------------
# ESTILOS SENCILLOS POR DEFECTO
# -----------------------------------------------------------------------------
.estilos_sencillos <- function() {
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("Se requiere 'openxlsx' para los estilos.", call. = FALSE)
  }
  list(
    titulo = openxlsx::createStyle(
      fontSize       = 14,
      textDecoration = "bold",
      halign         = "left",
      valign         = "center",
      wrapText       = TRUE
    ),
    notas = openxlsx::createStyle(
      fontSize       = 9,
      italic         = TRUE,
      halign         = "left",
      valign         = "top",
      wrapText       = TRUE
    ),
    header = openxlsx::createStyle(
      fontSize       = 10,
      textDecoration = "bold",
      border         = "Bottom",
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

# -----------------------------------------------------------------------------
# ESCRIBIR UN INDICADOR EN EXCEL
# -----------------------------------------------------------------------------
.escribir_indicador_excel <- function(wb,
                                      sheet,
                                      indicador,
                                      row_start = 1L,
                                      col_start = 1L,
                                      estilos) {

  tabla   <- indicador$tabla
  titulo  <- indicador$titulo %||% indicador$id
  notas   <- indicador$notas %||% character(0)

  n_filas_tabla <- nrow(tabla)
  n_cols_tabla  <- ncol(tabla)

  r <- row_start

  # Título
  openxlsx::writeData(wb, sheet, titulo, startRow = r, startCol = col_start)
  openxlsx::addStyle(
    wb, sheet, estilos$titulo,
    rows = r, cols = col_start,
    gridExpand = TRUE
  )
  r <- r + 2L

  # Tabla
  openxlsx::writeData(
    wb, sheet, x = tabla,
    startRow = r, startCol = col_start,
    headerStyle = estilos$header
  )

  # Estilo de cuerpo
  openxlsx::addStyle(
    wb, sheet, estilos$cuerpo,
    rows = seq(r + 1L, length.out = n_filas_tabla),
    cols = seq(col_start, length.out = n_cols_tabla),
    gridExpand = TRUE
  )

  r <- r + n_filas_tabla + 1L

  # Notas (si existen)
  if (length(notas) > 0) {
    txt_nota <- paste0("Notas: ", paste(notas, collapse = " | "))
    openxlsx::writeData(wb, sheet, txt_nota, startRow = r, startCol = col_start)
    openxlsx::addStyle(wb, sheet, estilos$notas, rows = r, cols = col_start, gridExpand = TRUE)
    r <- r + 2L
  } else {
    r <- r + 1L
  }

  # Fila en blanco entre indicadores
  r <- r + 1L

  list(next_row = r)
}
