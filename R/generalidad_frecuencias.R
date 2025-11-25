# ============================
# Helpers internos (no exportar)
# ============================

#' @noRd
.peso_vec <- function(data){
  if (!("peso" %in% names(data))) return(rep(1, nrow(data)))
  w <- suppressWarnings(as.numeric(data$peso))
  w[!is.finite(w) | is.na(w)] <- 0
  w
}

#' @noRd
.auto_row_height <- function(text, chars_per_line = 70, base = 24, per_line = 16){
  if (length(text) == 0 || is.na(text)) return(base)
  txt <- gsub("\\r?\\n", " ", as.character(text))
  lines <- max(1, ceiling(nchar(txt) / chars_per_line))
  base + (lines - 1) * per_line
}

#' @noRd
.move_ns_pref_last <- function(tab){
  if (!nrow(tab) || !"Opciones" %in% names(tab)) return(tab)
  idx <- which(trimws(tab$Opciones) == "No sé / Prefiero no decir")
  if (length(idx) == 0) return(tab)
  dplyr::bind_rows(tab[-idx, , drop = FALSE],
                   tab[ idx, , drop = FALSE])
}

#' @noRd
.map_from_attr_labels <- function(tab, var, df){
  if (is.null(df) || !(var %in% names(df))) return(tab)

  lab_attr <- attr(df[[var]], "labels", exact = TRUE)
  if (is.null(lab_attr) || length(lab_attr) == 0) return(tab)

  codes_vec  <- as.character(names(lab_attr))
  labels_vec <- as.character(unname(lab_attr))

  if (!"Opciones" %in% names(tab)) return(tab)

  is_total <- tab$Opciones %in% c("Total", "")
  body  <- if (any(is_total)) tab[!is_total, , drop = FALSE] else tab
  total <- if (any(is_total)) tab[ is_total, , drop = FALSE] else NULL

  idx <- match(as.character(body$Opciones), codes_vec)
  body$Opciones <- ifelse(!is.na(idx), labels_vec[idx], body$Opciones)

  if (!is.null(total) && nrow(total)) dplyr::bind_rows(body, total) else body
}

#' @noRd
.map_to_labels <- function(tab, var, orders_list){
  if (is.null(orders_list)) return(tab)
  if (!"Opciones" %in% names(tab)) return(tab)

  is_total <- tab$Opciones == "Total"
  body <- if (any(is_total)) tab[!is_total, , drop = FALSE] else tab
  total <- if (any(is_total)) tab[ is_total, , drop = FALSE] else NULL
  if (!nrow(body)) return(tab)

  ord_lbl <- tryCatch(orders_list[[var]]$labels, error = function(e) NULL)
  ord_nam <- tryCatch(orders_list[[var]]$names,  error = function(e) NULL)

  if (!is.null(ord_nam) && !is.null(ord_lbl)) {
    idx_code <- match(body$Opciones, ord_nam)
    body$Opciones <- ifelse(!is.na(idx_code), ord_lbl[idx_code], body$Opciones)
  }

  if (!is.null(total) && nrow(total)) dplyr::bind_rows(body, total) else body
}

#' @noRd
.reordenar_por_instrumento <- function(tab, var, orders_list){
  if (is.null(orders_list) || !(var %in% names(orders_list))) return(tab)
  if (!all(c("Opciones","n","pct") %in% names(tab))) return(tab)

  is_total <- tab$Opciones == "Total"
  body <- if (any(is_total)) tab[!is_total, , drop = FALSE] else tab
  total <- if (any(is_total)) tab[ is_total, , drop = FALSE] else NULL
  if (!nrow(body)) return(tab)

  ord_lbl <- tryCatch(orders_list[[var]]$labels, error = function(e) NULL)
  ord_nam <- tryCatch(orders_list[[var]]$names,  error = function(e) NULL)

  if (!is.null(ord_lbl)) {
    body <- dplyr::mutate(body, .orden_aux = match(Opciones, ord_lbl))
  } else {
    body$.orden_aux <- NA_integer_
  }

  if (all(is.na(body$.orden_aux)) && !is.null(ord_nam)) {
    body <- dplyr::mutate(body, .orden_aux = match(Opciones, ord_nam))
  }

  body <- body |>
    dplyr::mutate(
      .orden_aux = ifelse(
        is.na(.orden_aux),
        max(.orden_aux, na.rm = TRUE) + dplyr::row_number(),
        .orden_aux
      )
    ) |>
    dplyr::arrange(.orden_aux) |>
    dplyr::select(-.orden_aux)

  if (!is.null(total) && nrow(total)) dplyr::bind_rows(body, total) else body
}

#' @noRd
titulo_var <- function(var, dic_vars = NULL, labels_override = NULL,
                       orders_list = NULL, df = NULL) {
  if (!is.null(labels_override) && var %in% names(labels_override)) {
    return(as.character(labels_override[[var]]))
  }

  if (!is.null(df) && var %in% names(df)) {
    vl <- attr(df[[var]], "label", exact = TRUE)
    if (!is.null(vl) && nzchar(as.character(vl))) return(as.character(vl))
  }

  if (!is.null(orders_list) && var %in% names(orders_list)) {
    ordv <- orders_list[[var]]
    if (!is.null(ordv$label) && nzchar(as.character(ordv$label))) {
      return(as.character(ordv$label))
    }
    lab_attr <- tryCatch(attr(ordv$labels, "label"), error = function(e) NULL)
    if (!is.null(lab_attr) && nzchar(as.character(lab_attr))) {
      return(as.character(lab_attr))
    }
  }

  if (!is.null(dic_vars) && all(c("name","label") %in% names(dic_vars))) {
    lab <- dic_vars$label[dic_vars$name == var]
    if (length(lab) && !all(is.na(lab))) return(as.character(lab[1]))
  }

  return(as.character(var))
}

#' @noRd
tipo_pregunta_spss <- function(var, survey, sm_vars_force = NULL) {
  if (!is.null(sm_vars_force) && var %in% sm_vars_force) return("sm")

  if (!is.null(survey) && all(c("type","name") %in% names(survey))) {
    tps <- unique(stats::na.omit(survey$type[survey$name == var]))
    if (length(tps)) {
      if (any(grepl("^select_multiple(\\s|$)", tps))) return("sm")
      if (any(grepl("^select_one(\\s|$)", tps)))      return("so")
    }
  }
  if (!is.null(survey) &&
      any(grepl(paste0("^", stringr::fixed(var), "/"), names(survey)))) {
    return("sm")
  }
  "so_or_open"
}

#' @noRd
split_sm_tokens <- function(x) {
  stringr::str_split(x, "\\s*[;]+\\s*", simplify = FALSE)
}

# ============================
# freq_table_spss (pública)
# ============================

#' Tabla de frecuencias ponderadas para una variable
#'
#' Calcula una tabla de frecuencias (n y porcentaje) para una variable de la
#' base de reporte. Soporta variables de tipo elección múltiple (`select_multiple`)
#' codificadas como:
#' \itemize{
#'   \item Variable madre con códigos separados por `";"` (p. ej. `"1;3;4"`).
#'   \item Dummies derivadas (`var/cod`) recodificadas a 0/1.
#' }
#'
#' La función utiliza, cuando están disponibles:
#' \itemize{
#'   \item Atributos `labels` de la variable en `data` (para mapear códigos a
#'         etiquetas).
#'   \item `orders_list` (si se proporciona) para ordenar las categorías según
#'         el instrumento.
#'   \item Una variable de peso llamada `peso` en `data`. Si no existe, se
#'         asume peso 1 para todos los casos.
#' }
#'
#' @param data Data frame o tibble con la base de datos.
#' @param var Nombre de la variable (como cadena) para la que se desea la tabla.
#' @param survey Tibble con metadatos del instrumento (hoja `survey`), que
#'   debe contener al menos las columnas `name` y `type`. Se utiliza para
#'   diferenciar `select_one` y `select_multiple`. Puede ser `NULL`.
#' @param sm_vars_force Vector opcional de nombres de variables que deben tratarse
#'   como `select_multiple` aunque el instrumento no las marque como tales.
#' @param orders_list Lista opcional con información de orden de categorías
#'   por variable (por ejemplo, `instrumento$orders_list`).
#'
#' @return Un tibble con las columnas:
#' \describe{
#'   \item{Opciones}{Código o etiqueta de la categoría.}
#'   \item{n}{Frecuencia ponderada.}
#'   \item{pct}{Porcentaje relativo (0–1).}
#' }
#'
#' @export
freq_table_spss <- function(data, var, survey = NULL, sm_vars_force = NULL,
                            orders_list = NULL) {

  stopifnot(var %in% names(data))
  tipo <- tipo_pregunta_spss(var, survey, sm_vars_force)
  w <- .peso_vec(data)

  if (tipo == "sm") {

    # Caso 1: madre "pegada" (ej. "1;3;4")
    if (is.character(data[[var]]) || is.factor(data[[var]])) {
      vec <- as.character(data[[var]])
      df_long <- tibble::tibble(id = seq_len(nrow(data)), valor = vec) |>
        dplyr::filter(!is.na(valor) & nzchar(valor) & valor != "NA") |>
        dplyr::mutate(tokens = split_sm_tokens(valor)) |>
        dplyr::select(-valor) |>
        tidyr::unnest_longer(tokens, values_to = "op") |>
        dplyr::mutate(op = trimws(op)) |>
        dplyr::filter(nzchar(op)) |>
        dplyr::distinct(id, op)

      if (!nrow(df_long)) {
        return(tibble::tibble(Opciones = character(), n = numeric(), pct = numeric()))
      }

      ids_con_marca <- sort(unique(df_long$id))
      denom <- sum(w[ids_con_marca], na.rm = TRUE)

      tab <- df_long |>
        dplyr::left_join(
          tibble::tibble(id = seq_len(nrow(data)), peso = w),
          by = "id"
        ) |>
        dplyr::group_by(op) |>
        dplyr::summarise(n = sum(peso, na.rm = TRUE), .groups = "drop") |>
        dplyr::arrange(dplyr::desc(n)) |>
        dplyr::transmute(
          Opciones = op,
          n        = as.numeric(n),
          pct      = if (denom > 0) n/denom else NA_real_
        )

      tab <- .map_from_attr_labels(tab, var, data)
      tab <- .map_to_labels(tab, var, orders_list)
      tab <- .reordenar_por_instrumento(tab, var, orders_list)
      tab <- .move_ns_pref_last(tab)

      total_row <- tibble::tibble(
        Opciones = "Total",
        n        = as.numeric(denom),
        pct      = 1
      )
      return(dplyr::bind_rows(tab, total_row))

    } else {
      # Caso 2: dummies var/cod ya en columnas
      subvars <- names(data)[grepl(paste0("^", stringr::fixed(var), "/"), names(data))]
      if (!length(subvars)) {
        return(tibble::tibble(Opciones = character(), n = numeric(), pct = numeric()))
      }
      mat <- as.data.frame(data[, subvars, drop = FALSE])
      mat[] <- lapply(mat, function(v) suppressWarnings(as.numeric(as.character(v))))

      has_any <- rowSums(mat == 1, na.rm = TRUE) > 0
      denom <- sum(w[has_any], na.rm = TRUE)

      n_w <- vapply(subvars, function(sv){
        v <- suppressWarnings(as.numeric(as.character(mat[[sv]])))
        sum(w[v == 1 & !is.na(v)], na.rm = TRUE)
      }, numeric(1))

      tab <- tibble::tibble(subvar = subvars, n = as.numeric(n_w)) |>
        dplyr::mutate(
          Opciones = sub(
            paste0("^", gsub("([\\W])", "\\\\\\1", var), "/"),
            "",
            subvar
          )
        ) |>
        dplyr::arrange(dplyr::desc(n)) |>
        dplyr::transmute(
          Opciones,
          n,
          pct = if (denom > 0) n/denom else NA_real_
        )

      tab <- .map_to_labels(tab, var, orders_list)
      tab <- .reordenar_por_instrumento(tab, var, orders_list)
      tab <- .move_ns_pref_last(tab)

      total_row <- tibble::tibble(
        Opciones = "Total",
        n        = as.numeric(denom),
        pct      = 1
      )
      return(dplyr::bind_rows(tab, total_row))
    }

  } else {

    col <- var
    tib <- data |>
      dplyr::transmute(.op = as.character(.data[[col]]), peso = w) |>
      dplyr::filter(!is.na(.op) & nzchar(.op) & .op != "NA")

    if (!nrow(tib)) {
      return(tibble::tibble(Opciones = character(), n = numeric(), pct = numeric()))
    }

    denom <- sum(tib$peso, na.rm = TRUE)

    tab <- tib |>
      dplyr::group_by(.op) |>
      dplyr::summarise(n = sum(peso, na.rm = TRUE), .groups = "drop") |>
      dplyr::arrange(dplyr::desc(n)) |>
      dplyr::mutate(pct = if (denom > 0) n/denom else NA_real_) |>
      dplyr::rename(Opciones = .op)

    tab <- .map_from_attr_labels(tab, var, data)
    tab <- .map_to_labels(tab, var, orders_list)
    tab <- .reordenar_por_instrumento(tab, var, orders_list)
    tab <- .move_ns_pref_last(tab)

    total_row <- tibble::tibble(
      Opciones = "Total",
      n        = sum(tab$n, na.rm = TRUE),
      pct      = 1
    )
    dplyr::bind_rows(tab, total_row)
  }
}

# ============================
# Estilos y escritura en Excel
# ============================

#' @noRd
mk_styles_spss <- function() {
  list(
    sec_title = openxlsx::createStyle(
      fontSize     = 18,
      textDecoration = NULL,
      halign       = "center",
      valign       = "center",
      wrapText     = TRUE,
      fgFill       = "#FFFFFF",
      fontColour   = "#000000",
      fontName     = "Arial"
    ),
    q_title = openxlsx::createStyle(
      fontSize     = 11,
      textDecoration = "italic",
      halign       = "left",
      valign       = "center",
      wrapText     = TRUE,
      fgFill       = "#FFFFFF",
      fontColour   = "#000000",
      fontName     = "Arial"
    ),
    header = openxlsx::createStyle(
      fontSize     = 10,
      textDecoration = NULL,
      border       = c("top", "bottom"),
      borderStyle  = "thin",
      borderColour = "#000000",
      halign       = "center",
      valign       = "center",
      fgFill       = "#FFFFFF",
      fontName     = "Arial"
    ),
    body_txt = openxlsx::createStyle(
      fontSize     = 10,
      textDecoration = NULL,
      border       = c(),
      halign       = "left",
      valign       = "center",
      fgFill       = "#FFFFFF",
      fontName     = "Arial",
      wrapText     = TRUE
    ),
    body_int = openxlsx::createStyle(
      fontSize     = 10,
      textDecoration = NULL,
      numFmt       = "#,##0",
      border       = c(),
      halign       = "right",
      valign       = "center",
      fgFill       = "#FFFFFF",
      fontName     = "Arial"
    ),
    body_pct = openxlsx::createStyle(
      fontSize     = 10,
      textDecoration = NULL,
      numFmt       = "0.0%",
      border       = c(),
      halign       = "right",
      valign       = "center",
      fgFill       = "#FFFFFF",
      fontName     = "Arial"
    ),
    total_row = openxlsx::createStyle(
      fontSize     = 10,
      textDecoration = NULL,
      numFmt       = "#,##0",
      halign       = "right",
      valign       = "center",
      fgFill       = "#FFFFFF",
      fontName     = "Arial"
    ),
    total_label = openxlsx::createStyle(
      fontSize     = 10,
      textDecoration = NULL,
      halign       = "left",
      valign       = "center",
      fgFill       = "#FFFFFF",
      fontName     = "Arial"
    ),
    table_end = openxlsx::createStyle(
      border       = c("bottom"),
      borderStyle  = "thin",
      borderColour = "#000000"
    )
  )
}

#' @noRd
write_one_freq <- function(wb, sheet, data, var, dic_vars,
                           survey = NULL, sm_vars_force = NULL,
                           labels_override = NULL,
                           start_row = 1, start_col = 1,
                           fuente = "Pulso PUCP",
                           orders_list = NULL) {
  st <- mk_styles_spss()
  fila <- start_row

  label_q <- titulo_var(
    var,
    dic_vars,
    labels_override,
    orders_list = orders_list,
    df         = data
  )

  openxlsx::writeData(wb, sheet, label_q,
                      startRow = fila, startCol = start_col, colNames = FALSE)
  openxlsx::mergeCells(wb, sheet,
                       cols = start_col:(start_col + 2),
                       rows = fila:fila)
  openxlsx::addStyle(wb, sheet, st$q_title,
                     rows = fila, cols = start_col,
                     gridExpand = TRUE, stack = TRUE)
  openxlsx::setRowHeights(
    wb, sheet, rows = fila,
    heights = .auto_row_height(label_q, chars_per_line = 70,
                               base = 24, per_line = 16)
  )
  fila <- fila + 1

  header_vec <- c("", "n", "%")
  openxlsx::writeData(wb, sheet, t(header_vec),
                      startRow = fila, startCol = start_col, colNames = FALSE)
  openxlsx::addStyle(wb, sheet, st$header,
                     rows = fila,
                     cols = start_col:(start_col + 2),
                     gridExpand = TRUE, stack = TRUE)
  fila <- fila + 1

  tab <- freq_table_spss(
    data,
    var,
    survey        = survey,
    sm_vars_force = sm_vars_force,
    orders_list   = orders_list
  )

  if (!nrow(tab)) {
    openxlsx::writeData(wb, sheet, "Sin datos",
                        startRow = fila, startCol = start_col)
    return(fila + 2)
  }

  is_total <- tab$Opciones == "Total"
  body_rows <- if (any(is_total)) tab[!is_total, , drop = FALSE] else tab
  total_row <- if (any(is_total)) tab[ is_total, , drop = FALSE] else NULL

  if (nrow(body_rows)) {
    openxlsx::writeData(wb, sheet, body_rows,
                        startRow = fila, startCol = start_col,
                        colNames = FALSE)
    r_ini <- fila; r_fin <- fila + nrow(body_rows) - 1

    openxlsx::addStyle(wb, sheet, st$body_txt,
                       rows = r_ini:r_fin, cols = start_col,
                       gridExpand = TRUE)
    openxlsx::addStyle(wb, sheet, st$body_int,
                       rows = r_ini:r_fin, cols = start_col + 1,
                       gridExpand = TRUE)
    openxlsx::addStyle(wb, sheet, st$body_pct,
                       rows = r_ini:r_fin, cols = start_col + 2,
                       gridExpand = TRUE)

    fila <- r_fin + 1
  }

  if (!is.null(total_row) && nrow(total_row)) {
    openxlsx::writeData(wb, sheet, total_row,
                        startRow = fila, startCol = start_col,
                        colNames = FALSE)
    openxlsx::addStyle(wb, sheet, st$total_label,
                       rows = fila, cols = start_col,
                       gridExpand = TRUE)
    openxlsx::addStyle(wb, sheet, st$total_row,
                       rows = fila, cols = start_col + 1,
                       gridExpand = TRUE)
    openxlsx::addStyle(wb, sheet, st$total_row,
                       rows = fila, cols = start_col + 2,
                       gridExpand = TRUE)

    openxlsx::addStyle(wb, sheet, st$table_end,
                       rows = fila,
                       cols = start_col:(start_col + 2),
                       gridExpand = TRUE, stack = TRUE)

    fila <- fila + 1
  } else {
    openxlsx::addStyle(wb, sheet, st$table_end,
                       rows = max(start_row + 1, fila - 1),
                       cols = start_col:(start_col + 2),
                       gridExpand = TRUE, stack = TRUE)
  }

  openxlsx::writeData(wb, sheet, paste0("Fuente: ", fuente),
                      startRow = fila, startCol = start_col, colNames = FALSE)
  openxlsx::addStyle(
    wb, sheet,
    openxlsx::createStyle(
      fontSize   = 9,
      fontColour = "#666666",
      halign     = "left",
      fontName   = "Arial"
    ),
    rows = fila, cols = start_col, gridExpand = TRUE
  )
  openxlsx::setColWidths(wb, sheet, cols = start_col,     widths = 55)
  openxlsx::setColWidths(wb, sheet, cols = start_col + 1, widths = 14)
  openxlsx::setColWidths(wb, sheet, cols = start_col + 2, widths = 14)

  return(fila + 2)
}

# ============================
# exportar_frecuencias_spss (pública)
# ============================

#' Exportar tablas de frecuencias a Excel por secciones
#'
#' Función de nivel intermedio que recibe una base de datos, un diccionario
#' de variables y una lista de secciones, y genera un archivo Excel con tablas
#' de frecuencias simples (n y %) para cada variable indicada.
#'
#' Suele llamarse desde [reporte_frecuencias()], que se encarga de construir
#' `dic_vars` y `SECCIONES` a partir del instrumento.
#'
#' @param data Data frame o tibble con la base de datos.
#' @param dic_vars Tibble con al menos las columnas `name` y `label`.
#' @param SECCIONES Lista nombrada donde cada elemento es un vector de nombres
#'   de variables a incluir en esa sección.
#' @param labels_override Vector nombrado opcional para sobrescribir etiquetas
#'   de algunas variables en los títulos de tabla.
#' @param path_xlsx Ruta del archivo Excel a generar.
#' @param orden Criterio de ordenamiento interno de las categorías: `"desc"`,
#'   `"asc"` o `"original"`.
#' @param sm_vars_force Vector opcional de variables que deben tratarse como
#'   `select_multiple` aunque el instrumento no las marque como tales.
#' @param fuente Texto que se mostrará como fuente al pie de cada tabla.
#' @param orders_list Lista opcional con información de orden de categorías
#'   por variable.
#' @param survey Tibble con la hoja `survey` del instrumento (para detectar
#'   `select_one` y `select_multiple`).
#'
#' @return Invisiblemente, la ruta normalizada del archivo Excel generado.
#'
#' @export
exportar_frecuencias_spss <- function(
    data,
    dic_vars,
    SECCIONES,
    labels_override = NULL,
    path_xlsx = "frecuencias_spss.xlsx",
    orden = c("desc","asc","original"),
    sm_vars_force = NULL,
    fuente = "Pulso PUCP",
    orders_list = NULL,
    survey = NULL
){
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("El paquete 'openxlsx' es necesario para `exportar_frecuencias_spss()`. ",
         "Instálalo con install.packages('openxlsx').", call. = FALSE)
  }

  orden <- match.arg(orden)
  wb <- openxlsx::createWorkbook()
  sheet <- "Frecuencias"
  openxlsx::addWorksheet(wb, sheet)
  st <- mk_styles_spss()

  fila <- 1L

  for (sec in names(SECCIONES)) {
    vars_sec <- SECCIONES[[sec]]
    vars_sec <- vars_sec[vars_sec %in% names(data)]
    if (!length(vars_sec)) next

    openxlsx::writeData(wb, sheet, toupper(sec),
                        startRow = fila, startCol = 1)
    openxlsx::mergeCells(wb, sheet, cols = 1:3, rows = fila:fila)
    openxlsx::addStyle(wb, sheet, st$sec_title,
                       rows = fila, cols = 1,
                       gridExpand = TRUE, stack = TRUE)
    openxlsx::setRowHeights(
      wb, sheet, rows = fila,
      heights = .auto_row_height(toupper(sec),
                                 chars_per_line = 70,
                                 base = 28, per_line = 18)
    )
    fila <- fila + 2

    for (v in vars_sec) {

      tab <- freq_table_spss(
        data,
        v,
        survey        = survey,
        sm_vars_force = sm_vars_force,
        orders_list   = orders_list
      )

      if (nrow(tab)) {
        is_total <- tab$Opciones == "Total"
        body <- tab[!is_total, , drop = FALSE]
        total <- tab[ is_total, , drop = FALSE]

        if (orden %in% c("asc","desc") && nrow(body)) {
          body <- dplyr::arrange(body,
                                 if (orden == "asc") n else dplyr::desc(n))
          tab <- dplyr::bind_rows(body, total)
        } else {
          tab <- dplyr::bind_rows(body, total)
        }
      }

      fila <- write_one_freq(
        wb, sheet,
        data  = data,
        var   = v,
        dic_vars = dic_vars,
        survey   = survey,
        sm_vars_force = sm_vars_force,
        labels_override = labels_override,
        start_row = fila,
        start_col = 1,
        fuente = fuente,
        orders_list = orders_list
      )
    }

    fila <- fila + 1
  }

  openxlsx::saveWorkbook(wb, path_xlsx, overwrite = TRUE)
  message("Frecuencias exportadas a: ", normalizePath(path_xlsx, winslash = "/"))
  invisible(normalizePath(path_xlsx, winslash = "/"))
}

# ============================
# reporte_frecuencias (wrapper alto nivel)
# ============================

#' Generar reporte de frecuencias en Excel a partir de una base de reporte
#'
#' `reporte_frecuencias()` toma una base ya adaptada para reporte (típicamente
#' el resultado de [reporte_data()]) y genera un archivo Excel con tablas de
#' frecuencias simples (n y %), organizadas por secciones.
#'
#' Utiliza la información del instrumento (`survey`, `orders_list`) para
#' construir etiquetas de variables y ordenar categorías. Internamente llama
#' a [exportar_frecuencias_spss()].
#'
#' @param data Data frame o tibble, idealmente el objeto devuelto por
#'   [reporte_data()], que contiene los atributos `label` y `labels`.
#' @param instrumento Objeto devuelto por [reporte_instrumento()]. Si es
#'   `NULL`, se intentará recuperarlo desde `attr(data, "instrumento_reporte")`.
#' @param secciones Lista nombrada que define qué variables se incluyen en cada
#'   sección. Si es `NULL`, se intentará inferirla desde una columna
#'   `section` o `seccion` del `survey`.
#' @param path_xlsx Ruta del archivo Excel a generar.
#' @param orden `"desc"`, `"asc"` o `"original"` para el orden de las
#'   categorías dentro de cada tabla.
#' @param sm_vars_force Vector opcional de variables que deben tratarse como
#'   `select_multiple` aunque el instrumento no las marque como tales.
#' @param fuente Texto de fuente que se mostrará al pie de cada tabla.
#'
#' @return Invisiblemente, la ruta normalizada del archivo Excel generado.
#'
#' @export
reporte_frecuencias <- function(data,
                                instrumento = NULL,
                                secciones   = NULL,
                                path_xlsx   = "frecuencias_spss.xlsx",
                                orden       = c("desc", "asc", "original"),
                                sm_vars_force = NULL,
                                fuente      = "Pulso PUCP") {

  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("El paquete 'openxlsx' es necesario para `reporte_frecuencias()`. ",
         "Instálalo con install.packages('openxlsx').", call. = FALSE)
  }

  if (!is.data.frame(data)) {
    stop("`data` debe ser un data.frame o tibble.", call. = FALSE)
  }

  if (is.null(instrumento)) {
    instrumento <- attr(data, "instrumento_reporte", exact = TRUE)
    if (is.null(instrumento)) {
      stop("No se proporcionó `instrumento` y `data` no tiene atributo ",
           "`instrumento_reporte`.", call. = FALSE)
    }
  }

  survey <- instrumento$survey
  if (is.null(survey) || !all(c("name", "label") %in% names(survey))) {
    stop("El `instrumento` no contiene un `survey` con columnas `name` y `label`.",
         call. = FALSE)
  }

  dic_vars <- survey |>
    dplyr::filter(!is.na(.data$name), .data$name != "") |>
    dplyr::select(name, label) |>
    dplyr::mutate(label = trimws(as.character(.data$label))) |>
    dplyr::distinct(name, .keep_all = TRUE)

  orders_list <- if (!is.null(instrumento$orders_list)) instrumento$orders_list else NULL

  orden <- match.arg(orden)

  # inferir secciones si no se pasan
  if (is.null(secciones)) {
    seccion_col <- NULL
    if ("section" %in% names(survey)) {
      seccion_col <- "section"
    } else if ("seccion" %in% names(survey)) {
      seccion_col <- "seccion"
    }

    if (!is.null(seccion_col)) {
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
    } else {
      stop("No se especificaron `secciones` y el `survey` no tiene columna ",
           "`section` ni `seccion`.", call. = FALSE)
    }
  }

  SECCIONES <- lapply(secciones, function(vars) {
    intersect(vars, names(data))
  })
  SECCIONES <- SECCIONES[vapply(SECCIONES, length, integer(1)) > 0L]

  if (length(SECCIONES) == 0L) {
    stop("Después de intersectar con `data`, ninguna sección tiene variables. ",
         "Revisar los nombres en `secciones`.", call. = FALSE)
  }

  exportar_frecuencias_spss(
    data           = data,
    dic_vars       = dic_vars,
    SECCIONES      = SECCIONES,
    labels_override = NULL,
    path_xlsx      = path_xlsx,
    orden          = orden,
    sm_vars_force  = sm_vars_force,
    fuente         = fuente,
    orders_list    = orders_list,
    survey         = survey
  )

  invisible(normalizePath(path_xlsx, winslash = "/"))
}
