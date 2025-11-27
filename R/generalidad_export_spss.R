#' Exportar una base de reporte a formato SPSS (.sav)
#'
#' `reporte_spss()` toma una base ya adaptada para reporte (típicamente el
#' resultado de [reporte_data()]) y la convierte a un objeto compatible con
#' SPSS, aplicando:
#' \itemize{
#'   \item Conversión de variables con etiquetas de valor (`attr(, "labels")`)
#'         a objetos `haven::labelled_spss()`, usando códigos numéricos y
#'         manteniendo los `label` de variable y el nivel de medición
#'         (`attr(, "measure")`).
#'   \item Conversión de variables de fecha, hora y fecha-hora según los
#'         metadatos del instrumento (`vars_fecha`, `vars_hora`, `vars_datetime`).
#'   \item Renombrado de `_uuid` a `uuid` si fuese necesario.
#'   \item Escritura a disco en formato `.sav` mediante [haven::write_sav()].
#' }
#'
#' La lógica asume que la base de entrada ya pasó por las etapas de:
#' \enumerate{
#'   \item Evaluación de consistencia.
#'   \item Recodificación/adaptación de instrumento y data.
#'   \item Preparación para reporte con [reporte_data()], donde se asignan
#'         `label`, `labels` y `measure`, y se identifican variables de fecha
#'         y hora.
#' }
#'
#' Además, la función realiza una comprobación básica de la aplicación de
#' etiquetas: cuenta cuántas variables con `labels` fueron convertidas a
#' `labelled_spss`, e identifica posibles problemas al convertir el contenido
#' a numérico (por ejemplo, si todo quedó en `NA`).
#'
#' @param data Un `data.frame` o `tibble`, preferentemente el objeto devuelto
#'   por [reporte_data()] (clase `"prosecnur_reporte_tbl"`).
#' @param path_sav Ruta del archivo `.sav` a generar. Debe incluir la extensión,
#'   por ejemplo `"estudio_final.sav"`.
#' @param compress Lógico; se pasa a [haven::write_sav()]. Por defecto `TRUE`.
#' @param ... Argumentos adicionales que se pasan directamente a
#'   [haven::write_sav()].
#'
#' @return Invisiblemente, el `data.frame` ya transformado (con clases
#'   `labelled_spss`, tipos de fecha/hora ajustados, etc.). Como efecto
#'   secundario, se escribe el archivo `.sav` en `path_sav` y se imprime en
#'   consola un breve resumen sobre la aplicación de labels.
#'
#' @examples
#' \dontrun{
#'   rp_inst <- reporte_instrumento("instrumento.xlsx", ...)
#'   rp_data <- reporte_data(data_cruda_adaptada, rp_inst)
#'
#'   reporte_spss(rp_data, path_sav = "estudio_2025.sav")
#' }
#'
#' @export
reporte_spss <- function(data,
                         path_sav,
                         compress = TRUE,
                         ...) {

  if (!requireNamespace("haven", quietly = TRUE)) {
    stop("El paquete 'haven' es necesario para `reporte_spss()`. ",
         "Instálalo con install.packages('haven').", call. = FALSE)
  }
  if (!requireNamespace("hms", quietly = TRUE)) {
    stop("El paquete 'hms' es necesario para `reporte_spss()`. ",
         "Instálalo con install.packages('hms').", call. = FALSE)
  }

  if (missing(path_sav) || !nzchar(path_sav)) {
    stop("Debe especificarse `path_sav` (ruta al archivo .sav a generar).",
         call. = FALSE)
  }

  if (!is.data.frame(data)) {
    stop("`data` debe ser un data.frame o tibble.", call. = FALSE)
  }

  # Trabajar sobre una copia local
  df <- data

  # ---------------------------------------------------------------------------
  # 1) Recuperar metadatos del instrumento desde atributos (si existen)
  # ---------------------------------------------------------------------------
  instr         <- attr(df, "instrumento_reporte", exact = TRUE)
  vars_fecha    <- attr(df, "vars_fecha",    exact = TRUE)
  vars_hora     <- attr(df, "vars_hora",     exact = TRUE)
  vars_datetime <- attr(df, "vars_datetime", exact = TRUE)

  `%||%` <- function(x, y) if (!is.null(x)) x else y

  if (is.null(vars_fecha) && !is.null(instr)) {
    vars_fecha <- instr$vars_fecha
  }
  if (is.null(vars_hora) && !is.null(instr)) {
    vars_hora <- instr$vars_hora
  }
  if (is.null(vars_datetime) && !is.null(instr)) {
    vars_datetime <- instr$vars_datetime
  }

  # ---------------------------------------------------------------------------
  # 2) Convertir tipos especiales (fecha, hora, fecha-hora)
  # ---------------------------------------------------------------------------
  # Fechas
  if (!is.null(vars_fecha) && length(vars_fecha) > 0L) {
    vars_fecha <- intersect(vars_fecha, names(df))
    for (v in vars_fecha) {
      if (!inherits(df[[v]], "Date")) {
        df[[v]] <- as.Date(df[[v]])
      }
    }
  }

  # Horas
  if (!is.null(vars_hora) && length(vars_hora) > 0L) {
    vars_hora <- intersect(vars_hora, names(df))
    for (v in vars_hora) {
      if (!inherits(df[[v]], "hms")) {
        df[[v]] <- hms::as_hms(df[[v]])
      }
    }
  }

  # Fecha-hora (datetime)
  if (!is.null(vars_datetime) && length(vars_datetime) > 0L) {
    vars_datetime <- intersect(vars_datetime, names(df))
    for (v in vars_datetime) {
      if (!inherits(df[[v]], "POSIXct")) {
        df[[v]] <- as.POSIXct(df[[v]])
      }
    }
  }

  # ---------------------------------------------------------------------------
  # 3) Renombrar _uuid a uuid si corresponde
  # ---------------------------------------------------------------------------
  if ("_uuid" %in% names(df) && !"uuid" %in% names(df)) {
    names(df)[names(df) == "_uuid"] <- "uuid"
  }

  # ---------------------------------------------------------------------------
  # 4) Convertir variables con value-labels a haven::labelled_spss
  #    + auditoría básica de conversión
  # ---------------------------------------------------------------------------
  vars_with_labels <- names(df)[vapply(
    df,
    function(x) !is.null(attr(x, "labels", exact = TRUE)),
    logical(1)
  )]

  info_labels <- list()

  if (length(vars_with_labels) > 0L) {
    for (v in vars_with_labels) {
      x     <- df[[v]]
      labs  <- attr(x, "labels",  exact = TRUE)
      v_lab <- attr(x, "label",   exact = TRUE)
      meas  <- attr(x, "measure", exact = TRUE)

      if (is.null(labs) || length(labs) == 0L) next

      # En este flujo: names(labs) = códigos (character), labs[] = etiquetas (character)
      codigos <- suppressWarnings(as.numeric(names(labs)))
      textos  <- as.character(unname(labs))

      ok      <- !is.na(codigos)
      codigos <- codigos[ok]
      textos  <- textos[ok]

      if (!length(codigos)) {
        info_labels[[v]] <- list(
          var      = v,
          n_labels = length(labs),
          converted = FALSE,
          motivo   = "No se pudieron convertir los códigos a numérico (todos NA)."
        )
        next
      }

      # Eliminar posibles códigos duplicados
      dup     <- duplicated(codigos)
      codigos <- codigos[!dup]
      textos  <- textos[!dup]

      labs_new <- stats::setNames(codigos, textos)  # nombres = etiquetas, valores = códigos

      x_num <- suppressWarnings(as.numeric(x))
      all_na_num <- all(is.na(x_num)) && any(!is.na(x))

      df[[v]] <- haven::labelled_spss(
        x_num,
        labels = labs_new
      )

      if (!is.null(v_lab)) {
        attr(df[[v]], "label") <- v_lab
      }
      if (!is.null(meas)) {
        attr(df[[v]], "measure") <- meas
      }

      info_labels[[v]] <- list(
        var       = v,
        n_labels  = length(labs_new),
        converted = !all_na_num,
        motivo    = if (all_na_num) "Contenido convertido a numérico quedó todo en NA." else ""
      )
    }
  }

  # ---------------------------------------------------------------------------
  # 5) Escritura a disco en formato .sav
  # ---------------------------------------------------------------------------
  haven::write_sav(
    data    = df,
    path    = path_sav,
    compress = compress,
    ...
  )

  message("Archivo SPSS guardado en: ", normalizePath(path_sav, winslash = "/"))

  # ---------------------------------------------------------------------------
  # 6) Resumen sobre la aplicación de labels
  # ---------------------------------------------------------------------------
  if (length(info_labels) > 0L) {
    df_info <- do.call(rbind, lapply(info_labels, function(z) {
      data.frame(
        var       = z$var,
        n_labels  = z$n_labels,
        converted = z$converted,
        motivo    = z$motivo,
        stringsAsFactors = FALSE
      )
    }))

    n_total <- nrow(df_info)
    n_ok    <- sum(df_info$converted, na.rm = TRUE)
    n_prob  <- sum(!df_info$converted, na.rm = TRUE)

    message(
      "Resumen etiquetas SPSS: ",
      n_ok, " variable(s) con labels convertidas correctamente; ",
      n_prob, " con posibles problemas (ver 'motivo')."
    )

    if (n_prob > 0) {
      print(df_info[df_info$converted == FALSE, ], row.names = FALSE)
    }
  } else {
    message("No se encontraron variables con `attr(, \"labels\")` para convertir a labelled_spss.")
  }

  invisible(df)
}



#' Generar sintaxis SPSS para niveles de medida (ordinal / scale)
#'
#' A partir de una base de reporte (típicamente devuelta por [reporte_data()]),
#' identifica las variables que tienen el atributo `measure` y genera un
#' archivo `.sps` con instrucciones `VARIABLE LEVEL` para SPSS.
#'
#' En particular:
#' - `measure = "ordinal"` -> `VARIABLE LEVEL ... (ORDINAL).`
#' - `measure = "scale"`   -> `VARIABLE LEVEL ... (SCALE).`
#'
#' @param data Un `data.frame` o `tibble`, preferentemente el objeto devuelto
#'   por [reporte_data()] (clase `"prosecnur_reporte_tbl"`).
#' @param path_sps Ruta del archivo `.sps` a generar.
#' @param verbose Lógico; si `TRUE` imprime un mensaje con la ruta generada.
#'
#' @return Invisiblemente, una lista con los vectores de variables ordinales y
#'   de escala, junto con la ruta del archivo `.sps`.
#' @export
generar_spss_niveles <- function(data,
                                 path_sps = "niveles_medida.sps",
                                 verbose = TRUE) {

  if (!is.data.frame(data)) {
    stop("`data` debe ser un data.frame o tibble.", call. = FALSE)
  }

  # Detectar medida desde el atributo 'measure'
  vars_ordinal <- names(data)[vapply(
    data,
    function(x) identical(attr(x, "measure", exact = TRUE), "ordinal"),
    logical(1)
  )]

  vars_scale_raw <- names(data)[vapply(
    data,
    function(x) identical(attr(x, "measure", exact = TRUE), "scale"),
    logical(1)
  )]

  # Excluir fechas y horas del grupo SCALE
  vars_scale <- vars_scale_raw[!vapply(
    data[vars_scale_raw],
    function(x) inherits(x, c("Date", "POSIXct", "hms")),
    logical(1)
  )]

  lineas <- character(0L)

  # Helper para partir en bloques de hasta 20 vars
  add_variable_level_blocks <- function(vars, level) {
    if (length(vars) == 0L) return(character(0L))
    split_vars <- split(vars, ceiling(seq_along(vars) / 20))
    vapply(split_vars, function(v) {
      sprintf("VARIABLE LEVEL %s (%s).", paste(v, collapse = " "), toupper(level))
    }, character(1L))
  }

  lineas <- c(
    lineas,
    add_variable_level_blocks(vars_ordinal, "ordinal"),
    add_variable_level_blocks(vars_scale,   "scale")
  )

  if (length(lineas) == 0L) {
    warning("No se encontraron variables con atributo 'measure' = 'ordinal' o 'scale'.")
  }

  lineas <- c(lineas, "EXECUTE.")

  writeLines(lineas, path_sps, useBytes = TRUE)

  if (verbose) {
    message("Sintaxis SPSS guardada en: ", normalizePath(path_sps, winslash = "/"))
  }

  invisible(list(
    vars_ordinal = vars_ordinal,
    vars_scale   = vars_scale,
    path_sps     = path_sps
  ))
}
