# =============================================================================
# CONSTRUCTORES DE INDICADORES CATEGÓRICOS
# =============================================================================

#' Definir un nivel de un indicador
#'
#' Constructor de un nivel individual dentro de un indicador categórico.
#' Cada nivel tiene un código interno, una etiqueta humana y una regla lógica
#' (fórmula de un lado) que se evalúa contra el data.frame para determinar
#' qué filas pertenecen a este nivel.
#'
#' @param code Código interno del nivel (character). Debe ser único dentro
#'   del indicador. Ejemplo: \code{"asistio_reciente"}.
#' @param label Etiqueta humana del nivel. Ejemplo:
#'   \code{"Asistió en últimos 6 meses"}. Por defecto usa \code{code}.
#' @param regla Fórmula de un lado (\code{~ expresión}) que se evalúa contra
#'   el data.frame. Debe retornar un vector lógico. Ejemplo:
#'   \code{~ p17 \%in\% c("1", "2")}.
#'
#' @return Lista de clase \code{"prosecnur_nivel"}.
#' @family indicador
#' @export
nivel <- function(code, label = code, regla) {
  code  <- as.character(code)[1]
  label <- as.character(label)[1]


  if (!nzchar(code)) stop("`code` no puede estar vacío.", call. = FALSE)
  if (!inherits(regla, "formula")) {
    stop("`regla` debe ser una fórmula (~ expresión).", call. = FALSE)
  }

  out <- list(code = code, label = label, regla = regla)
  class(out) <- c("prosecnur_nivel", "list")
  out
}

#' Definir un indicador categórico
#'
#' Constructor de la especificación completa de un indicador categórico.
#' Un indicador agrupa múltiples niveles definidos con \code{\link{nivel}()},
#' cada uno con una regla lógica. Al aplicarse con
#' \code{\link{reporte_indicadores}()}, se crea una nueva variable categórica
#' en el data.frame.
#'
#' Los indicadores son primariamente categóricos (niveles sin puntaje).
#' Si se necesitan puntajes numéricos (escala 0-100), el indicador resultante
#' puede pasarse a \code{\link{reporte_recodificar_items}()} como cualquier
#' otra variable \code{select_one}.
#'
#' @param nombre Nombre de la variable resultante (sin prefijo). Ejemplo:
#'   \code{"acceso_salud"}.
#' @param etiqueta Etiqueta humana del indicador. Ejemplo:
#'   \code{"Acceso al servicio de salud"}. Por defecto usa \code{nombre}.
#' @param niveles Lista de objetos \code{\link{nivel}()}.
#' @param measure Nivel de medición: \code{"NOMINAL"} (default) u
#'   \code{"ORDINAL"}.
#' @param prioridad Resolución cuando una fila matchea múltiples niveles:
#'   \code{"primero"} (default) asigna el primer nivel que matchea según
#'   el orden de declaración.
#'
#' @return Lista de clase \code{"prosecnur_indicador"}.
#' @family indicador
#' @export
indicador <- function(
    nombre,
    etiqueta  = nombre,
    niveles,
    measure   = c("NOMINAL", "ORDINAL"),
    prioridad = c("primero", "ultimo")
) {
  nombre   <- as.character(nombre)[1]
  etiqueta <- as.character(etiqueta)[1]
  measure  <- match.arg(measure)
  prioridad <- match.arg(prioridad)

  if (!nzchar(nombre)) stop("`nombre` no puede estar vacío.", call. = FALSE)

  if (!is.list(niveles) || !length(niveles)) {
    stop("`niveles` debe ser una lista no vacía de objetos `nivel()`.", call. = FALSE)
  }
  for (i in seq_along(niveles)) {
    if (!inherits(niveles[[i]], "prosecnur_nivel")) {
      stop(
        sprintf("`niveles[[%d]]` no es un objeto `nivel()`. Usa nivel() para construirlo.", i),
        call. = FALSE
      )
    }
  }

  codes <- vapply(niveles, function(n) n$code, character(1))
  if (anyDuplicated(codes)) {
    dups <- codes[duplicated(codes)]
    stop(
      sprintf("Códigos duplicados en indicador '%s': %s", nombre, paste(dups, collapse = ", ")),
      call. = FALSE
    )
  }

  out <- list(
    nombre    = nombre,
    etiqueta  = etiqueta,
    niveles   = niveles,
    measure   = measure,
    prioridad = prioridad
  )
  class(out) <- c("prosecnur_indicador", "list")
  out
}

# =============================================================================
# FUNCIÓN PRINCIPAL
# =============================================================================

#' Construir indicadores categóricos a partir de reglas lógicas
#'
#' Toma un data.frame y una lista de definiciones de indicadores (construidos
#' con \code{\link{indicador}()} y \code{\link{nivel}()}) y crea nuevas
#' variables categóricas evaluando las reglas lógicas de cada nivel.
#'
#' Las variables resultantes se integran automáticamente con el instrumento
#' almacenado en \code{attr(data, "instrumento_reporte")}, lo que permite
#' usarlas directamente con \code{\link{reporte_frecuencias}()},
#' \code{\link{reporte_cruces}()}, \code{p_barras_apiladas()}, etc.
#'
#' @param data \code{data.frame} o \code{tibble}, típicamente la salida de
#'   \code{\link{reporte_data}()}.
#' @param indicadores Lista de objetos \code{\link{indicador}()}.
#' @param instrumento Objeto devuelto por \code{\link{reporte_instrumento}()}.
#'   Si es \code{NULL}, se toma desde \code{attr(data, "instrumento_reporte")}.
#' @param prefijo Prefijo para las columnas creadas. Por defecto \code{"ind_"}.
#'
#' @return El mismo \code{data.frame} con las nuevas columnas categóricas.
#'   Además actualiza el instrumento en \code{attr(data, "instrumento_reporte")}
#'   y agrega el atributo \code{indicadores_meta} con trazabilidad.
#'
#' @details
#' Para cada indicador, se evalúan las reglas de sus niveles en orden.
#' Con \code{prioridad = "primero"} (default), la primera regla que matchea
#' determina el nivel asignado. Filas que no matchean ningún nivel reciben
#' \code{NA}.
#'
#' Si una regla produce \code{NA} (por ejemplo, porque una variable tiene
#' \code{NA} en esa fila), se trata como \code{FALSE} (no matchea).
#'
#' @seealso \code{\link{indicador}}, \code{\link{nivel}},
#'   \code{\link{reporte_recodificar_items}},
#'   \code{\link{reporte_frecuencias}}, \code{\link{reporte_cruces}}
#' @family indicador
#' @export
reporte_indicadores <- function(
    data,
    indicadores,
    instrumento = NULL,
    prefijo     = "ind_"
) {
  `%||%` <- function(x, y) if (!is.null(x)) x else y

  if (!is.data.frame(data)) {
    stop("`data` debe ser un data.frame o tibble.", call. = FALSE)
  }

  if (!is.list(indicadores) || !length(indicadores)) {
    stop("`indicadores` debe ser una lista no vacía de objetos `indicador()`.", call. = FALSE)
  }

  # Validar que todos sean indicadores

  for (i in seq_along(indicadores)) {
    if (!inherits(indicadores[[i]], "prosecnur_indicador")) {
      stop(
        sprintf("`indicadores[[%d]]` no es un objeto `indicador()`. Usa indicador() para construirlo.", i),
        call. = FALSE
      )
    }
  }

  # Resolver instrumento
  if (is.null(instrumento)) {
    instrumento <- attr(data, "instrumento_reporte", exact = TRUE)
  }

  prefijo <- as.character(prefijo %||% "ind_")[1]
  n_rows  <- nrow(data)
  meta    <- list()

  for (ind in indicadores) {
    var_name <- paste0(prefijo, ind$nombre)
    list_name_sintetico <- paste0(ind$nombre, "_list")

    # --- Validar variables referenciadas ---
    vars_ref <- unique(unlist(lapply(ind$niveles, function(niv) {
      all.vars(niv$regla)
    })))
    faltantes <- setdiff(vars_ref, names(data))
    if (length(faltantes)) {
      stop(
        sprintf(
          "Indicador '%s': variables no encontradas en data: %s",
          ind$nombre, paste(faltantes, collapse = ", ")
        ),
        call. = FALSE
      )
    }

    # --- Evaluar reglas ---
    resultado <- rep(NA_character_, n_rows)
    asignado  <- rep(FALSE, n_rows)
    conteos   <- c(n_total = n_rows)
    n_niveles <- length(ind$niveles)

    for (j in seq_along(ind$niveles)) {
      niv <- ind$niveles[[j]]

      mask <- tryCatch(
        eval(niv$regla[[2]], envir = data, enclos = environment(niv$regla)),
        error = function(e) {
          stop(
            sprintf(
              "Indicador '%s', nivel '%s': error al evaluar regla: %s",
              ind$nombre, niv$code, conditionMessage(e)
            ),
            call. = FALSE
          )
        }
      )

      if (!is.logical(mask) || length(mask) != n_rows) {
        stop(
          sprintf(
            "Indicador '%s', nivel '%s': la regla debe retornar un vector lógico de largo %d.",
            ind$nombre, niv$code, n_rows
          ),
          call. = FALSE
        )
      }

      mask[is.na(mask)] <- FALSE

      if (identical(ind$prioridad, "primero")) {
        eligible <- mask & !asignado
      } else {
        eligible <- mask
      }

      resultado[eligible] <- niv$code
      asignado[eligible]  <- TRUE

      conteos[paste0("n_", niv$code)] <- sum(eligible)
    }

    conteos["n_sin_nivel"] <- sum(!asignado)

    # Warning si hay filas sin nivel
    if (conteos["n_sin_nivel"] > 0L) {
      message(
        sprintf(
          "Indicador '%s': %d fila(s) no matchearon ningún nivel (quedan como NA).",
          ind$nombre, conteos["n_sin_nivel"]
        )
      )
    }

    # --- Crear columna categórica ---
    codes  <- vapply(ind$niveles, function(n) n$code,  character(1))
    labels <- vapply(ind$niveles, function(n) n$label, character(1))

    # attr(labels) formato: c(code1 = "Label1", code2 = "Label2")
    labs_attr <- stats::setNames(labels, codes)

    data[[var_name]] <- resultado
    attr(data[[var_name]], "label")   <- ind$etiqueta
    attr(data[[var_name]], "labels")  <- labs_attr
    attr(data[[var_name]], "measure") <- tolower(ind$measure)

    # --- Actualizar instrumento ---
    if (!is.null(instrumento) && is.list(instrumento)) {
      # a) survey
      if (is.data.frame(instrumento$survey)) {
        new_survey_row <- data.frame(
          name      = var_name,
          type      = paste0("select_one ", list_name_sintetico),
          list_name = list_name_sintetico,
          label     = ind$etiqueta,
          stringsAsFactors = FALSE
        )
        # Asegurar columnas que falten
        for (col in setdiff(names(instrumento$survey), names(new_survey_row))) {
          new_survey_row[[col]] <- NA
        }
        new_survey_row <- new_survey_row[, names(instrumento$survey), drop = FALSE]
        instrumento$survey <- rbind(instrumento$survey, new_survey_row)
      }

      # b) orders_list
      if (is.list(instrumento$orders_list)) {
        instrumento$orders_list[[var_name]] <- list(
          names  = codes,
          labels = labels,
          label  = ind$etiqueta
        )
      }

      # c) dicc_code_to_label
      if (is.list(instrumento$dicc_code_to_label)) {
        instrumento$dicc_code_to_label[[list_name_sintetico]] <- labs_attr
      }

      # d) dicc_label_to_code
      if (is.list(instrumento$dicc_label_to_code)) {
        instrumento$dicc_label_to_code[[list_name_sintetico]] <- stats::setNames(codes, labels)
      }

      # e) var_labels
      if (!is.null(instrumento$var_labels)) {
        instrumento$var_labels[var_name] <- ind$etiqueta
      }
    }

    # --- Metadata ---
    reglas_texto <- vapply(ind$niveles, function(niv) {
      deparse(niv$regla[[2]], width.cutoff = 500L)
    }, character(1))

    meta[[ind$nombre]] <- list(
      nombre       = ind$nombre,
      etiqueta     = ind$etiqueta,
      variable     = var_name,
      measure      = ind$measure,
      prioridad    = ind$prioridad,
      niveles      = lapply(ind$niveles, function(niv) {
        list(code = niv$code, label = niv$label)
      }),
      conteos      = conteos,
      reglas_texto = reglas_texto
    )
  }

  # Guardar instrumento actualizado
  if (!is.null(instrumento)) {
    attr(data, "instrumento_reporte") <- instrumento
  }

  # Guardar metadata
  existing_meta <- attr(data, "indicadores_meta", exact = TRUE) %||% list()
  attr(data, "indicadores_meta") <- c(existing_meta, meta)

  data
}
