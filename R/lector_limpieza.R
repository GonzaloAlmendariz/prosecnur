# =============================================================================
# lector_limpieza.R
# =============================================================================
# Función orquestadora para leer un Excel de datos ODK/XLSForm con secciones
# repetidas (repeats), detectar hoja principal, enlazar padre-hijo y producir
#:
#  - data$main: dataframe principal (con columnas n_<repeat>)
#  - data$repeats: lista de dataframes por repeat
#  - meta$keys: llaves detectadas (tabla, key, parent_key)
#  - meta$parent_map: orden: principal, repeats
#  - meta$counts: conteos por parent_key en cada repeat
#
# Soporta:
#  - _submission__id / _id (enlace por submission)
#  - _parent_index / _index (enlace por índice)
#  - Nombres de hojas y columnas con espacios/tildes/cambios de mayúsculas
#  - Hojas sin repeats
#
# Requiere: readxl, dplyr, tibble, purrr (stringi opcional)
# =============================================================================

# ---- Dependencias recomendadas ----
# install.packages(c("readxl","dplyr","tibble","purrr","stringi"))


# ---- Operador de coalescencia ----
`%||%` <- function(x, y) if (is.null(x) || (length(x) == 0)) y else x


# ---- Normalización de nombres de columnas ----
.tolower_names <- function(n) {
  n <- gsub("\\.+", "_", tolower(n))
  n <- gsub("\\s+", "_", n)
  n <- gsub("[^a-z0-9_]+", "_", n)
  n <- gsub("_+", "_", n)
  n <- gsub("^_|_$", "", n)
  n
}

.std_names <- function(df, tolower_names = .tolower_names) {
  if (!is.data.frame(df)) return(df)
  names(df) <- tolower_names(names(df))
  df
}

# ---- Lectura de todas las hojas del Excel ----
.read_all_sheets <- function(path) {
  if (!requireNamespace("readxl", quietly = TRUE)) {
    stop("Se requiere el paquete 'readxl'. Instálalo con install.packages('readxl').")
  }
  sh <- readxl::excel_sheets(path)
  out <- purrr::map(sh, function(s) readxl::read_excel(path, sheet = s, .name_repair = "minimal"))
  names(out) <- sh
  out
}

# ---- Normalización robusta de nombres de hojas (insensible a tildes, espacios, mayúsculas) ----
.strip_accents <- function(x) {
  if (!requireNamespace("stringi", quietly = TRUE)) return(tolower(trimws(x)))
  tolower(trimws(stringi::stri_trans_general(x, "Latin-ASCII")))
}
.canon_sheet <- function(x) {
  x <- .strip_accents(x)
  x <- gsub("\\s+", "_", x)
  x <- gsub("[^a-z0-9_]+", "_", x)
  x <- gsub("_+", "_", x)
  gsub("^_|_$", "", x)
}

# ---- ID seguro para tabla/columna ----
.safe_table_id <- function(x) {
  x <- tolower(gsub("\\s+", "_", x))
  x <- gsub("[^a-z0-9_]+", "_", x)
  x <- gsub("_+", "_", x)
  gsub("^_|_$", "", x)
}

# ---- Detección de hoja principal y repeats ----
.detect_main_sheet <- function(hojas, prefer_name = NULL) {
  if (!is.null(prefer_name) && prefer_name %in% names(hojas)) return(prefer_name)
  candidates <- names(hojas)[vapply(hojas, function(d) {
    d <- .std_names(d)
    (("_id" %in% names(d) || "_uuid" %in% names(d))) && !("_parent_index" %in% names(d))
  }, logical(1))]
  if (length(candidates) >= 1) return(candidates[1L])
  lens <- vapply(hojas, nrow, integer(1))
  names(hojas)[which.max(lens)]
}

.detect_repeats <- function(hojas, main_name, provided = NULL) {
  if (!is.null(provided)) return(setdiff(intersect(unique(provided), names(hojas)), main_name))
  reps <- names(hojas)[vapply(hojas, function(d) {
    d <- .std_names(d)
    ("_parent_index" %in% names(d)) || ("_submission__id" %in% names(d))
  }, logical(1))]
  setdiff(unique(reps), main_name)
}

# ---- Elección del enlace padre-hijo ----
# Devuelve: list(child_key, parent_key, mode = "submission"|"index"|NA)
.choose_parent_link <- function(child_df, parent_df) {
  child_df <- .std_names(child_df)
  parent_df <- .std_names(parent_df)

  # Modo 1: por submission (_submission__id en hijo) ↔ (_id en padre)
  if (("_submission__id" %in% names(child_df)) && ("_id" %in% names(parent_df))) {
    return(list(child_key = "_submission__id", parent_key = "_id", mode = "submission"))
  }

  # Modo 2: por índice (_parent_index en hijo) ↔ (_index en padre)
  if (("_parent_index" %in% names(child_df)) && ("_index" %in% names(parent_df))) {
    return(list(child_key = "_parent_index", parent_key = "_index", mode = "index"))
  }

  # Modo 3: fallback por UUID (raro, pero por si acaso)
  if (("_submission__uuid" %in% names(child_df)) && ("_uuid" %in% names(parent_df))) {
    return(list(child_key = "_submission__uuid", parent_key = "_uuid", mode = "uuid"))
  }

  list(child_key = NA_character_, parent_key = NA_character_, mode = NA_character_)
}

# ---- Conteos por parent_key ----
.compute_counts <- function(child_df, child_parent_key) {
  child_df <- .std_names(child_df)
  if (!(child_parent_key %in% names(child_df))) {
    return(tibble::tibble(parent_key = character(0), n = integer(0)))
  }
  child_df |>
    dplyr::mutate(!!child_parent_key := as.character(.data[[child_parent_key]])) |>
    dplyr::group_by(.data[[child_parent_key]], .add = FALSE) |>
    dplyr::summarise(n = dplyr::n(), .groups = "drop") |>
    dplyr::rename(parent_key = dplyr::all_of(child_parent_key)) |>
    dplyr::mutate(parent_key = as.character(.data$parent_key))
}

# ---- Inyección de conteos a la hoja principal ----
.left_join_counts_into_main <- function(main_df, parent_key, counts_df, ncol_name) {
  main_df <- .std_names(main_df)
  if (!(parent_key %in% names(main_df))) {
    # Intento de crear parent key si falta
    if (parent_key == "_index" && !("_index" %in% names(main_df))) {
      main_df$`_index` <- seq_len(nrow(main_df))
    } else if (parent_key == "_id" && !("_id" %in% names(main_df))) {
      main_df$`_id` <- NA_character_
    } else if (parent_key == "_uuid" && !("_uuid" %in% names(main_df))) {
      main_df$`_uuid` <- NA_character_
    }
  }

  if (!(parent_key %in% names(main_df))) {
    # No se puede unir
    return(main_df)
  }

  tmp <- counts_df
  if (!"n" %in% names(tmp)) tmp$n <- 0L
  names(tmp) <- dplyr::case_when(
    names(tmp) == "parent_key" ~ parent_key,
    TRUE ~ names(tmp)
  )

  out <- main_df |>
    dplyr::left_join(tmp, by = setNames(parent_key, parent_key))

  if (!ncol_name %in% names(out)) {
    # si ya existe columna "n" por el join, renombrar
    if ("n" %in% names(out)) {
      names(out)[names(out) == "n"] <- ncol_name
    } else {
      out[[ncol_name]] <- 0L
    }
  } else {
    # si ya existía, sobrescribir con el "n" que trajo el join
    if ("n" %in% names(out)) {
      out[[ncol_name]] <- out$n
    }
  }

  # Limpiar columna auxiliar "n" si quedó
  if ("n" %in% names(out)) out$n <- NULL
  out
}

#' Lector de datos ODK/XLSForm con repeats (limpieza y metadatos)
#'
#' @description
#' Lee un archivo Excel exportado desde ODK/Enketo/Kobo, detecta la hoja principal
#' y las hojas de secciones repetidas (repeats), enlaza padre-hijo usando
#' `_submission__id` ↔ `_id` (si existe) o `_parent_index` ↔ `_index` (fallback),
#' construye conteos `n_<repeat>` y devuelve un objeto con datos y metadatos.
#'
#' @param path Ruta al archivo Excel.
#' @param hoja_principal (opcional) Nombre de la hoja principal tal como aparece en Excel
#'   (no sensible a mayúsculas, espacios o tildes).
#' @param grupos_repetidos (opcional) Vector con nombres de hojas repeat.
#'   Si no se provee, se detectan automáticamente por presencia de `_parent_index`
#'   o `_submission__id` en sus columnas.
#'
#' @return Una lista con:
#' \itemize{
#'   \item \code{data$main}: dataframe principal con columnas `n_<repeat>`.
#'   \item \code{data$repeats}: lista de dataframes por repeat.
#'   \item \code{meta$keys}: tibble con columnas `table`, `key`, `parent_key`.
#'   \item \code{meta$parent_map}: vector con orden de tablas: principal + repeats.
#'   \item \code{meta$counts}: tibble largo con `parent_key`, `n`, `repeat_name`, `parent_table`.
#' }
#'
#' @examples
#' \dontrun{
#' datos <- lector_limpieza("datos_RMS.xlsx",
#'                          hoja_principal   = "RMS 2025 Perú - Q4",
#'                          grupos_repetidos = c("rpt_hhmnames","childedupe","s1"))
#' datos$meta$keys
#' names(datos$data$main) |> grep("^n_", ., value = TRUE)
#' }
#' @export
lector_limpieza <- function(path,
                            hoja_principal   = NULL,
                            grupos_repetidos = NULL) {

  # 1) Leer todas las hojas y crear mapa original ↔ canónico
  hojas_raw <- .read_all_sheets(path)
  if (length(hojas_raw) == 0) stop("No se encontraron hojas en el Excel: ", path)

  sheets_orig <- names(hojas_raw)
  sheets_can  <- .canon_sheet(sheets_orig)
  names(hojas_raw) <- sheets_can

  # Estandarizar columnas
  hojas <- lapply(hojas_raw, .std_names)

  # Resolver nombres canónicos del usuario (si se pasan)
  hoja_principal_can   <- if (!is.null(hoja_principal)) .canon_sheet(hoja_principal) else NULL
  grupos_repetidos_can <- if (!is.null(grupos_repetidos)) unique(.canon_sheet(grupos_repetidos)) else NULL

  # 2) Detectar principal y repeats (en nombres canónicos)
  main_name  <- .detect_main_sheet(hojas, prefer_name = hoja_principal_can)
  repeat_vec <- .detect_repeats(hojas, main_name, provided = grupos_repetidos_can)

  # Mensajería útil
  message("Hoja principal detectada: ", sheets_orig[match(main_name, sheets_can)])
  if (length(repeat_vec)) {
    pretty_reps <- sheets_orig[match(repeat_vec, sheets_can)]
    message("Repeats detectados: ", paste(pretty_reps, collapse = ", "))
  } else {
    message("No se detectaron repeats.")
  }

  # 3) Preparar dataframes principal/repeats
  main_df <- .std_names(hojas[[main_name]])
  if (!("_index" %in% names(main_df))) main_df$`_index` <- seq_len(nrow(main_df))

  # Detectar key en principal (prioridad: _id, _uuid, _index)
  parent_key_candidates <- c("_id", "_uuid", "_index")
  parent_key_in_parent  <- parent_key_candidates[parent_key_candidates %in% names(main_df)]
  parent_key_in_parent  <- if (length(parent_key_in_parent)) parent_key_in_parent[1] else NA_character_

  # meta$keys: fila para principal
  keys_tbl <- tibble::tibble(
    table      = .safe_table_id(main_name),
    key        = parent_key_in_parent,
    parent_key = NA_character_
  )

  # Lista de repeats “limpia” (sin duplicados)
  repeat_vec <- unique(repeat_vec)

  # 4) Construir metadatos de enlace y preparar n_<repeat>
  child_info <- list()
  for (rnm in repeat_vec) {
    ch <- .std_names(hojas[[rnm]])

    # Asegurar key local hija
    if (!("_index" %in% names(ch))) ch$`_index` <- seq_len(nrow(ch))
    child_key <- "_index"

    # Detectar enlace padre-hijo
    link <- .choose_parent_link(ch, main_df)

    child_info[[rnm]] <- list(
      child_key        = child_key,
      child_parent_key = link$child_key,
      parent_key       = link$parent_key,
      mode             = link$mode
    )

    # Añadir a meta$keys
    keys_tbl <- dplyr::bind_rows(
      keys_tbl,
      tibble::tibble(
        table      = .safe_table_id(rnm),
        key        = child_key,
        parent_key = link$child_key %||% NA_character_
      )
    )
  }

  # Deduplicar meta$keys por tabla
  keys_tbl <- keys_tbl |>
    dplyr::group_by(table) |>
    dplyr::slice(1) |>
    dplyr::ungroup()

  parent_map <- c(.safe_table_id(main_name), .safe_table_id(repeat_vec))

  # 5) Pre-crear columnas n_<repeat> en la principal
  if (length(repeat_vec)) {
    for (rnm in repeat_vec) {
      ncol_name <- paste0("n_", .safe_table_id(rnm))
      if (!(ncol_name %in% names(main_df))) main_df[[ncol_name]] <- 0L
    }
  }

  # 6) Calcular conteos por repeat y unir a la principal
  counts_long <- tibble::tibble(parent_key = character(), n = integer(),
                                repeat_name = character(), parent_table = character())
  repeats_out <- list()

  for (rnm in repeat_vec) {
    ch <- .std_names(hojas[[rnm]])
    repeats_out[[.safe_table_id(rnm)]] <- ch

    info <- child_info[[rnm]]

    # Validación del enlace
    if (is.na(info$child_parent_key) || is.na(info$parent_key) ||
        !(info$child_parent_key %in% names(ch)) || !(info$parent_key %in% names(main_df))) {

      warning(sprintf("No se pudo detectar enlace padre-hijo para '%s'. No se sumarán n_%s.",
                      sheets_orig[match(rnm, sheets_can)], .safe_table_id(rnm)))
      next
    }

    # Conteos
    cnt <- .compute_counts(ch, info$child_parent_key)
    if (!nrow(cnt)) next

    counts_long <- dplyr::bind_rows(
      counts_long,
      cnt |>
        dplyr::mutate(repeat_name = .safe_table_id(rnm),
                      parent_table = .safe_table_id(main_name))
    )

    # Inyectar n_<repeat> a la principal
    ncol_name <- paste0("n_", .safe_table_id(rnm))
    main_df <- .left_join_counts_into_main(main_df, info$parent_key, cnt, ncol_name)
  }

  # Asegurar NA→0 en todas las n_<repeat>
  if (length(repeat_vec)) {
    for (rnm in repeat_vec) {
      ncol_name <- paste0("n_", .safe_table_id(rnm))
      if (!(ncol_name %in% names(main_df))) main_df[[ncol_name]] <- 0L
      main_df[[ncol_name]][is.na(main_df[[ncol_name]])] <- 0L
    }
  }

  # 7) Salida
  list(
    data = list(
      main    = main_df,
      repeats = repeats_out
    ),
    meta = list(
      keys       = keys_tbl,
      parent_map = parent_map,
      counts     = counts_long
    )
  )
}
