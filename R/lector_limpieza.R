# lector_limpieza.R
# ============================================================
# Cargadores necesarios
# ============================================================
#' @importFrom readxl excel_sheets read_excel
#' @importFrom dplyr mutate select rename left_join group_by summarise n ungroup bind_rows distinct
#' @importFrom tibble tibble as_tibble
#' @importFrom purrr map set_names compact
#' @importFrom rlang .data
NULL

# ============================================================
# Helpers internos con prefijo ll_ (evitan choques de nombres)
# ============================================================

ll_std_names <- function(df) {
  if (!is.data.frame(df)) return(df)
  nn <- names(df)
  norm <- function(s) {
    s0 <- gsub("\\s+", "_", trimws(tolower(as.character(s))))
    s0 <- chartr("áéíóúñüÁÉÍÓÚÑÜ", "aeiounuAEIOUNU", s0)
    s0
  }
  names(df) <- vapply(nn, norm, character(1))
  df
}

ll_sheet_key_candidates <- function(df) {
  nms <- names(df)
  pick <- function(x) any(x == nms)
  if (pick("_id"))    return("_id")
  if (pick("_uuid"))  return("_uuid")
  if (pick("__id"))   return("__id")
  if (pick("__uuid")) return("__uuid")
  if (pick("_index")) return("_index")
  NA_character_
}

ll_find_col <- function(df, candidates) {
  nms <- names(df)
  i <- which(tolower(nms) %in% tolower(candidates))[1]
  if (length(i) == 1 && !is.na(i)) nms[i] else NA_character_
}

ll_canon <- function(x) {
  x <- tolower(trimws(as.character(x)))
  x <- gsub("\\s+", " ", x)
  x <- chartr("áéíóúñüÁÉÍÓÚÑÜ", "aeiounuAEIOUNU", x)
  x
}

ll_match_sheet <- function(target, pool) {
  if (is.null(target) || !nzchar(target)) return(NA_character_)
  ct <- ll_canon(target)
  canon_pool <- ll_canon(pool)
  j <- match(ct, canon_pool)
  if (is.na(j)) NA_character_ else pool[j]
}

ll_detect_main_sheet <- function(hojas, hoja_principal = NULL) {
  nms <- names(hojas)
  if (length(nms) == 0) return(NA_character_)
  if (nzchar(hoja_principal %||% "")) {
    mm <- ll_match_sheet(hoja_principal, nms)
    if (isTRUE(nzchar(mm))) return(mm)
    warning(sprintf("No se encontró la hoja principal '%s'. Se intentará inferirla.", hoja_principal), call. = FALSE)
  }
  sizes <- vapply(hojas, nrow, integer(1))
  order_idx <- order(sizes, decreasing = TRUE)
  for (i in order_idx) {
    key <- ll_sheet_key_candidates(hojas[[i]])
    if (!is.na(key)) return(nms[i])
  }
  nms[order_idx[1]]
}

ll_detect_repeats_infer <- function(hojas, main_name) {
  setdiff(names(hojas)[vapply(hojas, function(d) {
    d <- ll_std_names(d)
    any(c("_parent_index","_submission__id","_parent_table_name") %in% names(d))
  }, logical(1))], main_name)
}

ll_link_children <- function(parent_df, child_df, child_name, parent_label) {
  p <- ll_std_names(parent_df)
  c <- ll_std_names(child_df)

  parent_key <- ll_find_col(p, c("_index","_id","_uuid","__id","__uuid"))
  child_pk   <- ll_find_col(c, c("_index"))
  child_fk   <- ll_find_col(c, c("_parent_index","parent_index"))

  counts <- tibble::tibble()
  if (!is.na(child_fk)) {
    counts <- c %>%
      dplyr::group_by(.data[[child_fk]]) %>%
      dplyr::summarise(n = dplyr::n(), .groups = "drop") %>%
      dplyr::rename(parent_key = !!child_fk) %>%
      # >>> CORRECCIÓN: forzar parent_key a character para evitar choques de tipo
      dplyr::mutate(parent_key = as.character(.data$parent_key),
                    repeat_name = child_name,
                    parent_table = parent_label)
  }

  parent_aug <- p
  if (!is.na(parent_key) && !is.na(child_fk) && identical(parent_key, "_index")) {
    parent_aug <- parent_aug %>%
      # asegurar tipos comparables en el join
      dplyr::mutate(`_index` = as.character(.data[["_index"]])) %>%
      dplyr::left_join(counts %>% dplyr::select(parent_key, n),
                       by = c("_index" = "parent_key")) %>%
      dplyr::rename(!!paste0("n_", child_name) := n)
  } else if (!is.na(parent_key) && !is.na(child_fk)) {
    tmp_counts <- counts %>% dplyr::mutate(parent_key = as.character(.data$parent_key))
    parent_aug <- parent_aug %>%
      dplyr::mutate(`__join_key__` = as.character(.data[[parent_key]])) %>%
      dplyr::left_join(tmp_counts, by = c("__join_key__" = "parent_key")) %>%
      dplyr::rename(!!paste0("n_", child_name) := n) %>%
      dplyr::select(-"__join_key__")
  }

  list(counts = counts, parent_aug = parent_aug)
}

`%||%` <- function(a,b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a

# ============================================================
# Orquestador principal
# ============================================================

#' Lector robusto de Excel RMS/ODK con hojas repetidas
#'
#' Lee un archivo Excel exportado de KoBo/ODK (o similar), normaliza nombres
#' de columnas y detecta relaciones padre–hijo entre hoja principal y grupos
#' repetidos. Agrega conteos `n_<repeat>` a la hoja principal cuando existe
#' el vínculo, y devuelve metadatos útiles para validación.
#'
#' @param archivo Ruta al .xlsx
#' @param hoja_principal Nombre de la hoja principal (tolerante a acentos/espacios).
#'   Si no se proporciona, se infiere.
#' @param grupos_repetidos Vector opcional con nombres de hojas repetidas (tolerante).
#'   Si no se proporciona, se detectan por columnas típicas (`_parent_index`, etc.).
#' @return Una lista con:
#'   - \code{data}: lista de data.frames por hoja (normalizados)
#'   - \code{meta$keys}: tibble con claves por hoja (key y parent_key)
#'   - \code{meta$parent_map}: tibble con relaciones hijo→padre
#'   - \code{counts}: tibble con conteos por padre para cada repeat
#' @export
lector_limpieza <- function(archivo,
                            hoja_principal   = NULL,
                            grupos_repetidos = NULL) {

  # -------- 1) Leer todas las hojas ----------
  hojas_nombres <- readxl::excel_sheets(archivo)
  hojas <- purrr::map(hojas_nombres, ~ readxl::read_excel(archivo, sheet = .x))
  names(hojas) <- hojas_nombres
  hojas <- purrr::map(hojas, ll_std_names)

  # -------- 2) Resolver hoja principal ----------
  main_name <- ll_detect_main_sheet(hojas, hoja_principal)
  if (is.na(main_name)) stop("No se pudo determinar la hoja principal.", call. = FALSE)

  # -------- 3) Resolver grupos repetidos ----------
  grupos_repetidos_can <- character(0)
  if (!is.null(grupos_repetidos) && length(grupos_repetidos)) {
    m <- vapply(grupos_repetidos, ll_match_sheet, character(1), pool = names(hojas))
    grupos_repetidos_can <- unique(na.omit(m))
  } else {
    grupos_repetidos_can <- ll_detect_repeats_infer(hojas, main_name)
  }

  # -------- 4) meta$keys (padre) ----------
  main_df  <- hojas[[main_name]]
  main_key <- ll_sheet_key_candidates(main_df)
  if (is.na(main_key)) {
    main_df <- dplyr::mutate(main_df, `_index` = dplyr::row_number())
    hojas[[main_name]] <- main_df
    main_key <- "_index"
  }

  keys_rows <- list(tibble::tibble(
    table = as.character(main_name),
    key   = as.character(main_key),
    parent_key = NA_character_
  ))

  # -------- 5) Vinculación y conteos ----------
  counts_all <- tibble::tibble(
    parent_key  = character(),
    n           = integer(),
    repeat_name = character(),
    parent_table= character()
  )

  parent_aug <- main_df
  for (child in grupos_repetidos_can) {
    child_df <- hojas[[child]]

    child_fk <- ll_find_col(child_df, c("_parent_index","parent_index"))
    child_pk <- ll_find_col(child_df, c("_index"))

    keys_rows[[length(keys_rows) + 1]] <- tibble::tibble(
      table = as.character(child),
      key   = if (!is.na(child_pk)) as.character(child_pk) else NA_character_,
      parent_key = if (!is.na(child_fk)) as.character(child_fk) else NA_character_
    )

    link <- ll_link_children(main_df, child_df, child_name = child, parent_label = main_name)
    counts_all <- dplyr::bind_rows(counts_all, link$counts)
    parent_aug <- link$parent_aug
  }

  hojas[[main_name]] <- parent_aug

  # -------- 6) meta$parent_map ----------
  parent_map <- tibble::tibble(
    child  = grupos_repetidos_can,
    parent = if (length(grupos_repetidos_can)) rep(main_name, length(grupos_repetidos_can)) else character(0)
  )

  # -------- 7) meta$keys ----------
  meta_keys <- dplyr::bind_rows(keys_rows) %>% dplyr::distinct()

  # -------- 8) Salida ----------
  list(
    data = hojas,
    meta = list(
      main  = main_name,
      keys  = meta_keys,
      parent_map = parent_map
    ),
    counts = counts_all
  )
}
