# =============================================================================
# ppra_adaptar_instrumento
# =============================================================================

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(stringr)
  library(tidyr)
  library(tibble)
})

`%||%` <- function(a,b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a

# ---------- helpers básicos ---------------------------------------------------

.guess_label_col <- function(df){
  nms <- tolower(names(df))
  hit <- match(TRUE, nms %in% c(
    "label::spanish (es)","label::spanish(es)","label::spanish_es",
    "label_spanish_es","label::spanish","label","label::es"
  ))
  if (is.na(hit)) "label" else names(df)[hit]
}

.norm_tokens <- function(x){
  x <- as.character(x)
  x <- gsub("[,;|/]+"," ", x, perl = TRUE)
  x <- gsub("\\s+"," ", x, perl = TRUE)
  x <- trimws(x)
  x[x==""] <- NA_character_
  x
}

.split_tokens <- function(v){
  v <- .norm_tokens(v)
  strsplit(ifelse(is.na(v),"",v), "\\s+")
}

.sanitize <- function(x){
  x <- gsub("[^A-Za-z0-9_]+","_", x)
  x <- gsub("_+","_", x)
  x <- sub("^_+","", x); x <- sub("_+$","", x)
  if (!nzchar(x)) "ppra_list" else x
}

.row_after <- function(df, row_idx, new_row){
  if (is.na(row_idx) || row_idx<=0 || row_idx>=nrow(df)) {
    dplyr::bind_rows(df, new_row)
  } else {
    dplyr::bind_rows(df[seq_len(row_idx), , drop = FALSE],
                     new_row,
                     df[(row_idx+1):nrow(df), , drop = FALSE])
  }
}

.extract_listname <- function(type_str){
  s <- as.character(type_str) %||% ""
  s <- trimws(s)
  parts <- strsplit(s, "\\s+")[[1]]
  if (!length(parts)) return(NA_character_)
  if (length(parts)>=2 && parts[1] %in% c("select_one","select_multiple")) parts[2] else NA_character_
}

.read_all_sheets <- function(path_xlsx){
  sh <- readxl::excel_sheets(path_xlsx)
  stats::setNames(lapply(sh, function(s){
    tryCatch(readxl::read_xlsx(path_xlsx, sheet = s), error = function(e) NULL)
  }), sh)
}

.collect_tokens_from_col <- function(df_or_path, col_name){
  # Une tokens de col_name en todas las hojas donde exista.
  if (is.character(df_or_path) && file.exists(df_or_path)) {
    lst <- .read_all_sheets(df_or_path)
  } else if (is.data.frame(df_or_path)) {
    lst <- list(DATA = df_or_path)
  } else stop("path_data_adaptada debe ser data.frame o ruta a XLSX con la data adaptada.")
  toks <- character(0)
  for (nm in names(lst)){
    d <- lst[[nm]]; if (is.null(d) || !ncol(d)) next
    if (col_name %in% names(d)) {
      v  <- d[[col_name]]
      vv <- unique(unlist(.split_tokens(v)))
      toks <- c(toks, vv)
    }
  }
  unique(toks[nzchar(toks)])
}

.collect_child_cols <- function(df_or_path, parent){
  # Devuelve nombres de columnas hijas *_recod a lo largo de todas las hojas:
  # ^<parent>_.+_recod$, excluyendo <parent>_recod
  rx   <- paste0("^", gsub("([\\W])","\\\\\\1", parent), "_.+_recod$")
  main <- paste0(parent, "_recod")

  if (is.character(df_or_path) && file.exists(df_or_path)) {
    lst <- .read_all_sheets(df_or_path)
  } else if (is.data.frame(df_or_path)) {
    lst <- list(DATA = df_or_path)
  } else stop("path_data_adaptada debe ser data.frame o ruta a XLSX con la data adaptada.")

  out <- character(0)
  for (nm in names(lst)){
    d <- lst[[nm]]; if (is.null(d) || !ncol(d)) next
    hits <- grep(rx, names(d), value = TRUE, perl = TRUE)
    hits <- setdiff(hits, main)
    out  <- c(out, hits)
  }
  unique(out)
}

.style <- function(hex) openxlsx::createStyle(fgFill = hex)

.paint_new <- function(wb, sheet, rows, ncols, hex){
  rows <- unique(rows[!is.na(rows)])
  if (!length(rows)) return()
  openxlsx::addStyle(wb, sheet, .style(hex),
                     rows = rows, cols = 1:ncols,
                     gridExpand = TRUE, stack = TRUE)
}

# ---------- núcleo: insertar una pregunta *_recod + lista ---------------------

.add_recoded_q <- function(survey, choices,
                           base_name,
                           kind = c("multiple","one"),
                           list_name_hint = NULL,
                           tokens_from_data = character(0),
                           lab_col_s,
                           lab_col_c,
                           choices_order = c("original_first","by_first_seen","alphabetical"),
                           insert_below_original = TRUE,
                           copy_from_original    = TRUE,
                           new_name_override     = NULL){

  kind          <- match.arg(kind)
  choices_order <- match.arg(choices_order)

  # fila base (si existe)
  i_base <- match(base_name, survey$name)
  base_type  <- if (!is.na(i_base)) as.character(survey$type[i_base]) else {
    if (kind=="multiple") "select_multiple" else "select_one"
  }
  base_label <- if (!is.na(i_base)) {
    as.character(survey[[lab_col_s]][i_base] %||% base_name)
  } else base_name

  # list original (si existe)
  ln_orig <- .extract_listname(base_type)

  # list destino
  base_list <- if (!is.null(list_name_hint)) {
    list_name_hint
  } else if (!is.na(ln_orig)) {
    paste0(ln_orig, "_recod")
  } else {
    paste0(.sanitize(base_name), "_recod")
  }

  # nombre nuevo de la pregunta
  new_name <- if (!is.null(new_name_override)) new_name_override else paste0(base_name, "_recod")
  new_type <- if (kind=="multiple") {
    paste("select_multiple", base_list)
  } else {
    paste("select_one", base_list)
  }

  # catálogo a crear
  codes   <- character(0)
  labels  <- character(0)
  copied_original <- FALSE

  # copiar catálogo original si aplica (SM / SO padre)
  if (copy_from_original && is.null(list_name_hint) && !is.na(ln_orig) &&
      ln_orig %in% choices$list_name) {
    orig <- choices %>% dplyr::filter(.data$list_name == ln_orig)
    if (nrow(orig)) {
      copied_original <- TRUE
      codes  <- c(codes, as.character(orig$name))
      labcol <- lab_col_c
      labels <- c(labels, as.character(orig[[labcol]]))
    }
  }

  # añadir tokens observados (de la data)
  if (length(tokens_from_data)) {
    seen <- unique(tokens_from_data[nzchar(tokens_from_data)])
    new_codes <- setdiff(seen, codes)
    if (length(new_codes)) {
      codes  <- c(codes, new_codes)
      labels <- c(labels, rep(NA_character_, length(new_codes)))
    }
  }

  # completar labels desde códigos cuando no se copió catálogo original
  if (!copied_original && length(codes)) {
    na_lab <- is.na(labels) | !nzchar(labels)
    labels[na_lab] <- codes[na_lab]
  }

  # ordenar catálogo
  if (choices_order == "alphabetical" && length(codes)) {
    o <- order(codes)
    codes  <- codes[o]
    labels <- labels[o]
  }

  # fila nueva en survey (debajo de la base si existe)
  new_row <- survey[0,]
  new_row[1, setdiff(names(survey), character(0))] <- NA
  new_row$type <- new_type
  new_row$name <- new_name
  new_row[[lab_col_s]] <- base_label  # misma etiqueta que la base

  survey2 <- .row_after(survey, i_base, new_row)

  # inyectar choices del list destino (evitando duplicados exactos)
  if (length(codes)) {
    add_choices <- tibble::tibble(list_name = base_list,
                                  name      = codes)
    add_choices[[lab_col_c]] <- labels

    dup_mask <- paste(choices$list_name, choices$name) %in%
      paste(add_choices$list_name, add_choices$name)
    if (any(dup_mask)) {
      choices <- choices[!dup_mask, , drop = FALSE]
    }

    # posición: debajo del catálogo original (si se desea y existe), o al final
    below <- NA_integer_
    if (insert_below_original && !is.na(ln_orig) && ln_orig %in% choices$list_name) {
      hit <- which(choices$list_name == ln_orig)
      if (length(hit)) below <- max(hit)
    }

    if (!is.na(below) && below < nrow(choices)) {
      choices2 <- dplyr::bind_rows(
        choices %>% dplyr::slice(1:below),
        add_choices,
        choices %>% dplyr::slice((below+1):dplyr::n())
      )
    } else {
      choices2 <- dplyr::bind_rows(choices, add_choices)
    }
  } else {
    choices2 <- choices
  }

  list(
    survey    = survey2,
    choices   = choices2,
    new_name  = new_name,
    list_name = base_list
  )
}

# =============================================================================
#' @title Adaptar instrumento XLSForm a partir de una data recodificada
#'
#' @description
#' Genera una versión adaptada de un instrumento XLSForm (hojas \code{survey} y
#' \code{choices}) a partir de una data ya recodificada (output de
#' \code{ppra_adaptar_data}), creando preguntas y listas nuevas para documentar
#' los códigos de recodificación.
#'
#' La función:
#' \itemize{
#'   \item Para variables \strong{select\_multiple} (\code{sm_vars}): crea una nueva
#'   pregunta \code{<parent>_recod} de tipo \code{select_multiple} con una lista
#'   derivada del catálogo original más los códigos observados en la data. La nueva
#'   lista se inserta \emph{debajo} del catálogo original en \code{choices}.
#'
#'   \item Para variables \strong{select\_one} que recodifican al padre
#'   (\code{so_parent_vars}): crea una nueva pregunta \code{<parent>_recod} de tipo
#'   \code{select_one} con una lista derivada del catálogo original más los códigos
#'   observados en la data. La nueva lista también se inserta \emph{debajo} del
#'   catálogo original en \code{choices}.
#'
#'   \item Para variables \strong{select\_one} cuyo recodificado está en el hijo
#'   (\code{so_child_vars}): detecta columnas hijas
#'   \code{^<parent>_.+_recod$} (excluyendo \code{<parent>_recod}) en todas
#'   las hojas de la data adaptada y crea, para cada una, una nueva pregunta
#'   cuyo nombre coincide con la columna hija (por ejemplo,
#'   \code{calidad_docente_why_recod}) de tipo \code{select_one} con una lista
#'   propia construida a partir de los valores observados en dicha columna.
#'   Estas listas se añaden al final de \code{choices}, con nombres del tipo
#'   \code{<child>_lista} (por ejemplo, \code{calidad_docente_why_recod_lista},
#'   sin sufijo extra \code{_recod}).
#'
#'   \item Para variables \strong{integer} (\code{integer_vars}): crea una nueva
#'   pregunta \code{<var>_recod} de tipo \code{select_one <var>_recod_lista}.
#'   La lista \code{<var>_recod_lista} se construye a partir de los valores
#'   únicos observados en \code{<var>_recod} en la data adaptada (todas las hojas)
#'   y se añade al final de \code{choices}.
#' }
#'
#' En todos los casos, la nueva pregunta hereda exactamente la misma etiqueta
#' (\code{label}) que la pregunta base, sin añadir sufijos como \code{"[recod]"}.
#'
#' Además, si \code{paint = TRUE}, la función colorea:
#' \itemize{
#'   \item Las nuevas filas de \code{survey}: verde para \code{sm_vars}, azul
#'   para \code{so_parent_vars} y \code{so_child_vars}, y morado para
#'   \code{integer_vars}.
#'   \item Las filas de \code{choices} correspondientes a las nuevas listas:
#'   verde claro (SM), azul claro (SO) y morado claro (integer).
#' }
#'
#' @param path_instrumento_in Ruta al XLSX del instrumento original. Debe contener,
#'   al menos, las hojas \code{"survey"} y \code{"choices"}.
#' @param path_data_adaptada Data adaptada sobre la cual se construyen los nuevos
#'   catálogos y preguntas. Puede ser:
#'   \itemize{
#'     \item una ruta a un archivo XLSX (potencialmente con varias hojas), o
#'     \item un \code{data.frame} con las columnas \code{*_recod}.
#'   }
#'   Si es XLSX, la función lee todas las hojas para recolectar los tokens.
#' @param path_instrumento_out Ruta del XLSX de salida con el instrumento
#'   adaptado. Por defecto \code{"instrumento_adaptado.xlsx"}.
#' @param sm_vars Vector de nombres de variables \code{select_multiple} (padres)
#'   para las que se creará \code{<parent>_recod} como \code{select_multiple}.
#' @param so_parent_vars Vector de nombres de variables \code{select_one} (padres)
#'   para las que se creará \code{<parent>_recod} como \code{select_one}, usando
#'   el catálogo del padre más los códigos observados en la data.
#' @param so_child_vars Vector de nombres de variables \code{select_one} cuyos
#'   recodificados están en columnas hijas \code{<parent>_<alias>_recod}. Para
#'   cada hija detectada se crea una nueva pregunta (con el mismo nombre de la
#'   columna hija) con lista propia, basada en los valores observados en la data
#'   adaptada.
#' @param integer_vars Vector de nombres de variables \code{integer} para las
#'   que se creará una nueva pregunta \code{<var>_recod} de tipo
#'   \code{select_one <var>_recod_lista}, donde \code{<var>_recod_lista} se
#'   construye a partir de los valores únicos observados en \code{<var>_recod}.
#' @param choices_order Criterio de orden para los códigos en las nuevas listas.
#'   Puede ser \code{"original_first"} (por defecto), \code{"alphabetical"} o
#'   \code{"by_first_seen"}.
#' @param paint Lógico; si es \code{TRUE}, colorea las filas nuevas en
#'   \code{survey} y las listas nuevas en \code{choices} para facilitar la
#'   inspección visual.
#'
#' @return Invisiblemente, una lista con:
#'   \itemize{
#'     \item \code{survey}: data.frame de la hoja \code{survey} adaptada.
#'     \item \code{choices}: data.frame de la hoja \code{choices} adaptada.
#'     \item \code{out_path}: ruta al XLSX escrito en disco.
#'   }
#'
#' @export
# =============================================================================
ppra_adaptar_instrumento <- function(path_instrumento_in,
                                     path_data_adaptada,
                                     path_instrumento_out = "instrumento_adaptado.xlsx",
                                     sm_vars        = character(0),
                                     so_parent_vars = character(0),
                                     so_child_vars  = character(0),
                                     integer_vars   = character(0),
                                     choices_order  = c("original_first","by_first_seen","alphabetical"),
                                     paint = TRUE){

  choices_order <- match.arg(choices_order)
  stopifnot(file.exists(path_instrumento_in))

  # --- leer instrumento base ---
  survey  <- readxl::read_excel(path_instrumento_in, sheet = "survey")
  choices <- readxl::read_excel(path_instrumento_in, sheet = "choices")
  if (!all(c("type","name") %in% names(survey)))
    stop("survey debe tener columnas 'type' y 'name'.")
  if (!all(c("list_name","name") %in% names(choices)))
    stop("choices debe tener columnas 'list_name' y 'name'.")

  lab_col_s <- .guess_label_col(survey)
  lab_col_c <- .guess_label_col(choices)

  survey$name       <- as.character(survey$name)
  choices$name      <- as.character(choices$name)
  choices$list_name <- as.character(choices$list_name)

  # --- preparar acceso multi-hojas ---
  df_is_xlsx <- is.character(path_data_adaptada) && file.exists(path_data_adaptada)
  df_single  <- is.data.frame(path_data_adaptada)
  if (!df_is_xlsx && !df_single)
    stop("path_data_adaptada debe ser data.frame o ruta a XLSX con la data adaptada.")

  # --- logs de NUEVOS NOMBRES (survey) y listas nuevas (choices) -------------
  new_names_sm   <- character(0)   # nombres nuevos para SM
  new_names_so   <- character(0)   # nombres nuevos para SO (padre + hijo)
  new_names_int  <- character(0)   # nombres nuevos para INTEGER

  new_lists_sm   <- character(0)   # list_name nuevas SM
  new_lists_so   <- character(0)   # list_name nuevas SO (padre + hijo)
  new_lists_int  <- character(0)   # list_name nuevas INTEGER

  # =======================
  # 1) SELECT MULTIPLE (padre)
  # =======================
  if (length(sm_vars)) {
    for (pv in sm_vars) {
      col_rec <- paste0(pv, "_recod")
      toks <- .collect_tokens_from_col(path_data_adaptada, col_rec)

      res <- .add_recoded_q(survey, choices,
                            base_name      = pv,
                            kind           = "multiple",
                            list_name_hint = NULL,  # usa <ln_orig>_recod
                            tokens_from_data = toks,
                            lab_col_s      = lab_col_s,
                            lab_col_c      = lab_col_c,
                            choices_order  = choices_order,
                            insert_below_original = TRUE,
                            copy_from_original    = TRUE,
                            new_name_override     = NULL)

      survey  <- res$survey
      choices <- res$choices

      new_names_sm   <- c(new_names_sm,   res$new_name)
      new_lists_sm   <- c(new_lists_sm,   res$list_name)
    }
  }

  # =======================
  # 2) SELECT ONE (padre)
  # =======================
  if (length(so_parent_vars)) {
    for (pv in so_parent_vars) {
      col_rec <- paste0(pv, "_recod")
      toks <- .collect_tokens_from_col(path_data_adaptada, col_rec)

      res <- .add_recoded_q(survey, choices,
                            base_name      = pv,
                            kind           = "one",
                            list_name_hint = NULL,  # usa <ln_orig>_recod
                            tokens_from_data = toks,
                            lab_col_s      = lab_col_s,
                            lab_col_c      = lab_col_c,
                            choices_order  = choices_order,
                            insert_below_original = TRUE,
                            copy_from_original    = TRUE,
                            new_name_override     = NULL)

      survey  <- res$survey
      choices <- res$choices

      new_names_so   <- c(new_names_so,   res$new_name)
      new_lists_so   <- c(new_lists_so,   res$list_name)
    }
  }

  # =======================
  # 3) SELECT ONE (hijo)
  # =======================
  if (length(so_child_vars)) {
    for (pv in so_child_vars) {
      child_cols <- .collect_child_cols(path_data_adaptada, pv)
      if (!length(child_cols)) next

      for (cc in child_cols) {
        toks <- .collect_tokens_from_col(path_data_adaptada, cc)
        toks <- toks[nzchar(toks)]

        # base para la etiqueta: versión sin _recod
        base_child_name <- sub("(?i)_recod$", "", cc, perl = TRUE)

        res <- .add_recoded_q(survey, choices,
                              base_name      = base_child_name,                 # para type/label
                              kind           = "one",
                              list_name_hint = paste0(.sanitize(cc), "_lista"), # ej: calidad_docente_why_recod_lista
                              tokens_from_data = toks,
                              lab_col_s      = lab_col_s,
                              lab_col_c      = lab_col_c,
                              choices_order  = choices_order,
                              insert_below_original = FALSE,   # listas al final
                              copy_from_original    = FALSE,   # sin catálogo original
                              new_name_override     = cc)      # nombre EXACTO de la col hija recod

        survey  <- res$survey
        choices <- res$choices

        new_names_so   <- c(new_names_so,   res$new_name)
        new_lists_so   <- c(new_lists_so,   res$list_name)
      }
    }
  }

  # =======================
  # 4) INTEGER
  # =======================
  if (length(integer_vars)) {
    for (pv in integer_vars) {
      col_rec <- paste0(pv, "_recod")
      toks <- .collect_tokens_from_col(path_data_adaptada, col_rec)

      res <- .add_recoded_q(survey, choices,
                            base_name      = pv,
                            kind           = "one",
                            list_name_hint = paste0(pv, "_recod_lista"),
                            tokens_from_data = toks,
                            lab_col_s      = lab_col_s,
                            lab_col_c      = lab_col_c,
                            choices_order  = choices_order,
                            insert_below_original = FALSE,  # listas al final
                            copy_from_original    = FALSE,  # no hay catálogo original
                            new_name_override     = NULL)   # genera <var>_recod

      survey  <- res$survey
      choices <- res$choices

      new_names_int  <- c(new_names_int,  res$new_name)
      new_lists_int  <- c(new_lists_int,  res$list_name)
    }
  }

  # =======================
  # 5) Exportar + colorear
  # =======================
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    warning("No se encontró 'openxlsx'. Se guarda sin color.")
    openxlsx::write.xlsx(list(survey = survey, choices = choices),
                         path_instrumento_out, overwrite = TRUE)
  } else {
    wb <- openxlsx::createWorkbook()
    openxlsx::addWorksheet(wb, "survey")
    openxlsx::addWorksheet(wb, "choices")
    openxlsx::writeData(wb, "survey", survey)
    openxlsx::writeData(wb, "choices", choices)

    if (isTRUE(paint)) {
      # Colores pastel: SM (verde), SO (azul), INTEGER (morado)
      verde_s  <- "#DFF5DF"
      azul_s   <- "#DCEBFF"
      morado_s <- "#E6D9F2"

      verde_c  <- "#EFFAEF"
      azul_c   <- "#EEF5FF"
      morado_c <- "#F2E6FF"

      # survey: localizar filas por NOMBRE (columna name)
      rows_sm  <- which(survey$name %in% unique(new_names_sm))
      rows_so  <- which(survey$name %in% unique(new_names_so))
      rows_int <- which(survey$name %in% unique(new_names_int))

      # en Excel, fila 1 = encabezados → sumar 1
      .paint_new(wb, "survey", rows_sm  + 1L, ncol(survey), verde_s)
      .paint_new(wb, "survey", rows_so  + 1L, ncol(survey), azul_s)
      .paint_new(wb, "survey", rows_int + 1L, ncol(survey), morado_s)

      # choices: pintar por list_name nuevas
      paint_choices <- function(list_names, hex){
        list_names <- unique(list_names)
        if (!length(list_names)) return()
        ln_idx <- which(choices$list_name %in% list_names)
        if (length(ln_idx)) {
          openxlsx::addStyle(wb, "choices", .style(hex),
                             rows = ln_idx + 1L, cols = 1:ncol(choices),
                             gridExpand = TRUE, stack = TRUE)
        }
      }

      paint_choices(unique(new_lists_sm),  verde_c)
      paint_choices(unique(new_lists_so),  azul_c)
      paint_choices(unique(new_lists_int), morado_c)
    }

    openxlsx::freezePane(wb, "survey",  firstActiveRow = 2)
    openxlsx::freezePane(wb, "choices", firstActiveRow = 2)
    openxlsx::setColWidths(wb, "survey",  cols = 1:ncol(survey),  widths = "auto")
    openxlsx::setColWidths(wb, "choices", cols = 1:ncol(choices), widths = "auto")
    openxlsx::saveWorkbook(wb, path_instrumento_out, overwrite = TRUE)
  }

  invisible(list(
    survey   = survey,
    choices  = choices,
    out_path = path_instrumento_out
  ))
}
