# =============================================================================
# ppra_adaptar_instrumento — ahora lee tokens desde TODAS las hojas de la data
# (main + repeats), sin romper el comportamiento anterior.
# =============================================================================
suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(stringr)
  library(tidyr)
  library(tibble)
})

`%||%` <- function(a,b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a

# ---------- helpers básicos ----------
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
    bind_rows(df, new_row)
  } else {
    bind_rows(df[seq_len(row_idx), , drop = FALSE],
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

# ---------- lectura multi-hojas ----------
.read_all_sheets <- function(path_xlsx){
  sh <- readxl::excel_sheets(path_xlsx)
  setNames(lapply(sh, function(s){
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
      v <- d[[col_name]]
      vv <- unique(unlist(.split_tokens(v)))
      toks <- c(toks, vv)
    }
  }
  unique(toks[nzchar(toks)])
}

.collect_child_cols <- function(df_or_path, parent){
  # Devuelve nombres de columnas hijas *_recod a lo largo de todas las hojas:
  # ^<parent>_.+_recod$, excluyendo <parent>_recod
  rx  <- paste0("^", gsub("([\\W])","\\\\\\1", parent), "_.+_recod$")
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

# ---------- pintura ----------
.style <- function(hex) openxlsx::createStyle(fgFill = hex)

.paint_new <- function(wb, sheet, rows, ncols, hex){
  if (!length(rows)) return()
  openxlsx::addStyle(wb, sheet, .style(hex),
                     rows = rows, cols = 1:ncols,
                     gridExpand = TRUE, stack = TRUE)
}

# ---------- núcleo: insertar una pregunta *_recod debajo de su base ----------
.add_recoded_q <- function(survey, choices, base_name, kind = c("multiple","one"),
                           list_name_hint = NULL,
                           tokens_from_data = character(0),
                           lab_col_s, lab_col_c,
                           choices_order = c("original_first","by_first_seen","alphabetical")){
  kind <- match.arg(kind)
  choices_order <- match.arg(choices_order)

  # fila base (si existe)
  i_base <- match(base_name, survey$name)
  base_type  <- if (!is.na(i_base)) as.character(survey$type[i_base]) else if (kind=="multiple") "select_multiple" else "select_one"
  base_label <- if (!is.na(i_base)) as.character(survey[[lab_col_s]][i_base] %||% base_name) else base_name

  # list original (si existe)
  ln_orig <- .extract_listname(base_type)

  # list destino
  base_list <- if (!is.null(list_name_hint)) list_name_hint else {
    if (!is.na(ln_orig)) paste0(ln_orig, "_recod") else paste0(.sanitize(base_name), "_recod")
  }

  # nombre nuevo
  new_name <- paste0(base_name, "_recod")
  new_type <- if (kind=="multiple") paste("select_multiple", base_list) else paste("select_one", base_list)

  # catálogo a crear
  codes <- character(0); labels <- character(0)

  # copiar catálogo original si aplica
  if (is.null(list_name_hint) && !is.na(ln_orig) && ln_orig %in% choices$list_name) {
    orig <- choices %>% filter(list_name == ln_orig)
    codes  <- c(codes, as.character(orig$name))
    labcol <- .guess_label_col(orig)  # por si difiere
    labels <- c(labels, as.character(orig[[labcol]]))
  }

  # añadir tokens observados
  if (length(tokens_from_data)) {
    seen <- unique(tokens_from_data[nzchar(tokens_from_data)])
    new_codes <- setdiff(seen, codes)
    if (length(new_codes)) {
      codes  <- c(codes, new_codes)
      labels <- c(labels, rep(NA_character_, length(new_codes)))
    }
  }

  # ordenar catálogo
  if (choices_order == "alphabetical") {
    o <- order(codes); codes <- codes[o]; labels <- labels[o]
  }

  # fila survey
  new_row <- survey[0,]; new_row[1, setdiff(names(survey), character(0))] <- NA
  new_row$type <- new_type
  new_row$name <- new_name
  new_row[[lab_col_s]] <- paste0(base_label, " [recod]")

  survey2 <- .row_after(survey, i_base, new_row)

  # inyectar choices del list destino (evitando duplicar filas existentes)
  if (length(codes)) {
    lab_col_c <- .guess_label_col(choices)
    add_choices <- tibble(list_name = base_list, name = codes)
    add_choices[[lab_col_c]] <- labels

    # evitar duplicados exactos (list_name + name)
    dup_mask <- paste(choices$list_name, choices$name) %in% paste(add_choices$list_name, add_choices$name)
    if (any(dup_mask)) {
      choices <- choices[!dup_mask, , drop = FALSE]
    }

    # posición: debajo del catálogo original si existe, si no, al final
    below <- if (!is.na(ln_orig)) {
      hit <- which(choices$list_name == ln_orig)
      if (length(hit)) max(hit) else NA_integer_
    } else NA_integer_

    if (!is.na(below) && below < nrow(choices)) {
      choices2 <- bind_rows(
        choices %>% slice(1:below),
        add_choices,
        choices %>% slice((below+1):n())
      )
    } else {
      choices2 <- bind_rows(choices, add_choices)
    }
  } else {
    choices2 <- choices
  }

  list(survey = survey2, choices = choices2, new_name = new_name, list_name = base_list,
       base_row = if (is.na(i_base)) nrow(survey2) else i_base + 1)
}

# =============================================================================
#' @title ppra_adaptar_instrumento
#' @description
#' Genera un **XLSForm adaptado** desde un instrumento base y una **data adaptada**
#' (producida por tu flujo), creando:
#' - **SM**: `<parent>_recod` (select_multiple).
#' - **SO-padre**: `<parent>_recod` (select_one) usando catálogo del padre + tokens observados.
#' - **SO-hijo**: `<parent>_<alias>_recod` (select_one), lista propia a partir de tokens
#'   observados en **todas las hojas** donde aparezca.
#' Inserta cada nueva pregunta **debajo** de su base y colorea: SM (verde), SO (azul).
#' Lee **todas las hojas** del Excel de data adaptada; si es un data.frame, actúa como antes.
#' @param path_instrumento_in XLSX original (hojas `survey` y `choices`)
#' @param path_data_adaptada XLSX (multihoja) o `data.frame` con columnas *_recod
#' @param path_instrumento_out XLSX de salida
#' @param sm_vars vector de padres select_multiple
#' @param so_parent_vars vector de select_one que recodifican el **padre**
#' @param so_child_vars vector de select_one cuyo recodificado está en el **hijo**;
#'   se detectan columnas `^<parent>_.+_recod$` (excluyendo `<parent>_recod`) en TODAS las hojas
#' @param choices_order "original_first" (default), "by_first_seen", "alphabetical"
#' @param paint TRUE para colorear filas nuevas
#' @return lista con `survey`, `choices`, `out_path`
#' @export
# =============================================================================
ppra_adaptar_instrumento <- function(path_instrumento_in,
                                     path_data_adaptada,
                                     path_instrumento_out = "instrumento_adaptado.xlsx",
                                     sm_vars        = character(0),
                                     so_parent_vars = character(0),
                                     so_child_vars  = character(0),
                                     choices_order  = c("original_first","by_first_seen","alphabetical"),
                                     paint = TRUE){

  choices_order <- match.arg(choices_order)
  stopifnot(file.exists(path_instrumento_in))

  # --- leer instrumento base ---
  survey  <- readxl::read_excel(path_instrumento_in, sheet = "survey")
  choices <- readxl::read_excel(path_instrumento_in, sheet = "choices")
  if (!all(c("type","name") %in% names(survey)))  stop("survey debe tener 'type' y 'name'.")
  if (!all(c("list_name","name") %in% names(choices))) stop("choices debe tener 'list_name' y 'name'.")

  lab_col_s <- .guess_label_col(survey)
  lab_col_c <- .guess_label_col(choices)

  survey$name      <- as.character(survey$name)
  choices$name     <- as.character(choices$name)
  choices$list_name<- as.character(choices$list_name)

  # --- preparar acceso multi-hojas ---
  df_is_xlsx <- is.character(path_data_adaptada) && file.exists(path_data_adaptada)
  df_single  <- is.data.frame(path_data_adaptada)

  if (!df_is_xlsx && !df_single)
    stop("path_data_adaptada debe ser data.frame o ruta a XLSX con la data adaptada.")

  # --- logs de filas nuevas en survey (para pintar) ---
  new_rows_sm  <- integer(0)
  new_rows_so  <- integer(0)
  new_lists_sm <- character(0)
  new_lists_so <- character(0)

  # =======================
  # 1) SELECT MULTIPLE
  # =======================
  if (length(sm_vars)) {
    for (pv in sm_vars) {
      col_rec <- paste0(pv, "_recod")

      # tokens desde TODAS las hojas
      toks <- .collect_tokens_from_col(path_data_adaptada, col_rec)

      res <- .add_recoded_q(survey, choices,
                            base_name = pv,
                            kind      = "multiple",
                            list_name_hint = NULL,           # usa <ln_orig>_recod si existe
                            tokens_from_data = toks,
                            lab_col_s = lab_col_s,
                            lab_col_c = lab_col_c,
                            choices_order = choices_order)
      survey  <- res$survey
      choices <- res$choices
      new_rows_sm  <- c(new_rows_sm, res$base_row + 1)
      new_lists_sm <- c(new_lists_sm, res$list_name)
    }
  }

  # =======================
  # 2) SELECT ONE (PADRE)
  # =======================
  if (length(so_parent_vars)) {
    for (pv in so_parent_vars) {
      col_rec <- paste0(pv, "_recod")
      toks <- .collect_tokens_from_col(path_data_adaptada, col_rec)

      res <- .add_recoded_q(survey, choices,
                            base_name = pv,
                            kind      = "one",
                            list_name_hint = NULL,
                            tokens_from_data = toks,
                            lab_col_s = lab_col_s,
                            lab_col_c = lab_col_c,
                            choices_order = choices_order)
      survey  <- res$survey
      choices <- res$choices
      new_rows_so <- c(new_rows_so, res$base_row + 1)
      new_lists_so <- c(new_lists_so, res$list_name)
    }
  }

  # =======================
  # 3) SELECT ONE (HIJO)
  # =======================
  if (length(so_child_vars)) {
    for (pv in so_child_vars) {
      # columnas hijas en TODAS las hojas
      child_cols <- .collect_child_cols(path_data_adaptada, pv)
      if (!length(child_cols)) next

      for (cc in child_cols) {
        toks <- .collect_tokens_from_col(path_data_adaptada, cc)
        toks <- toks[nzchar(toks)]

        # Lista PROPIA para el hijo: <sanitize(cc)>_list
        list_hint <- paste0(.sanitize(cc), "_list")

        # Nota: por compatibilidad con tu flujo anterior, mantenemos
        # base_name = cc (esto generará cc_recod como nombre nuevo).
        res <- .add_recoded_q(survey, choices,
                              base_name = cc,
                              kind      = "one",
                              list_name_hint = list_hint,
                              tokens_from_data = toks,
                              lab_col_s = lab_col_s,
                              lab_col_c = lab_col_c,
                              choices_order = choices_order)
        survey  <- res$survey
        choices <- res$choices
        new_rows_so <- c(new_rows_so, res$base_row + 1)
        new_lists_so <- c(new_lists_so, res$list_name)
      }
    }
  }

  # =======================
  # 4) Exportar + colorear
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
      # Colores pastel: SM (verde), SO (azul)
      verde_s <- "#DFF5DF"; verde_c <- "#EFFAEF"
      azul_s  <- "#DCEBFF"; azul_c  <- "#EEF5FF"

      # pintar survey (filas nuevas)
      .paint_new(wb, "survey", new_rows_sm, ncol(survey), verde_s)
      .paint_new(wb, "survey", new_rows_so, ncol(survey), azul_s)

      # pintar choices por list_name creados
      paint_choices <- function(list_names, hex){
        if (!length(list_names)) return()
        ln_idx <- which(choices$list_name %in% unique(list_names))
        if (length(ln_idx)) {
          openxlsx::addStyle(wb, "choices", .style(hex),
                             rows = ln_idx + 1, cols = 1:ncol(choices),
                             gridExpand = TRUE, stack = TRUE)
        }
      }
      paint_choices(unique(new_lists_sm), verde_c)
      paint_choices(unique(new_lists_so), azul_c)
    }

    openxlsx::freezePane(wb, "survey", firstActiveRow = 2)
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
