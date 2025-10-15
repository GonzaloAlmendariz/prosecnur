# =============================================================================
# ppra_adaptar_instrumento: crea un XLSForm adaptado compatible con el flujo "simple"
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

  # list nuevo:
  # - SM y SO-padre: <ln_orig>_recod si hay ln_orig, si no: <sanitize(base)>_recod
  # - SO-hijo (usaremos esta misma función también): si se pasa list_name_hint,
  #   lo usamos; si no, generamos <sanitize(new_name)>_list.
  base_list <- if (!is.null(list_name_hint)) list_name_hint else {
    if (!is.na(ln_orig)) paste0(ln_orig, "_recod") else paste0(.sanitize(base_name), "_recod")
  }

  # nombre nuevo
  new_name <- paste0(base_name, "_recod")
  new_type <- if (kind=="multiple") paste("select_multiple", base_list) else paste("select_one", base_list)

  # catálogo a crear para el list_name destino
  codes <- character(0); labels <- character(0)

  # si ln_orig existe y coincidimos con SO-padre/SM, copiamos catálogo original
  if (is.null(list_name_hint) && !is.na(ln_orig) && ln_orig %in% choices$list_name) {
    orig <- choices %>% filter(list_name == ln_orig)
    codes  <- c(codes, as.character(orig$name))
    labels <- c(labels, as.character(orig[[lab_col_c]]))
  }

  # añadir tokens observados en data
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
  } # original_first/by_first_seen: dejamos el orden construido

  # construir fila survey
  new_row <- survey[0,]; new_row[1, setdiff(names(survey), character(0))] <- NA
  new_row$type <- new_type
  new_row$name <- new_name
  new_row[[lab_col_s]] <- paste0(base_label, " [recod]")

  survey2 <- .row_after(survey, i_base, new_row)

  # inyectar choices del list destino: debajo del catálogo original si aplica, o al final
  if (length(codes)) {
    add_choices <- tibble(list_name = base_list, name = codes)
    add_choices[[lab_col_c]] <- labels

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
       base_row = if (is.na(i_base)) nrow(survey2) else i_base + 1)  # +1 por la inserción
}

# =============================================================================
#' @title ppra_adaptar_instrumento
#' @description
#' Genera un **XLSForm adaptado** desde un instrumento base y una **data ya adaptada**
#' (salida de `ppra_adaptar_data_simple`), creando:
#' - **SM**: `<parent>_recod` (select_multiple).
#' - **SO-padre**: `<parent>_recod` (select_one) usando el catálogo del padre.
#' - **SO-hijo**: `<parent>_<alias>_recod` (select_one), con **lista propia**
#'   (tokens observados en la data).
#' Inserta cada nueva pregunta **debajo** de su base y **colorea**:
#' SM (verde), SO (azul). También colorea las filas de `choices` del list creado.
#' @param path_instrumento_in XLSX original (hojas `survey` y `choices`)
#' @param path_data_adaptada XLSX o `data.frame` con columnas *_recod generadas
#' @param path_instrumento_out XLSX de salida
#' @param sm_vars vector de padres select_multiple
#' @param so_parent_vars vector de select_one que recodifican el **padre**
#' @param so_child_vars vector de select_one cuyo recodificado está en el **hijo**;
#'   la/las columnas a agregar se detectan en la data como `^<parent>_.*_recod$`
#'   (excluyendo `<parent>_recod`)
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

  survey$name  <- as.character(survey$name)
  choices$name <- as.character(choices$name)
  choices$list_name <- as.character(choices$list_name)

  # --- leer data adaptada ---
  df <- if (is.character(path_data_adaptada) && file.exists(path_data_adaptada)) {
    readxl::read_xlsx(path_data_adaptada)
  } else if (is.data.frame(path_data_adaptada)) {
    path_data_adaptada
  } else stop("path_data_adaptada debe ser data.frame o ruta a XLSX con la data adaptada.")

  nms <- names(df)

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
      if (!col_rec %in% nms) next
      toks <- unique(unlist(.split_tokens(df[[col_rec]])))
      toks <- toks[nzchar(toks)]
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
      new_rows_sm  <- c(new_rows_sm, res$base_row + 1)  # la línea insertada justo después
      new_lists_sm <- c(new_lists_sm, res$list_name)
    }
  }

  # =======================
  # 2) SELECT ONE (PADRE)
  # =======================
  if (length(so_parent_vars)) {
    for (pv in so_parent_vars) {
      col_rec <- paste0(pv, "_recod")
      toks <- if (col_rec %in% nms) {
        v <- .split_tokens(df[[col_rec]])
        unique(unlist(v))
      } else character(0)

      res <- .add_recoded_q(survey, choices,
                            base_name = pv,
                            kind      = "one",
                            list_name_hint = NULL,           # usa <ln_orig>_recod si existe
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
      # detectar las columnas hijas *_recod: ^<parent>_.*_recod$ y != <parent>_recod
      rx <- paste0("^", gsub("([\\W])","\\\\\\1", pv), "_.+_recod$")
      child_cols <- grep(rx, nms, value = TRUE)
      child_cols <- setdiff(child_cols, paste0(pv, "_recod"))
      if (!length(child_cols)) next

      for (cc in child_cols) {
        # p.ej. CEPR01_why_recod -> base_name = ese mismo (tratamos cada hijo como “base” independiente)
        toks <- unique(unlist(.split_tokens(df[[cc]])))
        toks <- toks[nzchar(toks)]

        # Lista PROPIA para el hijo: <sanitize(cc)>_list
        list_hint <- paste0(.sanitize(cc), "_list")

        res <- .add_recoded_q(survey, choices,
                              base_name = cc,       # la pregunta nueva se llama == columna hija *_recod
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
        ln_idx <- which(choices$list_name %in% list_names)
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
