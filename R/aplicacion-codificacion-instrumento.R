# =====================================================================
# ppra_adaptar_instrumento: genera un XLSForm adaptado con *_recod
# =====================================================================
suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(stringr)
  library(tibble)
  library(tidyr)
})

# --- helpers breves ----
`%||%` <- function(a,b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a

.ppra_label_col_guess <- function(df){
  nms <- tolower(names(df))
  hit <- match(TRUE, nms %in% c(
    "label::spanish (es)","label::spanish(es)","label::spanish_es",
    "label_spanish_es","label::spanish","label","label::es"
  ))
  if (is.na(hit)) "label" else names(df)[hit]
}

.ppra_choices_label_col_guess <- function(df){
  nms <- tolower(names(df))
  hit <- match(TRUE, nms %in% c(
    "label::spanish (es)","label::spanish(es)","label::spanish_es",
    "label_spanish_es","label::spanish","label","label::es"
  ))
  if (is.na(hit)) "label" else names(df)[hit]
}

.ppra_norm_tokens <- function(x){
  x <- as.character(x)
  x <- gsub("[,;|/]+"," ", x, perl = TRUE) # separadores extra
  x <- gsub("\\s+"," ", x, perl = TRUE)
  x <- trimws(x)
  x[x==""] <- NA_character_
  x
}

.ppra_split_tokens <- function(v){
  v <- .ppra_norm_tokens(v)
  strsplit(ifelse(is.na(v),"",v), "\\s+")
}

.ppra_sanitize_listname <- function(x){
  x <- gsub("[^A-Za-z0-9_]+","_", x)  # todo lo no [A-Za-z0-9_] -> "_"
  x <- gsub("_+","_", x)              # colapsa múltiples "_"
  x <- sub("^_+","", x)               # quita "_" al inicio
  x <- sub("_+$","", x)               # quita "_" al final
  if (!nzchar(x)) x <- "ppra_list"
  x
}

.ppra_row_after <- function(df, row_idx, new_row){
  if (is.na(row_idx) || row_idx<=0 || row_idx>=nrow(df)) {
    bind_rows(df, new_row)
  } else {
    bind_rows(
      df[seq_len(row_idx), , drop = FALSE],
      new_row,
      df[(row_idx+1):nrow(df), , drop = FALSE]
    )
  }
}

.ppra_extract_listname_from_type <- function(type_str){
  # ej: "select_multiple mylist", "select_one mylist"
  s <- as.character(type_str)
  s <- trimws(s)
  parts <- strsplit(s, "\\s+")[[1]]
  if (!length(parts)) return(NA_character_)
  if (length(parts)>=2 && parts[1] %in% c("select_one","select_multiple"))
    parts[2] else NA_character_
}


.ppra_last_idx_choices_list <- function(choices, list_name){
  hit <- which(choices$list_name == list_name)
  if (length(hit)) max(hit) else NA_integer_
}

# --- colores (openxlsx) ----
.ppra_style_fill <- function(hex){
  openxlsx::createStyle(fgFill = hex)
}

#' Adaptar un XLSForm con variables recodificadas
#'
#' @description
#' Toma un XLSForm de entrada (`survey`/`choices`) y una base **ya adaptada**
#' (salida de \code{ppra_adaptar_data}) y genera un instrumento actualizado que
#' incluye las nuevas variables recodificadas:
#' \itemize{
#'   \item Para \emph{select_multiple}: agrega \code{<parent>_recod}.
#'   \item Para \emph{select_one}: agrega \code{<parent>_recod} (y opcionalmente su label).
#'   \item Para abiertas (\code{open_var}): se decide si su \code{<detail>_recod}
#'         se comporta como \emph{select_one} o \emph{select_multiple} según el
#'         número de tokens (separados por espacio) en la columna \code{*_recod}.
#' }
#' Las nuevas preguntas se insertan en \code{survey} inmediatamente **después**
#' de la original y las opciones nuevas se añaden en \code{choices} (si faltan
#' labels para opciones nuevas, se dejan vacíos para completarlos manualmente).
#'
#' @param path_instrumento_in Ruta al XLSForm original (XLSX con hojas \code{survey} y \code{choices}).
#' @param path_data_adaptada Ruta al XLSX con la data adaptada (o un \code{data.frame}). Debe contener solo
#'   las columnas padre \code{*_recod} (por ejemplo, salida \code{final_data} de \code{ppra_adaptar_data}).
#' @param path_instrumento_out Ruta de salida para el XLSForm adaptado (XLSX).
#' @param parent_var Vector de \emph{select_multiple} a reflejar como \code{<parent>_recod}.
#' @param uni_var Vector de \emph{select_one} a reflejar como \code{<parent>_recod}.
#' @param open_var Vector de variables abiertas \code{<detail>} cuya \code{<detail>_recod} debe incorporarse.
#' @param labels_append_recode \code{logical}. Si \code{TRUE}, en \code{choices} se intentan replicar las
#'   etiquetas de las opciones existentes para las nuevas \code{*_recod}; las desconocidas se crean con label vacío.
#' @param paint \code{logical}. Si \code{TRUE}, colorea las filas nuevas en \code{survey}/\code{choices}
#'   según el tipo: \emph{select_multiple} (verde pastel) y \emph{select_one} (azul pastel).
#' @param choices_order Cadena con el orden de mezcla en \code{choices}: \code{"original_first"},
#'   \code{"recode_first"} o \code{"append_only"}.
#'
#' @details
#' \strong{Determinación del tipo para \code{open_var}}:
#' si la columna \code{<detail>_recod} tiene, por fila, más de un código (tokens
#' separados por espacio), se clasifica como \emph{select_multiple}; si la gran
#' mayoría son de un único token, se considera \emph{select_one}.
#'
#' \strong{List names}: para \emph{select_multiple} se usa el \code{list_name}
#' del \code{parent} original; para nuevas opciones se añade el \code{name}
#' (código) y se deja el \code{label} vacío cuando no exista.
#'
#' @return (Invisiblemente) una lista con:
#' \itemize{
#'   \item \code{survey}: \code{tibble} de la hoja \code{survey} adaptada.
#'   \item \code{choices}: \code{tibble} de la hoja \code{choices} adaptada.
#'   \item \code{path}: ruta escrita del archivo XLSX de salida.
#' }
#'
#' @seealso \code{\link{ppra_adaptar_data}}, \code{\link{ppra_auditar}}
#' @examples
#' \dontrun{
#' res_inst <- ppra_adaptar_instrumento(
#'   path_instrumento_in  = "instrumento_pdm.xlsx",
#'   path_data_adaptada   = "DATA_ADAPTADA_SM_FINAL.xlsx",
#'   path_instrumento_out = "instrumento_pdm_ADAPTADO.xlsx",
#'   parent_var = c("Cash_purpose","Who_gave_help"),
#'   uni_var    = c("Possession_payment_mechanism"),
#'   open_var   = c("Increase_in_prices_details"),
#'   labels_append_recode = TRUE,
#'   paint = TRUE,
#'   choices_order = "original_first"
#' )
#' }
#' @export
ppra_adaptar_instrumento <- function(path_instrumento_in,
                                     path_data_adaptada,
                                     path_instrumento_out = "instrumento_adaptado.xlsx",
                                     parent_var = NULL,
                                     uni_var    = NULL,
                                     open_var   = NULL,
                                     labels_append_recode = TRUE,
                                     paint = TRUE,
                                     choices_order = c("original_first","by_first_seen","alphabetical")){
  stopifnot(file.exists(path_instrumento_in))
  choices_order <- match.arg(choices_order)

  # === leer instrumento ===
  survey  <- readxl::read_excel(path_instrumento_in, sheet = "survey")
  choices <- readxl::read_excel(path_instrumento_in, sheet = "choices")
  if (!all(c("type","name") %in% names(survey))) stop("La hoja 'survey' debe tener columnas 'type' y 'name'.")
  if (!all(c("list_name","name") %in% names(choices))) stop("La hoja 'choices' debe tener 'list_name' y 'name'.")

  lab_col_s <- .ppra_label_col_guess(survey)
  lab_col_c <- .ppra_choices_label_col_guess(choices)

  survey$name  <- as.character(survey$name)
  choices$name <- as.character(choices$name)
  choices$list_name <- as.character(choices$list_name)

  # === leer data adaptada ===
  if (is.character(path_data_adaptada) && file.exists(path_data_adaptada)) {
    df <- readxl::read_xlsx(path_data_adaptada)
  } else if (is.data.frame(path_data_adaptada)) {
    df <- path_data_adaptada
  } else {
    stop("path_data_adaptada debe ser data.frame o ruta a un XLSX.")
  }

  # === universo por data si no se pasa explícito ===
  nms <- names(df)
  if (is.null(parent_var)) {
    # padres SM detectables por <parent>_recod y existencia de <parent>/child en original: dejemos lo “explícito”
    parent_var <- unique(sub("_recod$","", grep("^[^/]+_recod$", nms, value = TRUE)))
  }
  if (is.null(uni_var)) {
    # también podrían estar en lo anterior; los distinguimos por XLSForm al copiar list_name original
    uni_var <- character(0)  # se decidirá cuando leamos el type original
  }
  if (is.null(open_var)) open_var <- character(0)

  # === índice de labels de choices por list_name ===
  ch_index <- choices %>% select(list_name, code = name, label = all_of(lab_col_c))

  # === log de lo creado ===
  log_rows <- list()


  # === función interna: agregar una pregunta recod SM/SO ===
  .add_recoded_question <- function(survey, choices, base_name, kind = c("multiple","one"),
                                    tokens_from_data = character(0),
                                    from_list = NULL,
                                    note = NULL){
    kind <- match.arg(kind)

    # fila base en survey
    idx <- match(base_name, survey$name)
    if (is.na(idx)) {
      # si no existe, la insertamos al final y avisamos
      idx <- nrow(survey)
      base_label <- base_name
      base_type  <- if (kind=="multiple") "select_multiple" else "select_one"
    } else {
      base_label <- as.character(survey[[lab_col_s]][idx] %||% base_name)
      base_type  <- as.character(survey$type[idx])
    }

    # list_name original (si existe)
    ln_orig <- .ppra_extract_listname_from_type(base_type)

    # === CAMBIO 1: decidir list_name nuevo como <lista_original>_recod o <base_saneado>_recod ===
    base_list     <- if (!is.na(ln_orig)) ln_orig else .ppra_sanitize_listname(base_name)
    list_name_new <- paste0(base_list, "_recod")

    # armar codes + labels
    codes <- character(0); labels <- character(0)

    # 1) si hay lista original y el tipo original era compatible, copiamos catálogo original
    if (!is.na(ln_orig) && ln_orig %in% choices$list_name) {
      orig <- choices %>% dplyr::filter(list_name == ln_orig)
      codes  <- c(codes, as.character(orig$name))
      labels <- c(labels, as.character(orig[[lab_col_c]]))
    }

    # 2) sumar tokens observados en la data adaptada
    if (length(tokens_from_data)) {
      seen <- unique(tokens_from_data[nzchar(tokens_from_data)])
      new_codes <- setdiff(seen, codes)
      if (length(new_codes)) {
        codes  <- c(codes, new_codes)
        labels <- c(labels, rep(NA_character_, length(new_codes)))
      }
    }

    # orden final
    if (choices_order == "alphabetical") {
      o <- order(codes)
      codes <- codes[o]; labels <- labels[o]
    } else if (choices_order == "by_first_seen") {
      # ya vienen en orden visto (orig + vistos) -> no tocar
    } else {
      # original_first -> ya cumplido
    }

    # construir fila nueva en survey
    new_type <- if (kind=="multiple") paste("select_multiple", list_name_new)
    else                    paste("select_one",      list_name_new)

    # === CAMBIO 2: name nuevo como <base>_recod (no "recod_<base>") ===
    new_name <- paste0(base_name, "_recod")

    new_row  <- survey[0,]
    new_row[1, setdiff(names(survey), character(0))] <- NA
    new_row$type <- new_type
    new_row$name <- new_name
    new_row[[lab_col_s]] <- if (isTRUE(labels_append_recode)) paste0(base_label, " [recod]") else base_label

    # insertar debajo de la base
    survey2 <- .ppra_row_after(survey, idx, new_row)

    # crear choices del list_name nuevo
    # Inserta justo debajo del list_name original si existe; si no, al final
    if (length(codes)) {
      add_choices <- tibble::tibble(list_name = list_name_new, name = codes)
      add_choices[[lab_col_c]] <- labels

      below <- .ppra_last_idx_choices_list(choices, ln_orig)  # ln_orig es la lista original
      if (!is.na(below)) {
        if (below >= nrow(choices)) {
          choices2 <- dplyr::bind_rows(choices, add_choices)
        } else {
          choices2 <- dplyr::bind_rows(
            choices %>% dplyr::slice(1:below),
            add_choices,
            choices %>% dplyr::slice((below+1):dplyr::n())
          )
        }
      } else {
        # si no hay list original, lo agregamos al final
        choices2 <- dplyr::bind_rows(choices, add_choices)
      }
    } else {
      choices2 <- choices
    }

    # log
    log_rows[[length(log_rows)+1]] <<- tibble::tibble(
      var = new_name,
      kind = kind,
      list_name_new = list_name_new,
      from_list = ln_orig %||% NA_character_,
      n_codes_orig = if (!is.na(ln_orig) && ln_orig %in% choices$list_name)
        nrow(choices %>% dplyr::filter(list_name==ln_orig)) else 0L,
      n_codes_new = length(setdiff(codes, choices %>% dplyr::filter(list_name==ln_orig) %>% dplyr::pull(name) %||% character(0))),
      note = note %||% NA_character_
    )

    list(survey = survey2, choices = choices2, new_name = new_name, list_name_new = list_name_new)
  }

  # === 1) SELECT-MULTIPLE y SELECT-ONE explícitos (usando survey/type para decidir) ===
  #    Si “uni_var” viene vacío, intentamos inferir por type original
  all_targets <- unique(c(parent_var, uni_var))

  for (pv in all_targets) {
    # tokens desde data: col <pv>_recod si existe
    tks <- character(0)
    pr <- paste0(pv, "_recod")
    if (pr %in% nms) {
      toks_list <- .ppra_split_tokens(df[[pr]])
      tks <- unique(unlist(toks_list))
      tks <- tks[nzchar(tks)]
    }

    # tipo original del survey
    idx <- match(pv, survey$name)
    if (!is.na(idx)) {
      ty <- as.character(survey$type[idx])
      base <- strsplit(ty, "\\s+")[[1]]
      base <- if (length(base)) base[1] else NA_character_
    } else {
      base <- NA_character_
    }

    if (!is.na(base) && base == "select_multiple") {
      res <- .add_recoded_question(survey, choices, base_name = pv, kind = "multiple",
                                   tokens_from_data = tks, note = "SM_from_survey")
    } else {
      # si el survey dice select_one o no hay type, construimos SO
      res <- .add_recoded_question(survey, choices, base_name = pv, kind = "one",
                                   tokens_from_data = tks, note = if (!is.na(base)) "SO_from_survey" else "SO_default")
    }
    survey <- res$survey; choices <- res$choices
  }

  # === 2) OPEN: inferir tipo por data y agregar al lado del detail ===
  for (dv in open_var) {
    rec_col <- paste0(dv, "_recod")
    if (!rec_col %in% nms) {
      # nada que hacer si no existe recod en la data
      next
    }
    # inferir tipo mirando n_tokens
    v <- .ppra_norm_tokens(df[[rec_col]])
    toks <- strsplit(ifelse(is.na(v),"",v), "\\s+")
    n_tok <- vapply(toks, function(z) sum(nzchar(z)), integer(1))
    kind <- if (max(n_tok, na.rm = TRUE) > 1) "multiple" else "one"
    if (all(is.na(v))) {
      kind <- "one"
      note_kind <- "open_empty_assumed_one"
    } else note_kind <- paste0("open_", kind, "_by_data")

    # tokens vistos
    tks <- unique(unlist(toks)); tks <- tks[nzchar(tks)]

    # insertar debajo del detail (fila text)
    idx <- match(dv, survey$name)
    if (is.na(idx)) {
      # si no existe, caeremos en “después” del final
      idx <- nrow(survey)
    }

    # parent_guess: fila anterior a la del detail (si existe)
    det_idx <- match(dv, survey$name)
    parent_guess <- if (!is.na(det_idx) && det_idx > 1) survey$name[det_idx - 1] else NA_character_
    # list_name del parent (si existiese)
    ln_parent <- if (!is.na(parent_guess)) {
      i_par <- match(parent_guess, survey$name)
      if (!is.na(i_par)) .ppra_extract_listname_from_type(survey$type[i_par]) else NA_character_
    } else NA_character_

    # Llamada normal (la función decidirá <base_list>_recod con ln_parent o con el propio detail saneado)
    res <- .add_recoded_question(survey, choices, base_name = dv, kind = kind,
                                 tokens_from_data = tks, note = note_kind)
    survey <- res$survey; choices <- res$choices
  }

  # === pintar y exportar ===
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    warning("No se encontró 'openxlsx'. Se guardará sin color.")
    openxlsx::write.xlsx(list(survey = survey, choices = choices), path_instrumento_out, overwrite = TRUE)
  } else {
    wb <- openxlsx::createWorkbook()
    openxlsx::addWorksheet(wb, "survey")
    openxlsx::addWorksheet(wb, "choices")
    openxlsx::writeData(wb, "survey", survey)
    openxlsx::writeData(wb, "choices", choices)

    if (isTRUE(paint)) {
      # Colores
      verde_s  <- "#DFF5DF"; verde_c  <- "#EFFAEF"  # multiple
      azul_s   <- "#DCEBFF"; azul_c   <- "#EEF5FF"  # one

      # pintar filas *_recod en survey según tipo por su type
      type_vec <- as.character(survey$type)
      name_vec <- as.character(survey$name)
      is_re    <- grepl("_recod$", name_vec)
      rows_re  <- which(is_re)

      for (r in rows_re) {
        ty <- strsplit(type_vec[r], "\\s+")[[1]]
        base <- if (length(ty)) ty[1] else ""
        if (identical(base, "select_multiple")) {
          openxlsx::addStyle(wb, "survey", .ppra_style_fill(verde_s),
                             rows = r+1, cols = 1:ncol(survey), gridExpand = TRUE, stack = TRUE)
          # también sus choices (list_name en col “type” pos2)
          ln <- .ppra_extract_listname_from_type(type_vec[r])
          if (!is.na(ln)) {
            crow <- which(as.character(choices$list_name) == ln)
            if (length(crow)) {
              openxlsx::addStyle(wb, "choices", .ppra_style_fill(verde_c),
                                 rows = crow+1, cols = 1:ncol(choices), gridExpand = TRUE, stack = TRUE)
            }
          }
        } else if (identical(base, "select_one")) {
          openxlsx::addStyle(wb, "survey", .ppra_style_fill(azul_s),
                             rows = r+1, cols = 1:ncol(survey), gridExpand = TRUE, stack = TRUE)
          ln <- .ppra_extract_listname_from_type(type_vec[r])
          if (!is.na(ln)) {
            crow <- which(as.character(choices$list_name) == ln)
            if (length(crow)) {
              openxlsx::addStyle(wb, "choices", .ppra_style_fill(azul_c),
                                 rows = crow+1, cols = 1:ncol(choices), gridExpand = TRUE, stack = TRUE)
            }
          }
        }
      }
    }

    openxlsx::freezePane(wb, "survey", firstActiveRow = 2)
    openxlsx::freezePane(wb, "choices", firstActiveRow = 2)
    openxlsx::setColWidths(wb, "survey",  cols = 1:ncol(survey),  widths = "auto")
    openxlsx::setColWidths(wb, "choices", cols = 1:ncol(choices), widths = "auto")
    openxlsx::saveWorkbook(wb, path_instrumento_out, overwrite = TRUE)
  }

  out_log <- dplyr::bind_rows(log_rows) %>% arrange(var)

  list(
    survey   = survey,
    choices  = choices,
    log      = out_log,
    out_path = path_instrumento_out
  )
}
