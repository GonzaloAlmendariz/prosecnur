# =============================================================================
# PPRA – Adaptación de datos con soporte para repeats y export preservando hojas
# (con armonización de clave para joins: evita _index double vs character)
# =============================================================================

suppressPackageStartupMessages({
  library(readxl); library(dplyr); library(stringr); library(purrr); library(tidyr); library(tibble)
})

`%||%` <- function(a,b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a
nz      <- function(x) !is.na(x) & nzchar(trimws(as.character(x)))

# -------- utilidades básicas ---------------------------------------------------
pick_join_key <- function(df){
  cands <- c("_uuid","uuid","meta_instance_id","instanceid","_id",
             "_index","Codigo pulso","Código pulso","Pulso_code","pulso_code")
  hit <- cands[cands %in% names(df)]
  if (length(hit)) hit[1] else NA_character_
}

# elige una clave que exista en x e y (en orden de prioridad)
pick_join_key_pair <- function(x, y){
  cands <- c("_uuid","uuid","meta_instance_id","instanceid","_id",
             "_index","Codigo pulso","Código pulso","Pulso_code","pulso_code")
  for (k in cands) {
    if ((k %in% names(x)) && (k %in% names(y))) return(k)
  }
  NA_character_
}

# asegura nombres únicos (evita "Input columns in `y` must be unique")
.ensure_unique_names <- function(df){
  names(df) <- make.unique(names(df), sep = "__")
  df
}

# ---- armonización de clave para joins (ambos lados a texto) ------------------
.to_key_char <- function(v){
  if (inherits(v, "POSIXct") || inherits(v, "Date")) return(as.character(v))
  if (is.numeric(v)) return(ifelse(is.na(v), NA_character_, format(v, trim = TRUE, scientific = FALSE, digits = 22)))
  as.character(v)
}

.harmonize_key <- function(df, key){
  if (!key %in% names(df)) return(df)
  df[[key]] <- .to_key_char(df[[key]])
  df
}

.safe_left_join_by <- function(x, y, key, cols_right = NULL){
  y <- .ensure_unique_names(y)
  x <- .harmonize_key(x, key)
  y <- .harmonize_key(y, key)
  if (!is.null(cols_right)) {
    keep <- unique(c(key, cols_right))
    keep <- keep[keep %in% names(y)]
    y <- y[, keep, drop = FALSE]
  }
  dplyr::left_join(x, y, by = key)
}

# normalizar valores 0/1/NA (NO nombres de columnas)
.norm01 <- function(v){
  if (is.numeric(v)) return(ifelse(is.na(v), NA_integer_, ifelse(v!=0,1L,0L)))
  s <- tolower(trimws(as.character(v)))
  s <- iconv(s, from="", to="ASCII//TRANSLIT")
  out <- rep(NA_integer_, length(s))
  out[s %in% c("1","t","true","si","s","yes","y","verdadero")] <- 1L
  out[s %in% c("0","f","false","no","n","falso")]               <- 0L
  out
}

# normaliza códigos (valores), NO se usa en nombres de columnas
.normcode <- function(x){
  x <- trimws(as.character(x))
  x <- iconv(x, from="", to="ASCII//TRANSLIT")
  tolower(x)
}

leer_datos_generico <- function(path_data, sheet = NULL){
  ext <- tolower(tools::file_ext(path_data))
  if (ext %in% c("csv","txt")) {
    suppressWarnings(readr::read_csv(path_data, show_col_types = FALSE))
  } else {
    readxl::read_excel(path_data, sheet = sheet %||% 1)
  }
}

# leer hoja con nombre insensible a mayúsculas
read_sheet_ci <- function(path_xlsx, sheet_name){
  sh <- readxl::excel_sheets(path_xlsx)
  i  <- match(tolower(sheet_name), tolower(sh))
  if (is.na(i)) stop("No existe la hoja '", sheet_name, "' en: ", path_xlsx)
  df <- readxl::read_excel(path_xlsx, sheet = sh[i])
  .ensure_unique_names(df)
}

# -------- plantilla: resolver hoja --------------------------------------------
.ppra_clean <- function(x) gsub("[^A-Za-z0-9]+", "", tolower(trimws(as.character(x))))
ppra_resolve_template_sheet <- function(path_plantilla, parent_var){
  sh <- readxl::excel_sheets(path_plantilla)
  hit <- which(sh == parent_var); if (length(hit)) return(sh[hit[1]])
  hit <- which(tolower(sh) == tolower(parent_var)); if (length(hit)) return(sh[hit[1]])
  p31 <- substr(parent_var, 1, 31)
  hit <- which(sh == p31 | tolower(sh) == tolower(p31)); if (length(hit)) return(sh[hit[1]])
  cl_parent <- .ppra_clean(parent_var); cl_sheets <- .ppra_clean(sh)
  hit <- which(cl_sheets == cl_parent); if (length(hit)) return(sh[hit[1]])
  d <- adist(cl_parent, cl_sheets); j <- which.min(d)
  if (length(j) && is.finite(d[j]) && d[j] <= 5) return(sh[j])
  NA_character_
}

# -------- localizar hoja de datos donde vive el parent -------------------------
locate_var_sheet <- function(parent, path_datos, path_familias = NULL){
  sheets <- readxl::excel_sheets(path_datos)
  hoja <- NA_character_
  if (!is.null(path_familias) && file.exists(path_familias)) {
    fam <- tryCatch(readxl::read_excel(path_familias), error = function(e) NULL)
    if (!is.null(fam) && ncol(fam)) {
      cn <- tolower(gsub("[^a-z0-9_]+","_", names(fam)))
      col_parent <- names(fam)[match("parent", cn)]
      col_hoja   <- names(fam)[match("hoja_datos", cn)]
      if (!is.na(col_parent) && !is.na(col_hoja)) {
        fila <- fam[trimws(as.character(fam[[col_parent]])) == parent, , drop = FALSE]
        if (nrow(fila)) hoja <- tolower(as.character(fila[[col_hoja]][1]))
      }
    }
  }
  if (is.na(hoja) || hoja %in% c("", "main")) {
    for (s in sheets) {
      hdr <- tryCatch(names(readxl::read_excel(path_datos, sheet = s, n_max = 0)),
                      error = function(e) character(0))
      if (length(hdr) && any(tolower(hdr) == tolower(parent))) {
        return(list(source = if (s == sheets[1]) "main" else "repeat", sheet = s))
      }
    }
    return(list(source = "main", sheet = sheets[1]))
  } else {
    mi <- match(hoja, tolower(sheets))
    if (!is.na(mi)) {
      s <- sheets[mi]
      return(list(source = if (s == sheets[1]) "main" else "repeat", sheet = s))
    } else {
      for (s in sheets) {
        hdr <- tryCatch(names(readxl::read_excel(path_datos, sheet = s, n_max = 0)),
                        error = function(e) character(0))
        if (length(hdr) && any(tolower(hdr) == tolower(parent))) {
          return(list(source = if (s == sheets[1]) "main" else "repeat", sheet = s))
        }
      }
      return(list(source = "main", sheet = sheets[1]))
    }
  }
}

# -------- XLSForm: choices por parent -----------------------------------------
ppra_get_choices_parent <- function(path_instrumento, parent_var){
  survey  <- readxl::read_excel(path_instrumento, sheet = "survey")
  choices <- readxl::read_excel(path_instrumento, sheet = "choices")

  i <- which(survey$name == parent_var)
  if (!length(i)) {
    clean <- function(x) gsub("[^A-Za-z0-9_]+","", as.character(x))
    i <- which(clean(as.character(survey$name)) == clean(parent_var))
  }
  if (!length(i)) stop("No encontré la pregunta en survey: ", parent_var)

  ty <- as.character(survey$type[i][1] %||% "")
  ln <- trimws(sub("^\\S+\\s+", "", ty))
  if (!nzchar(ln)) stop("No pude determinar list_name desde 'type' para: ", parent_var)

  nmsl <- tolower(names(choices))
  label_es_col <- names(choices)[match(TRUE, nmsl %in% c(
    "label::spanish (es)","label::spanish(es)","label::spanish_es",
    "label_spanish_es","label::spanish","label","label::es"
  ))]
  if (is.na(label_es_col)) label_es_col <- if ("label" %in% names(choices)) "label" else NA_character_

  ch <- choices[ trimws(choices$list_name) == trimws(ln), , drop = FALSE ]
  if (!nrow(ch)) stop("El list_name '", ln, "' no tiene choices en 'choices'.")

  tibble::tibble(
    order = seq_len(nrow(ch)),
    code  = as.character(ch$name),
    label = if (!is.na(label_es_col)) as.character(ch[[label_es_col]]) else as.character(ch$name)
  )
}

# -------- FAMILIAS: detectar text_col por variable (opcional) -----------------
ppra_get_textcol_from_familias <- function(path_familias, parent_var){
  if (is.null(path_familias) || !file.exists(path_familias)) return(NA_character_)
  sh <- readxl::excel_sheets(path_familias)
  out <- NA_character_
  for (s in sh){
    df <- tryCatch(readxl::read_excel(path_familias, sheet = s), error = function(e) NULL)
    if (is.null(df) || !ncol(df)) next
    cn <- tolower(trimws(names(df)))
    i_parent <- match(TRUE, cn %in% c("parent_col","parent","variable","var"))
    i_tipo   <- match(TRUE, cn %in% c("tipo","type"))
    i_text   <- match(TRUE, cn %in% c("text_col","texto_col","text","textcol","text_column"))
    if (is.na(i_parent) || is.na(i_tipo) || is.na(i_text)) next
    pv <- df[[i_parent]]
    if (is.factor(pv)) pv <- as.character(pv)
    hit <- which(trimws(as.character(pv)) == parent_var)
    if (length(hit)) {
      val <- df[[i_text]][hit[1]]
      if (length(val)) { out <- as.character(val); break }
    }
  }
  if (!nz(out)) NA_character_ else out
}

# -------- insertar a la derecha con fallback ----------------------------------
insert_right_of <- function(df, anchor, cols_to_insert){
  cols_to_insert <- cols_to_insert[cols_to_insert %in% names(df)]
  if (!length(cols_to_insert)) return(df)
  nms <- names(df)
  apos <- match(anchor, nms)
  if (is.na(apos)) {
    apos <- match("_index", nms)
    if (is.na(apos)) {
      base <- setdiff(nms, cols_to_insert)
      return(df[, c(base, cols_to_insert), drop = FALSE])
    }
  }
  base <- setdiff(nms, cols_to_insert)
  left <- if (apos<=0L) character(0) else base[seq_len(apos)]
  right<- if (apos<=0L) base else base[(apos+1):length(base)]
  df[, c(left, cols_to_insert, right), drop = FALSE]
}

# ====================== SM =====================================================
ppra_sm_parent_recod <- function(df, parent, path_instrumento, path_plantilla,
                                 path_datos = NULL, path_familias = NULL){
  where <- if (!is.null(path_datos)) locate_var_sheet(parent, path_datos, path_familias) else list(source="main", sheet=NA)
  df_work <- if (identical(where$source, "main")) df else read_sheet_ci(path_datos, where$sheet)
  nms <- names(df_work)

  # catálogo clásico
  ch <- ppra_get_choices_parent(path_instrumento, parent)
  classic      <- as.character(ch$code)
  classic_norm <- .normcode(classic)

  # hoja de plantilla (opcional, dedup nombres)
  tpl <- NULL
  if (file.exists(path_plantilla)) {
    sheet <- ppra_resolve_template_sheet(path_plantilla, parent)
    if (!is.na(sheet)) {
      tpl <- readxl::read_excel(path_plantilla, sheet = sheet)
      tpl <- .ensure_unique_names(tpl)
      names(tpl) <- trimws(names(tpl))
    }
  }

  # tmp para overrides (join armonizado)
  if (!is.null(tpl)) {
    kd <- pick_join_key_pair(df_work, tpl)
    if (is.na(kd)) stop("No hay clave común para SM en ", ifelse(identical(where$source,"main"), "main: ", paste0("repeat '", where$sheet, "': ")), parent)
    tmp <- .safe_left_join_by(df_work[, kd, drop = FALSE], tpl, kd)
  } else {
    kd <- pick_join_key(df_work)
    tmp <- df_work[, kd, drop = FALSE]
  }
  tnames <- names(tmp)

  # matriz clásico desde crudo
  mat <- matrix(NA_integer_, nrow = nrow(df_work), ncol = length(classic))
  colnames(mat) <- classic
  for (code in classic) {
    raw_col <- paste0(parent, "/", code)
    if (raw_col %in% nms) {
      v <- .norm01(df_work[[raw_col]])
      mat[, code] <- v
    }
  }

  # overrides clásico desde plantilla
  if (!is.null(tpl)) {
    for (code in classic) {
      rc1 <- paste0(parent, "/", code, "_RECOD")
      rc2 <- paste0(parent, "/", code, "_recod")
      if (rc1 %in% tnames || rc2 %in% tnames) {
        v <- if (rc1 %in% tnames) tmp[[rc1]] else tmp[[rc2]]
        v <- .norm01(v)
        w1 <- which(!is.na(v) & v==1L); if (length(w1)) mat[w1, code] <- 1L
        w0 <- which(!is.na(v) & v==0L); if (length(w0)) mat[w0, code] <- 0L
      }
    }
  }

  # nuevas hijas en plantilla
  new_tokens <- vector("list", nrow(df_work))
  if (!is.null(tpl)) {
    rx_new <- paste0("^", parent, "/[^/]+_(?i:recod)$")
    new_cols <- tnames[grepl(rx_new, tnames, perl = TRUE)]
    if (length(new_cols)) {
      for (cc in new_cols) {
        base <- sub("^.+/", "", cc)
        base <- sub("_(?i:recod)$", "", base)
        base_norm <- .normcode(base)
        v <- .norm01(tmp[[cc]])
        if (length(classic) && base_norm %in% classic_norm) {
          can_code <- classic[match(base_norm, classic_norm)]
          w1 <- which(!is.na(v) & v==1L); if (length(w1)) mat[w1, can_code] <- 1L
          w0 <- which(!is.na(v) & v==0L); if (length(w0)) mat[w0, can_code] <- 0L
        } else {
          sel <- which(!is.na(v) & v==1L)
          if (length(sel)) for (i in sel) new_tokens[[i]] <- unique(c(new_tokens[[i]], base))
        }
      }
    }
  }

  # tokens finales
  parent_recod <- character(nrow(df_work))
  for (i in seq_len(nrow(df_work))){
    from_classic <- colnames(mat)[which(mat[i, ] == 1L)]
    from_new     <- new_tokens[[i]] %||% character(0)
    other_idx <- match("other", .normcode(from_classic))
    if (is.na(other_idx)) {
      if ("Other" %in% classic) {
        j <- which(colnames(mat)=="Other")
        if (length(j) && !is.na(mat[i,j]) && mat[i,j]==0L) {
          from_new <- from_new[ .normcode(from_new) != "other" ]
        }
      }
    }
    allcodes <- unique(c(from_classic, from_new))
    parent_recod[i] <- if (length(allcodes)) paste(allcodes, collapse = " ") else NA_character_
  }

  out_col <- paste0(parent, "_recod")

  if (identical(where$source, "main")) {
    df[[out_col]] <- parent_recod
    return(list(
      df = df, new_col = out_col,
      repeat_sheet = NULL, repeat_df = NULL,
      repeat_cols_to_color = character(0)
    ))
  } else {
    rep_df2 <- df_work
    rep_df2[[out_col]] <- parent_recod
    rep_df2 <- insert_right_of(rep_df2, parent, out_col)
    return(list(
      df = df, new_col = character(0),
      repeat_sheet = where$sheet, repeat_df = rep_df2,
      repeat_cols_to_color = out_col
    ))
  }
}

# ====================== SO (padre) — Opción B con familias --------------------
ppra_so_parent <- function(df, parent, path_instrumento, path_plantilla,
                           path_familias = NULL, path_datos = NULL){
  where   <- if (!is.null(path_datos)) locate_var_sheet(parent, path_datos, path_familias) else list(source="main", sheet=NA)
  df_work <- if (identical(where$source, "main")) df else read_sheet_ci(path_datos, where$sheet)

  ch <- ppra_get_choices_parent(path_instrumento, parent)
  cat_codes <- as.character(ch$code); cat_labs <- as.character(ch$label)
  lab_by_code <- function(code){ i <- match(code, cat_codes); ifelse(is.na(i), NA_character_, cat_labs[i]) }

  # Plantilla
  tpl <- NULL
  if (file.exists(path_plantilla)) {
    sheet <- ppra_resolve_template_sheet(path_plantilla, parent)
    if (!is.na(sheet)) {
      tpl <- readxl::read_excel(path_plantilla, sheet = sheet)
      tpl <- .ensure_unique_names(tpl)
      names(tpl) <- trimws(names(tpl))
    }
  }

  # tmp (join armonizado)
  if (!is.null(tpl)) {
    kd <- pick_join_key_pair(df_work, tpl)
    if (is.na(kd)) stop("Sin clave común para SO-padre en ", ifelse(identical(where$source,"main"), "main: ", paste0("repeat '", where$sheet, "': ")), parent)
    tmp <- .safe_left_join_by(df_work[, kd, drop = FALSE], tpl, kd)
  } else {
    kd <- pick_join_key(df_work); stopifnot(!is.na(kd))
    tmp <- df_work[, kd, drop = FALSE]
  }
  tnames <- names(tmp)

  .pick <- function(cands) {
    tl <- tolower(tnames)
    for (cand in cands) {
      j <- match(tolower(cand), tl)
      if (!is.na(j)) return(tnames[j])
    }
    NA_character_
  }

  rec_tpl  <- .pick(c(paste0(parent,"_RECOD"), paste0(parent,"_recod")))
  labr_tpl <- .pick(c(paste0(parent,"_LABEL_RECOD"), paste0(parent,"_label_recod")))
  lab_tpl  <- .pick(c(paste0(parent,"_LABEL"),       paste0(parent,"_label")))
  orec_tpl <- .pick(c(paste0(parent,"_OTHER_RECOD"), paste0(parent,"_other_recod")))
  olab_tpl <- .pick(c(paste0(parent,"_OTHER_RECOD_LABEL"), paste0(parent,"_other_recod_label")))

  base_code <- as.character(df_work[[parent]]); base_code[base_code==""] <- NA_character_
  code_final <- base_code
  if (!is.na(rec_tpl)) {
    x <- trimws(as.character(tmp[[rec_tpl]])); x[x==""] <- NA_character_
    i <- which(!is.na(x)); if (length(i)) code_final[i] <- x[i]
  }

  is_other_legacy <- (tolower(code_final) == "other") | (tolower(base_code) == "other")
  is_other_legacy[is.na(is_other_legacy)] <- FALSE
  if (isTRUE(any(is_other_legacy, na.rm = TRUE)) && !is.na(orec_tpl)) {
    y <- trimws(as.character(tmp[[orec_tpl]])); y[y==""] <- NA_character_
    j <- which(is_other_legacy & !is.na(y)); if (length(j)) code_final[j] <- y[j]
  }

  # override por texto de familias
  text_col <- ppra_get_textcol_from_familias(path_familias, parent)
  txt_val <- rep(NA_character_, nrow(df_work))
  if (nz(text_col) && (text_col %in% names(df_work))) {
    txt_raw <- trimws(as.character(df_work[[text_col]])); txt_raw[txt_raw==""] <- NA_character_
    txt_rec_nm <- paste0(text_col, "_recod")
    txt_rec <- if (txt_rec_nm %in% tnames) {
      trimws(as.character(tmp[[txt_rec_nm]]))
    } else if (txt_rec_nm %in% names(df_work)) {
      trimws(as.character(df_work[[txt_rec_nm]]))
    } else NA_character_
    if (length(txt_rec)) txt_rec[txt_rec==""] <- NA_character_
    txt_val <- dplyr::coalesce(txt_rec, txt_raw)
  }
  if (any(nz(txt_val))) {
    idx <- which(nz(txt_val))
    code_final[idx] <- txt_val[idx]
  }

  label_final <- rep(NA_character_, nrow(df_work))
  if (!is.na(labr_tpl)) {
    z <- trimws(as.character(tmp[[labr_tpl]])); z[z==""] <- NA_character_
    label_final[!is.na(z)] <- z[!is.na(z)]
  }
  need <- is.na(label_final) & !is.na(code_final)
  if (any(need)) label_final[need] <- lab_by_code(code_final[need])
  need <- is.na(label_final)
  if (any(need) && !is.na(lab_tpl)) {
    w <- trimws(as.character(tmp[[lab_tpl]])); w[w==""] <- NA_character_
    label_final[need & !is.na(w)] <- w[need & !is.na(w)]
  }
  need <- is.na(label_final)
  if (any(need) && !is.na(olab_tpl)) {
    q <- trimws(as.character(tmp[[olab_tpl]])); q[q==""] <- NA_character_
    label_final[need & !is.na(q)] <- q[need & !is.na(q)]
  }

  out_col <- paste0(parent, "_recod")
  lab_col <- paste0(parent, "_recod_label")

  if (identical(where$source, "main")) {
    df[[out_col]] <- code_final
    if (any(nz(label_final))) df[[lab_col]] <- label_final
    return(list(
      df = df, new_col = out_col,
      repeat_sheet = NULL, repeat_df = NULL,
      repeat_cols_to_color = character(0)
    ))
  } else {
    rep_df2 <- df_work
    rep_df2[[out_col]] <- code_final
    if (any(nz(label_final))) rep_df2[[lab_col]] <- label_final
    rep_df2 <- insert_right_of(rep_df2, parent, c(out_col, lab_col))
    return(list(
      df = df, new_col = character(0),
      repeat_sheet = where$sheet, repeat_df = rep_df2,
      repeat_cols_to_color = c(out_col, lab_col)
    ))
  }
}

# ====================== SO (hijo) ==============================================
resolve_child_recod_col <- function(parent, colnames_tpl, text_col = NULL){
  nm <- trimws(colnames_tpl)
  if (!is.null(text_col) && nzchar(text_col)) {
    target <- paste0(text_col, "_recod")
    j <- which(tolower(nm) == tolower(target))
    if (length(j)) return(nm[j[1]])
  }
  rx <- paste0("^(?:", parent, "([_/]).+_recod)$")
  cand <- nm[grepl(rx, nm, ignore.case = TRUE, perl = TRUE)]
  if (!length(cand)) return(NA_character_)
  cand_base <- sub("(?i)_recod$", "", cand, perl = TRUE)
  have_base <- tolower(cand_base) %in% tolower(nm)
  if (any(have_base)) return(cand[which(have_base)[1]])
  cand[1]
}

ppra_so_child <- function(df, parent, path_plantilla, path_familias = NULL, path_datos = NULL){
  where   <- if (!is.null(path_datos)) locate_var_sheet(parent, path_datos, path_familias) else list(source="main", sheet=NA)
  df_work <- if (identical(where$source, "main")) df else read_sheet_ci(path_datos, where$sheet)

  tpl <- NULL
  if (file.exists(path_plantilla)) {
    sheet <- ppra_resolve_template_sheet(path_plantilla, parent)
    if (!is.na(sheet)) {
      tpl <- readxl::read_excel(path_plantilla, sheet = sheet)
      tpl <- .ensure_unique_names(tpl)
      names(tpl) <- trimws(names(tpl))
    }
  }
  if (is.null(tpl)) {
    message("[SO-hijo] No hay hoja para '", parent, "'. Omito.")
    return(list(df=df, new_col=character(0),
                repeat_sheet=NULL, repeat_df=NULL, repeat_cols_to_color=character(0)))
  }

  kd <- pick_join_key_pair(df_work, tpl)
  if (is.na(kd)) {
    message("[SO-hijo] Sin clave común para '", parent, "'. Omito.")
    return(list(df=df, new_col=character(0),
                repeat_sheet=NULL, repeat_df=NULL, repeat_cols_to_color=character(0)))
  }

  text_col <- ppra_get_textcol_from_familias(path_familias, parent)
  src <- resolve_child_recod_col(parent, names(tpl), text_col = text_col)
  if (is.na(src)) {
    message("[SO-hijo] No hallé hija *_recod para '", parent, "'.")
    return(list(df=df, new_col=character(0),
                repeat_sheet=NULL, repeat_df=NULL, repeat_cols_to_color=character(0)))
  }

  tmp <- .safe_left_join_by(df_work[, c(kd), drop = FALSE], tpl[, c(kd, src), drop = FALSE], kd)
  val <- trimws(as.character(tmp[[src]])); val[val==""] <- NA_character_

  alias  <- sub(paste0("^", parent, "([_/])"), "", sub("(?i)_recod$","", src, perl = TRUE))
  out_col <- paste0(parent, "_", alias, "_recod")

  if (identical(where$source, "main")) {
    if (!(out_col %in% names(df))) df[[out_col]] <- NA_character_
    i <- which(!is.na(val)); if (length(i)) df[[out_col]][i] <- val[i]
    df <- insert_right_of(df, parent, out_col)
    return(list(
      df = df, new_col = out_col,
      repeat_sheet = NULL, repeat_df = NULL,
      repeat_cols_to_color = character(0)
    ))
  } else {
    if (!(out_col %in% names(df_work))) df_work[[out_col]] <- NA_character_
    j <- which(!is.na(val)); if (length(j)) df_work[[out_col]][j] <- val[j]
    df_work <- insert_right_of(df_work, parent, out_col)
    return(list(
      df = df, new_col = character(0),
      repeat_sheet = where$sheet, repeat_df = df_work,
      repeat_cols_to_color = out_col
    ))
  }
}

# ====================== Export preservando hojas ===============================
ppra_export_preserving_sheets <- function(path_datos, out_path,
                                          df_main,
                                          main_cols_color = list(sm = character(0),
                                                                 sop = character(0),
                                                                 soh = character(0)),
                                          repeat_updates = list()) {
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    warning("No se encontró 'openxlsx'. No se exporta Excel.")
    return(invisible(NULL))
  }

  sheets <- readxl::excel_sheets(path_datos)
  main_sheet <- sheets[1]

  wb <- openxlsx::createWorkbook()

  paint_cols <- function(wb, sheet, df, cols, color){
    if (!length(cols)) return()
    idx <- match(cols, names(df))
    idx <- idx[!is.na(idx)]
    if (!length(idx)) return()
    openxlsx::addStyle(wb, sheet, openxlsx::createStyle(fgFill = color),
                       rows = 1:(nrow(df)+1), cols = idx,
                       gridExpand = TRUE, stack = TRUE)
  }

  for (s in sheets) {
    openxlsx::addWorksheet(wb, s)

    if (identical(s, main_sheet)) {
      openxlsx::writeData(wb, s, df_main)
      paint_cols(wb, s, df_main, unique(main_cols_color$sm),  "#DFF5DF")
      paint_cols(wb, s, df_main, unique(c(main_cols_color$sop, main_cols_color$soh)), "#DCEBFF")
    } else if (s %in% names(repeat_updates)) {
      df_rep  <- repeat_updates[[s]]$df
      cols_sm <- repeat_updates[[s]]$sm %||% character(0)
      cols_so <- repeat_updates[[s]]$so %||% character(0)
      openxlsx::writeData(wb, s, df_rep)
      paint_cols(wb, s, df_rep, unique(cols_sm), "#DFF5DF")
      paint_cols(wb, s, df_rep, unique(cols_so), "#DCEBFF")
    } else {
      df0 <- read_sheet_ci(path_datos, s)
      openxlsx::writeData(wb, s, df0)
    }

    openxlsx::freezePane(wb, s, firstActiveRow = 2)
    openxlsx::setColWidths(wb, s, cols = 1:200, widths = "auto")
  }

  openxlsx::saveWorkbook(wb, out_path, overwrite = TRUE)
  invisible(out_path)
}

# ====================== FUNCIÓN PRINCIPAL =====================================
#' @title ppra_adaptar_data
#' @description
#' - **SM**: `<parent>_recod` con overrides y nuevas hijas (verde).
#' - **SO-padre**: `<parent>_recod` con override por `text_col` de familias (azul).
#' - **SO-hijo**: `<parent>_<alias>_recod` (azul).
#' Escribe en la **hoja correcta** (main o repeat) e **inserta al lado del parent**.
#' Exporta un XLSX preservando **todas las hojas** del archivo de datos.
#' @export
ppra_adaptar_data <- function(path_instrumento,
                              path_datos,
                              path_plantilla,
                              sm_vars        = character(0),
                              so_parent_vars = character(0),
                              so_child_vars  = character(0),
                              out_path       = NULL,
                              path_familias  = NULL){

  stopifnot(file.exists(path_instrumento), file.exists(path_datos), file.exists(path_plantilla))

  df <- leer_datos_generico(path_datos)
  df <- .ensure_unique_names(df)

  sm_cols_to_color  <- character(0)
  sop_cols_to_color <- character(0)
  soh_cols_to_color <- character(0)

  # registro de hojas repeat modificadas
  rep_updates <- list()   # nombre_hoja -> list(df=..., sm=c(...), so=c(...))

  add_rep_update <- function(sheet, df_rep, cols, kind = c("sm","so")){
    kind <- match.arg(kind)
    if (!sheet %in% names(rep_updates)) {
      rep_updates[[sheet]] <<- list(df = df_rep, sm = character(0), so = character(0))
    } else {
      rep_updates[[sheet]]$df <<- df_rep
    }
    if (kind == "sm") {
      rep_updates[[sheet]]$sm <<- unique(c(rep_updates[[sheet]]$sm, cols))
    } else {
      rep_updates[[sheet]]$so <<- unique(c(rep_updates[[sheet]]$so, cols))
    }
  }

  # ---- SM
  if (length(sm_vars)) {
    for (pv in sm_vars) {
      res <- ppra_sm_parent_recod(df, pv, path_instrumento, path_plantilla,
                                  path_datos = path_datos, path_familias = path_familias)
      df  <- res$df
      if (length(res$new_col)) {
        df  <- insert_right_of(df, pv, res$new_col)
        sm_cols_to_color <- c(sm_cols_to_color, res$new_col)
      }
      if (!is.null(res$repeat_sheet)) {
        add_rep_update(res$repeat_sheet, res$repeat_df, res$repeat_cols_to_color, kind = "sm")
      }
    }
  }

  # ---- SO padre
  if (length(so_parent_vars)) {
    for (pv in so_parent_vars) {
      res <- ppra_so_parent(df, pv, path_instrumento, path_plantilla,
                            path_familias = path_familias, path_datos = path_datos)
      df  <- res$df
      if (length(res$new_col)) {
        df  <- insert_right_of(df, pv, res$new_col)
        sop_cols_to_color <- c(sop_cols_to_color, res$new_col)
      }
      if (!is.null(res$repeat_sheet)) {
        add_rep_update(res$repeat_sheet, res$repeat_df, res$repeat_cols_to_color, kind = "so")
      }
    }
  }

  # ---- SO hijo
  if (length(so_child_vars)) {
    for (pv in so_child_vars) {
      res <- ppra_so_child(df, pv, path_plantilla,
                           path_familias = path_familias, path_datos = path_datos)
      df  <- res$df
      if (length(res$new_col)) {
        df  <- insert_right_of(df, pv, res$new_col)
        soh_cols_to_color <- c(soh_cols_to_color, res$new_col)
      }
      if (!is.null(res$repeat_sheet)) {
        add_rep_update(res$repeat_sheet, res$repeat_df, res$repeat_cols_to_color, kind = "so")
      }
    }
  }

  # Export preservando TODAS las hojas
  if (!is.null(out_path)) {
    ppra_export_preserving_sheets(
      path_datos  = path_datos,
      out_path    = out_path,
      df_main     = df,
      main_cols_color = list(
        sm  = unique(sm_cols_to_color),
        sop = unique(sop_cols_to_color),
        soh = unique(soh_cols_to_color)
      ),
      repeat_updates = rep_updates
    )
  }

  df
}
