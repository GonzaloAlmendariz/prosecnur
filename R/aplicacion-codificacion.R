suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(stringr)
  library(tidyr)
})

# =============================
# Helpers básicos y robustos
# =============================
`%||%` <- function(a,b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a
nz      <- function(x) !is.na(x) & nzchar(trimws(as.character(x)))

last_word <- function(s) {
  v <- trimws(as.character(s))
  v[!nz(v)] <- NA_character_
  vapply(v, function(z){
    if (is.na(z)) return(NA_character_)
    toks <- strsplit(z, "\\s+", perl = TRUE)[[1]]
    toks <- toks[nzchar(toks)]
    if (!length(toks)) return(NA_character_)
    toks[length(toks)]
  }, FUN.VALUE = character(1))
}

norm_codes_1space <- function(x) {
  if (is.null(x)) return(NA_character_)
  s <- as.character(x)
  s <- gsub("[;,]", " ", s, perl = TRUE)
  s <- gsub("\\s+", " ", s, perl = TRUE)
  s <- trimws(s)
  s[s==""] <- NA_character_
  s
}

norm01 <- function(x) {
  if (is.numeric(x)) return(ifelse(is.na(x), NA_integer_, ifelse(x != 0, 1L, 0L)))
  s <- tolower(trimws(as.character(x)))
  s <- iconv(s, from = "", to = "ASCII//TRANSLIT")
  out <- rep(NA_integer_, length(s))
  out[s %in% c("1","t","true","si","s","yes","y","verdadero")] <- 1L
  out[s %in% c("0","f","false","no","n","falso")]               <- 0L
  out
}

pick_join_key <- function(df){
  cands <- c("_uuid","uuid","meta_instance_id","instanceid","_id",
             "_index","Codigo pulso","Código pulso","Pulso_code","pulso_code")
  hit <- cands[cands %in% names(df)]
  if (length(hit)) hit[1] else NA_character_
}

# =========================
# Resolver nombre de hoja en plantilla (31 chars, similitud, etc.)
# =========================
.ppra_clean <- function(x) gsub("[^A-Za-z0-9]+", "", tolower(trimws(as.character(x))))

# Resolver hoja por nombre (exacto / case-insensitive / truncado / limpio / similitud)
ppra_resolve_template_sheet <- function(path_plantilla, parent_var, verbose = TRUE, list_sheets_on_fail = FALSE){
  stopifnot(file.exists(path_plantilla))
  sh <- readxl::excel_sheets(path_plantilla)
  if (!length(sh)) stop("La plantilla no tiene hojas.")

  # 1) exacto
  hit <- which(sh == parent_var); if (length(hit)) return(sh[hit[1]])
  # 2) insensitive
  hit <- which(tolower(sh) == tolower(parent_var)); if (length(hit)) return(sh[hit[1]])
  # 3) truncamiento 31
  p31 <- substr(parent_var, 1, 31)
  hit <- which(sh == p31 | tolower(sh) == tolower(p31)); if (length(hit)) {
    if (verbose) message("Usando hoja por truncamiento a 31: '", sh[hit[1]], "'")
    return(sh[hit[1]])
  }
  # 4) limpieza
  cl_parent <- .ppra_clean(parent_var); cl_sheets <- .ppra_clean(sh)
  hit <- which(cl_sheets == cl_parent); if (length(hit)) return(sh[hit[1]])
  # 5) similitud
  d <- adist(cl_parent, cl_sheets); j <- which.min(d)
  if (length(j) && is.finite(d[j]) && d[j] <= 5) return(sh[j])

  if (verbose && list_sheets_on_fail) {
    message("No pude mapear '", parent_var, "' a una hoja en '", basename(path_plantilla), "'.")
  }
  return(NA_character_)
}

# Dado un select_one en survey, ubica su text "other" hija inmediata
ppra_guess_text_other_from_xlsform <- function(path_instrumento, parent){
  stopifnot(file.exists(path_instrumento))
  survey <- readxl::read_excel(path_instrumento, sheet = "survey")
  survey$type  <- as.character(survey$type)
  survey$name  <- as.character(survey$name)

  i <- which(survey$name == parent)
  if (!length(i)) return(NA_character_)

  # Busca hacia abajo la primera fila con type que empiece en "text"
  sub <- survey[seq.int(i+1, nrow(survey)), , drop = FALSE]
  if (!nrow(sub)) return(NA_character_)
  j <- which(grepl("^text\\b", tolower(sub$type)))
  if (!length(j)) return(NA_character_)
  text_name <- trimws(as.character(sub$name[j[1]]))
  if (!nzchar(text_name)) return(NA_character_)
  paste0("TEXT_", text_name)  # nombre de la hoja en plantilla
}

# Alias privado para compatibilidad con código que aún lo llama
.ppra_resolve_tpl_sheet <- function(path_plantilla, parent_var, verbose = TRUE){
  ppra_resolve_template_sheet(path_plantilla, parent_var, verbose = verbose)
}

# ========= Fallback: construir un select-one usando TEXT_<...>_other =========

ppra_build_select_one_from_text_other <- function(df, parent, choices_df, path_plantilla){
  nms <- names(df)

  # 1) adivinar la variable *_other y su hoja TEXT_<...>
  base_parent <- sub("^mand_", "", parent)
  other_cands <- c(
    paste0(base_parent, "_other"),
    paste0(parent, "_other"),
    sub("(payment_mechanism)$", "other", parent, ignore.case = TRUE)  # ej. Possession_payment_mechanism -> Possession_other
  )
  detail_var <- NA_character_
  sheet <- NA_character_
  for (c in other_cands) {
    ss <- ppra_resolve_template_sheet(path_plantilla, paste0("TEXT_", c), verbose = FALSE)
    if (!is.na(ss)) { detail_var <- c; sheet <- ss; break }
  }
  if (is.na(sheet)) stop("Fallback TEXT_* no disponible para ", parent)

  # 2) merge plantilla TEXT_<detail_var>
  tpl <- readxl::read_excel(path_plantilla, sheet = sheet)
  names(tpl) <- trimws(names(tpl))
  kd <- pick_join_key(df); kt <- pick_join_key(tpl)
  if (is.na(kd) || is.na(kt)) stop("Sin clave común para unir plantilla TEXT_* con datos en ", parent)
  names(tpl)[names(tpl) == kt] <- kd
  oth <- setdiff(names(tpl), kd)
  names(tpl)[names(tpl) %in% oth] <- paste0(oth, "__tpl")
  df <- dplyr::left_join(df, tpl, by = kd)
  nms <- names(df)

  # 3) columnas de recod y label provenientes del TEXT
  get_col <- function(...) { cands <- c(...); hit <- cands[cands %in% nms]; if (length(hit)) hit[1] else NA_character_ }
  rec_from_text <- get_col(paste0(detail_var, "_RECOD__tpl"), paste0(detail_var, "_recod__tpl"))
  lab_from_text <- get_col(paste0(detail_var, "_LABEL__tpl"), paste0(detail_var, "_label__tpl"),
                           paste0(detail_var, "_RECOD_LABEL__tpl"), paste0(detail_var, "_recod_label__tpl"))

  # 4) catálogo para labels
  cat_codes <- as.character(choices_df$code)
  cat_labs  <- as.character(choices_df$label)
  lab_by_code <- function(code){ i <- match(code, cat_codes); ifelse(is.na(i), NA_character_, cat_labs[i]) }

  # 5) construir <parent>_recod y <parent>_recod_label
  parent_recod_col <- paste0(parent, "_recod")
  parent_lab_col   <- paste0(parent, "_recod_label")

  base_code  <- trimws(as.character(df[[parent]])); base_code[base_code==""] <- NA_character_
  code_final <- base_code

  # Regla: si hay recod en TEXT, tiene la última palabra
  if (!is.na(rec_from_text)) {
    tok <- trimws(as.character(df[[rec_from_text]])); tok[tok==""] <- NA_character_
    i <- which(!is.na(tok))
    if (length(i)) code_final[i] <- tok[i]
  }

  # Label: prioridad label explícito en TEXT, si no, catálogo
  label_final <- rep(NA_character_, nrow(df))
  if (!is.na(lab_from_text)) {
    y <- trimws(as.character(df[[lab_from_text]])); y[y==""] <- NA_character_
    label_final[!is.na(y)] <- y[!is.na(y)]
  }
  need <- is.na(label_final) & !is.na(code_final)
  if (any(need)) label_final[need] <- lab_by_code(code_final[need])

  df[[parent_recod_col]] <- code_final
  df[[parent_lab_col]]   <- label_final

  # quitar __tpl
  keep <- !grepl("__tpl($|\\.[xy]$)", names(df), perl = TRUE)
  df <- df[, keep, drop = FALSE]

  list(df = df, new_cols = c(parent_recod_col, parent_lab_col))
}




# =============================
# Lecturas
# =============================
leer_datos_generico <- function(path_data, sheet = NULL){
  ext <- tolower(tools::file_ext(path_data))
  if (ext %in% c("csv","txt")) {
    suppressWarnings(readr::read_csv(path_data, show_col_types = FALSE))
  } else {
    readxl::read_excel(path_data, sheet = sheet %||% 1)
  }
}

ppra_get_choices_parent <- function(path_instrumento, parent_var){
  stopifnot(file.exists(path_instrumento))
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
    "label_spanish_es","label::spanish","label::es"
  ))]
  if (is.na(label_es_col)) label_es_col <- if ("label" %in% names(choices)) "label" else NA_character_

  ch <- choices[ trimws(choices$list_name) == trimws(ln), , drop = FALSE ]
  if (!nrow(ch)) stop("El list_name '", ln, "' no tiene choices en la hoja 'choices'.")

  tibble::tibble(
    order = seq_len(nrow(ch)),
    code  = as.character(ch$name),
    label = if (!is.na(label_es_col)) as.character(ch[[label_es_col]]) else as.character(ch$name)
  )
}

ppra_read_template_sheet <- function(path_plantilla, parent_var){
  stopifnot(file.exists(path_plantilla))
  sheet_name <- ppra_resolve_template_sheet(path_plantilla, parent_var, verbose = TRUE)
  if (is.na(sheet_name)) stop("Sheet para '", parent_var, "' no encontrado en la plantilla.")
  df <- readxl::read_excel(path_plantilla, sheet = sheet_name)
  names(df) <- trimws(names(df))

  nms <- names(df)
  list(
    df            = df,
    parent_exact  = parent_var,
    hijas_crudas  = nms[grepl(paste0("^", parent_var, "/[^/]+$"), nms, perl = TRUE)],
    hijas_recod   = nms[grepl(paste0("^", parent_var, "/[^/]+_(?i:recod)$"), nms, perl = TRUE)],
    other_col     = nms[grepl(paste0("^", parent_var, "/Other$"), nms, perl = TRUE)],
    other_recod   = nms[grepl(paste0("^", parent_var, "/(?i:Other_recod)$"), nms, perl = TRUE)],
    nuevas_recod  = setdiff(
      nms[grepl(paste0("^", parent_var, "/[^/]+_(?i:recod)$"), nms, perl = TRUE)],
      character(0)
    )
  )
}

.ppra_norm_recod_vec <- function(v){
  if (is.numeric(v)) {
    out <- ifelse(is.na(v), NA_integer_, ifelse(v != 0, 1L, 0L))
  } else {
    out <- ifelse(v %in% c(1,"1",TRUE,"true","TRUE"), 1L,
                  ifelse(v %in% c(0,"0",FALSE,"false","FALSE"), 0L, NA_integer_))
  }
  out
}

ppra_enforce_all_recods <- function(df, parents = NULL){
  nms <- names(df)
  rec_mask <- grepl("^[^/]+/[^/]+_recod$", nms, perl = TRUE)
  if (!any(rec_mask)) return(df)

  rec_cols <- nms[rec_mask]
  detected_parents <- unique(sub("^([^/]+)/.*$", "\\1", rec_cols))
  parents <- if (is.null(parents)) detected_parents else intersect(parents, detected_parents)

  for (pv in parents){
    cols_pv <- nms[rec_mask & startsWith(nms, paste0(pv, "/"))]
    if (!length(cols_pv)) next

    blk <- df[cols_pv]
    blk[] <- lapply(blk, .ppra_norm_recod_vec)

    any1 <- rowSums(blk == 1L, na.rm = TRUE) > 0
    if (any(any1)) {
      for (j in seq_along(blk)) {
        v <- blk[[j]]
        v[any1 & is.na(v)] <- 0L
        blk[[j]] <- v
      }
    }
    df[cols_pv] <- blk
  }
  df
}

ppra_parent_rebuild_fallback <- function(df, parents = NULL){
  nms <- names(df)
  rec_mask <- grepl("^[^/]+/[^/]+_recod$", nms, perl = TRUE)
  if (!any(rec_mask)) return(df)

  rec_cols <- nms[rec_mask]
  detected_parents <- unique(sub("^([^/]+)/.*$", "\\1", rec_cols))
  parents <- if (is.null(parents)) detected_parents else intersect(parents, detected_parents)

  for (pv in parents){
    parent_col <- paste0(pv, "_recod")
    child_rec  <- nms[rec_mask & startsWith(nms, paste0(pv, "/"))]
    if (!length(child_rec)) next

    raw_mask  <- grepl(paste0("^", pv, "/[^/]+$"), nms, perl = TRUE)
    raw_child <- nms[raw_mask]
    raw_codes <- sub("^.+/", "", raw_child, perl = TRUE)

    rec_codes <- sub("^.+/", "", child_rec, perl = TRUE)
    rec_codes <- sub("_recod$", "", rec_codes, perl = TRUE)

    ord     <- match(rec_codes, raw_codes)
    ord_idx <- order(ifelse(is.na(ord), 1e9L, ord), rec_codes)
    child_rec <- child_rec[ord_idx]
    rec_codes <- rec_codes[ord_idx]

    blk <- df[child_rec]
    blk[] <- lapply(blk, .ppra_norm_recod_vec)

    tokens <- apply(blk, 1, function(r){
      sel <- which(r == 1L)
      if (!length(sel)) return(NA_character_)
      paste(rec_codes[sel], collapse = " ")
    })

    if (!parent_col %in% nms) df[[parent_col]] <- NA_character_

    need <- is.na(df[[parent_col]]) | trimws(df[[parent_col]]) == ""
    has_tokens <- !is.na(tokens) & nzchar(tokens)
    fill_idx <- which(need & has_tokens)
    if (length(fill_idx)) df[[parent_col]][fill_idx] <- tokens[fill_idx]
  }
  df
}

# --- 2b) AUGMENT: añade al <Parent>_recod los códigos 1 de hijas (aunque ya tenga texto)
ppra_parent_augment_recod <- function(df, parents = NULL){
  nms <- names(df)

  # detectar parents desde hijas recod si no se pasan
  rec_mask <- grepl("^[^/]+/[^/]+_recod$", nms, perl = TRUE)
  if (!any(rec_mask)) return(df)
  rec_cols <- nms[rec_mask]
  detected_parents <- unique(sub("^([^/]+)/.*$", "\\1", rec_cols))
  parents <- if (is.null(parents)) detected_parents else intersect(parents, detected_parents)

  .split_tokens <- function(x){
    x <- trimws(as.character(x))
    x[is.na(x) | x==""] <- NA_character_
    strsplit(ifelse(is.na(x), "", x), "\\s+", perl = TRUE)
  }

  for (pv in parents){
    parent_col <- paste0(pv, "_recod")
    child_rec  <- nms[rec_mask & startsWith(nms, paste0(pv, "/"))]
    if (!length(child_rec)) next

    # orden de referencia del bloque crudo
    raw_mask  <- grepl(paste0("^", pv, "/[^/]+$"), nms, perl = TRUE)
    raw_child <- nms[raw_mask]
    raw_codes <- sub("^.+/", "", raw_child)  # ej: c("Weekly_market","Shelter",...)

    # códigos de hijas recod (en el orden actual de columnas)
    rec_codes <- sub("^.+/", "", child_rec)
    rec_codes <- sub("_recod$", "", rec_codes)

    # matriz 0/1 de hijas recod
    blk <- df[child_rec]
    blk[] <- lapply(blk, function(v) as.integer(v %in% c(1,"1",TRUE,"true","TRUE")))

    # tokens ya en el padre
    if (!parent_col %in% nms) df[[parent_col]] <- NA_character_
    cur_tokens <- .split_tokens(df[[parent_col]])

    # construir tokens por fila = union( tokens_padre , hijas==1 )
    new_parent <- character(nrow(df))
    for (i in seq_len(nrow(df))){
      from_parent <- cur_tokens[[i]]
      from_parent <- from_parent[!is.na(from_parent) & nzchar(from_parent)]

      sel <- which(blk[i, ] == 1L)
      from_children <- if (length(sel)) rec_codes[sel] else character(0)

      # unión de códigos, sin duplicados
      uni <- unique(c(from_parent, from_children))

      # ordenar por orden crudo; desconocidos al final por nombre
      if (length(uni)){
        ord <- match(uni, raw_codes)
        oix <- order(ifelse(is.na(ord), 1e9L, ord), uni)
        uni <- uni[oix]
        new_parent[i] <- paste(uni, collapse = " ")
      } else {
        new_parent[i] <- NA_character_
      }
    }

    # solo escribe si cambia (evita tocar filas innecesarias)
    change <- (is.na(df[[parent_col]]) & !is.na(new_parent)) |
      (!is.na(df[[parent_col]]) & is.na(new_parent)) |
      (!is.na(df[[parent_col]]) & !is.na(new_parent) & trimws(df[[parent_col]]) != new_parent)

    df[[parent_col]][change] <- new_parent[change]
  }

  df
}

# =============================
# Caso especial de hojas TEXT_
# =============================

# === 1A) Detecta el padre select_one de un campo other (text) mirando la hoja survey ===
ppra_survey_parent_of_other <- function(path_instrumento, other_name){
  stopifnot(file.exists(path_instrumento))
  survey <- readxl::read_excel(path_instrumento, sheet = "survey")
  if (!all(c("type","name") %in% names(survey))) return(NA_character_)

  survey$type <- as.character(survey$type)
  survey$name <- as.character(survey$name)

  i <- which(trimws(survey$name) == trimws(other_name) &
               grepl("^\\s*text\\b", survey$type, ignore.case = TRUE))
  if (!length(i)) return(NA_character_)

  # buscar hacia arriba la fila select_one más cercana
  j <- rev(seq_len(i[1]-1))
  sel <- j[grepl("^\\s*select_one\\b", survey$type[j], ignore.case = TRUE)]
  if (!length(sel)) return(NA_character_)
  trimws(survey$name[sel[1]])
}

# === 1B) Localiza hoja TEXT_<other> y sus columnas <other> / <other>_recod para un parent dado ===
.ppra_find_text_sheet_for_parent <- function(path_plantilla, path_instrumento, parent){
  stopifnot(file.exists(path_plantilla))
  shs <- readxl::excel_sheets(path_plantilla)
  text_sheets <- shs[grepl("^TEXT_", shs, ignore.case = TRUE)]
  if (!length(text_sheets)) return(NULL)

  base_parent <- sub("^mand_", "", parent)

  for (s in text_sheets) {
    base_other <- sub("^TEXT_","", s, ignore.case = TRUE)     # ej: "Possession_other"
    # mapear a su padre según survey
    parent_from_survey <- ppra_survey_parent_of_other(path_instrumento, base_other)
    if (is.na(parent_from_survey)) next

    if (!(parent_from_survey %in% c(parent, base_parent))) next

    hdr <- tryCatch(names(readxl::read_excel(path_plantilla, sheet = s, n_max = 0)), error = function(e) NULL)
    if (is.null(hdr)) next

    cand_other <- c(base_other, paste0(base_other, "__tpl"))
    cand_recod <- c(paste0(base_other,"_recod"), paste0(base_other,"_recod__tpl"))

    other_col <- cand_other[cand_other %in% hdr]
    rec_col   <- cand_recod[cand_recod %in% hdr]
    if (length(other_col) && length(rec_col)) {
      return(list(sheet = s, other_col = other_col[1], other_recod_col = rec_col[1]))
    }
  }
  NULL
}


# === 2) Construcción select-one con fallback automático a hoja TEXT_<other> ===
ppra_build_select_one_auto <- function(df, parent, choices_df, path_plantilla, path_instrumento){
  nms <- names(df)
  base_parent <- sub("^mand_","", parent)

  # ——— A) Intento estándar: hoja del parent ———
  sheet_name <- ppra_resolve_template_sheet(path_plantilla, parent, verbose = FALSE)
  tpl_std <- NULL
  if (!is.na(sheet_name)) {
    tpl_std <- readxl::read_excel(path_plantilla, sheet = sheet_name)
    names(tpl_std) <- trimws(names(tpl_std))
  }

  .pick <- function(cands, nms) { hits <- cands[cands %in% nms]; if (length(hits)) hits[1] else NA_character_ }

  use_text_fallback <- TRUE
  recod_tpl_col <- label_recod_tpl_col <- label_tpl_col <- other_rec_tpl_col <- other_labrec_tpl_col <- NA_character_

  if (!is.null(tpl_std)) {
    recod_tpl_col       <- .pick(c(paste0(parent,"_RECOD__tpl"),        paste0(parent,"_recod__tpl"),
                                   paste0(base_parent,"_RECOD__tpl"),   paste0(base_parent,"_recod__tpl")),
                                 names(tpl_std))
    label_recod_tpl_col <- .pick(c(paste0(parent,"_LABEL_RECOD__tpl"),      paste0(parent,"_label_recod__tpl"),
                                   paste0(base_parent,"_LABEL_RECOD__tpl"), paste0(base_parent,"_label_recod__tpl")),
                                 names(tpl_std))
    label_tpl_col       <- .pick(c(paste0(parent,"_LABEL__tpl"),        paste0(parent,"_label__tpl"),
                                   paste0(base_parent,"_LABEL__tpl"),   paste0(base_parent,"_label__tpl")),
                                 names(tpl_std))
    other_rec_tpl_col   <- .pick(c(paste0(parent,"_OTHER_RECOD__tpl"),      paste0(parent,"_other_recod__tpl"),
                                   paste0(base_parent,"_OTHER_RECOD__tpl"), paste0(base_parent,"_other_recod__tpl"),
                                   paste0(base_parent,"_other__tpl")),
                                 names(tpl_std))
    other_labrec_tpl_col<- .pick(c(paste0(parent,"_OTHER_RECOD_LABEL__tpl"),      paste0(parent,"_other_recod_label__tpl"),
                                   paste0(base_parent,"_OTHER_RECOD_LABEL__tpl"), paste0(base_parent,"_other_recod_label__tpl"),
                                   paste0(base_parent,"_other_label__tpl")),
                                 names(tpl_std))

    if (any(!is.na(c(recod_tpl_col, label_recod_tpl_col, label_tpl_col, other_rec_tpl_col, other_labrec_tpl_col)))) {
      use_text_fallback <- FALSE
    }
  }

  kd <- pick_join_key(df)

  if (!use_text_fallback) {
    # ===== Flujo estándar =====
    kt <- pick_join_key(tpl_std); stopifnot(!is.na(kt))
    tpl2 <- tpl_std
    names(tpl2)[names(tpl2)==kt] <- kd
    keep_cols <- c(kd, recod_tpl_col, label_recod_tpl_col, label_tpl_col, other_rec_tpl_col, other_labrec_tpl_col)
    tpl2 <- tpl2[, unique(keep_cols[!is.na(keep_cols)]), drop = FALSE]
    df <- dplyr::left_join(df, tpl2, by = kd)

    cat_codes <- as.character(choices_df$code)
    cat_labs  <- as.character(choices_df$label)
    lab_by_code <- function(code){ i <- match(code, cat_codes); ifelse(is.na(i), NA_character_, cat_labs[i]) }

    base_code <- as.character(df[[parent]]); base_code[base_code==""] <- NA_character_

    code_final <- base_code
    if (!is.na(recod_tpl_col)) {
      x <- trimws(as.character(df[[recod_tpl_col]])); x[x==""] <- NA_character_
      repl <- which(!is.na(x)); if (length(repl)) code_final[repl] <- x[repl]
    }

    is_other <- (!is.na(code_final) & tolower(code_final) == "other") |
      (!is.na(base_code)  & tolower(base_code)  == "other")
    if (any(is_other) && !is.na(other_rec_tpl_col)) {
      tok <- trimws(as.character(df[[other_rec_tpl_col]])); tok[tok==""] <- NA_character_
      fill <- which(is_other & !is.na(tok))
      if (length(fill)) code_final[fill] <- tok[fill]
    }

    label_final <- rep(NA_character_, nrow(df))
    if (!is.na(label_recod_tpl_col)) {
      y <- trimws(as.character(df[[label_recod_tpl_col]])); y[y==""] <- NA_character_
      label_final[!is.na(y)] <- y[!is.na(y)]
    }
    need <- is.na(label_final) & !is.na(code_final)
    if (any(need)) label_final[need] <- lab_by_code(code_final[need])
    need <- is.na(label_final)
    if (any(need) && !is.na(label_tpl_col)) {
      z <- trimws(as.character(df[[label_tpl_col]])); z[z==""] <- NA_character_
      label_final[need & !is.na(z)] <- z[need & !is.na(z)]
    }
    need <- is.na(label_final)
    if (any(need) && !is.na(other_labrec_tpl_col)) {
      w <- trimws(as.character(df[[other_labrec_tpl_col]])); w[w==""] <- NA_character_
      label_final[need & !is.na(w)] <- w[need & !is.na(w)]
    }

    rec_col <- paste0(parent, "_recod")
    lab_col <- paste0(parent, "_recod_label")
    df[[rec_col]] <- code_final
    df[[lab_col]] <- label_final
    return(list(df = df, new_cols = c(rec_col, lab_col)))
  }

  # ——— B) Fallback: hoja TEXT_<other> asociada al parent ———
  pr <- .ppra_find_text_sheet_for_parent(path_plantilla, path_instrumento, parent)
  if (is.null(pr)) {
    warning("SO auto: no hallé ni hoja estándar ni TEXT_<other> para '", parent, "'.")
    return(list(df = df, new_cols = character(0)))
  }

  tpl <- readxl::read_excel(path_plantilla, sheet = pr$sheet)
  names(tpl) <- trimws(names(tpl))
  kt <- pick_join_key(tpl); stopifnot(!is.na(kt))
  tpl2 <- tpl; names(tpl2)[names(tpl2)==kt] <- kd
  tpl2 <- tpl2[, unique(c(kd, pr$other_col, pr$other_recod_col)), drop = FALSE]
  df <- dplyr::left_join(df, tpl2, by = kd)

  # catálogo
  cat_codes <- as.character(choices_df$code)
  cat_labs  <- as.character(choices_df$label)
  lab_by_code <- function(code){ i <- match(code, cat_codes); ifelse(is.na(i), NA_character_, cat_labs[i]) }

  base_code   <- as.character(df[[parent]]); base_code[base_code==""] <- NA_character_
  other_text  <- trimws(as.character(df[[pr$other_col]]));        other_text[other_text==""] <- NA_character_
  other_recod <- trimws(as.character(df[[pr$other_recod_col]]));   other_recod[other_recod==""] <- NA_character_

  code_final <- base_code
  is_other_now <- !is.na(code_final) & tolower(code_final) == "other"
  fill_idx <- which(is_other_now & !is.na(other_recod))
  if (length(fill_idx)) code_final[fill_idx] <- other_recod[fill_idx]

  label_final <- lab_by_code(code_final)
  need <- is.na(label_final) & !is.na(other_text)
  if (any(need)) label_final[need] <- other_text[need]

  rec_col <- paste0(parent, "_recod")
  lab_col <- paste0(parent, "_recod_label")
  df[[rec_col]] <- code_final
  df[[lab_col]] <- label_final

  list(df = df, new_cols = c(rec_col, lab_col))
}




# =============================
# Construcción del bloque recodificado (padre + hijas)
# =============================
ppra_build_parent_block <- function(df_merged, parent_var, choices_df) {
  df             <- df_merged
  nms            <- names(df)
  out_parent_col <- paste0(parent_var, "_recod")
  child_cols     <- character(0)
  new_child_cols <- character(0)

  classic_codes <- as.character(choices_df$code %||% character(0))

  final_mat <- if (length(classic_codes)) {
    m <- matrix(NA_integer_, nrow = nrow(df), ncol = length(classic_codes))
    colnames(m) <- classic_codes
    m
  } else {
    matrix(, nrow = nrow(df), ncol = 0L)
  }

  if (length(classic_codes)) {
    for (code in classic_codes) {
      raw_col <- paste0(parent_var, "/", code)
      if (raw_col %in% nms) final_mat[, code] <- norm01(df[[raw_col]])
    }
  }

  if (length(classic_codes)) {
    for (code in classic_codes) {
      rc1 <- paste0(parent_var, "/", code, "_RECOD__tpl")
      rc2 <- paste0(parent_var, "/", code, "_recod__tpl")
      if (rc1 %in% nms || rc2 %in% nms) {
        vrec <- if (rc1 %in% nms) df[[rc1]] else df[[rc2]]
        vnum <- norm01(vrec)
        if (any(!is.na(vnum) & vnum == 1L)) final_mat[which(vnum==1L), code] <- 1L
        if (any(!is.na(vnum) & vnum == 0L)) final_mat[which(vnum==0L), code] <- 0L
      }
    }
  }

  if (length(classic_codes) && "Other" %in% classic_codes) {
    other_code <- "Other"
    or1 <- paste0(parent_var, "/Other_RECOD__tpl")
    or2 <- paste0(parent_var, "/Other_recod__tpl")
    if (or1 %in% nms || or2 %in% nms) {
      vraw <- if (or1 %in% nms) df[[or1]] else df[[or2]]
      vnum <- norm01(vraw)
      off <- which(!is.na(vnum) & vnum == 0L)
      if (length(off)) final_mat[off, other_code] <- 0L

      is_txt <- nz(as.character(vraw)) & is.na(vnum)
      if (any(is_txt)) {
        to_code <- last_word(vraw[is_txt])
        ok <- which(is_txt & to_code %in% classic_codes)
        if (length(ok)) {
          rows <- which(is_txt)[ok]
          dst  <- to_code[ok]
          for (k in seq_along(rows)) {
            r <- rows[k]; cd <- dst[k]
            final_mat[r, other_code] <- 0L
            final_mat[r, cd] <- 1L
          }
        }
      }
    }
  }

  rx_new_tpl   <- paste0("^", parent_var, "/[^/]+_(?i:recod)__tpl$")
  new_tpl_cols <- nms[grepl(rx_new_tpl, nms, perl = TRUE)]
  new_defs     <- list()

  if (length(new_tpl_cols)) {
    for (cc in new_tpl_cols) {
      leaf <- sub("__tpl$", "", cc)
      leaf <- sub("^.+/", "", leaf)
      base <- sub("_(?i:recod)$", "", leaf, perl = TRUE)
      if (length(classic_codes) && base %in% classic_codes) next

      rec_num <- norm01(df[[cc]])
      out_col <- paste0(parent_var, "/", base, "_recod")

      df[[out_col]] <- ifelse(is.na(rec_num), NA_character_, as.character(rec_num))
      new_defs[[length(new_defs)+1]] <- list(code = base, out_col = out_col, vec = rec_num)
    }
  }
  if (length(new_defs)) {
    new_child_cols <- vapply(new_defs, function(z) z$out_col, FUN.VALUE = character(1))
  }

  final_df <- as.data.frame(final_mat, check.names = FALSE)

  tokens_classic <- if (ncol(final_df)) {
    apply(final_df, 1, function(r){
      sel <- names(r)[which(r == 1L)]
      if (!length(sel)) character(0) else {
        ord <- match(sel, classic_codes)
        sel[order(ord)]
      }
    })
  } else {
    replicate(nrow(df), character(0), simplify = FALSE)
  }

  tokens_new <- replicate(nrow(df), character(0), simplify = FALSE)
  if (length(new_defs)) {
    new_order <- vapply(new_defs, function(z) z$code, FUN.VALUE = character(1))
    for (i in seq_len(nrow(df))) {
      add <- character(0)
      for (nd in new_defs) {
        if (!is.null(nd$vec) && length(nd$vec) == nrow(df) && !is.na(nd$vec[i]) && nd$vec[i] == 1L) {
          add <- c(add, nd$code)
        }
      }
      if (length(add)) add <- add[match(add, new_order)]
      tokens_new[[i]] <- add
    }
  }

  parent_tokens <- mapply(function(a,b) unique(c(a,b)), tokens_classic, tokens_new, SIMPLIFY = FALSE)
  parent_concat <- vapply(
    parent_tokens,
    function(v) {
      v <- v[!is.na(v) & nzchar(v)]
      if (length(v)) paste(v, collapse = " ") else NA_character_
    },
    FUN.VALUE = character(1)
  )
  df[[out_parent_col]] <- parent_concat

  if (ncol(final_df)) {
    for (code in colnames(final_df)) {
      outc <- paste0(parent_var, "/", code, "_recod")
      v    <- final_df[[code]]
      df[[outc]] <- ifelse(is.na(v), NA_character_, as.character(v))
      child_cols <- c(child_cols, outc)
    }
  }

  list(
    df       = df,
    new_cols = c(out_parent_col, child_cols, new_child_cols)
  )
}

# =============================
# Merge plantilla->datos con sufijo __tpl
# =============================
ppra_merge_template <- function(data_df, tpl_df){
  kd <- pick_join_key(data_df)
  kt <- pick_join_key(tpl_df)
  if (is.na(kd) || is.na(kt)) stop("No encontré claves de merge compatibles en data/plantilla.")
  tpl2 <- tpl_df
  names(tpl2)[names(tpl2)==kt] <- kd
  other_cols <- setdiff(names(tpl2), kd)
  names(tpl2)[names(tpl2) %in% other_cols] <- paste0(other_cols, "__tpl")
  data_df %>% left_join(tpl2, by = kd)
}

# =============================
# Limpia columnas __tpl
# =============================
# Reemplaza tu función por esta
ppra_strip_tpl_cols <- function(df){
  # Elimina cualquier columna que contenga "__tpl" en el nombre (en cualquier parte)
  keep <- !grepl("__tpl", names(df), perl = TRUE)
  df[, keep, drop = FALSE]
}

# =============================
# Colocación por MAPA (familiasv3.xlsx)
# =============================
.ppra_read_family_map <- function(path_map, sheet = 1){
  stopifnot(file.exists(path_map))
  mp <- readxl::read_excel(path_map, sheet = sheet)
  nms <- tolower(trimws(names(mp)))
  col_parent <- names(mp)[which(nms == "parent")][1]
  col_after  <- names(mp)[which(nms == "text_col")][1]
  if (is.na(col_parent) || is.na(col_after)) {
    stop("En '", basename(path_map), "' se requieren columnas EXACTAS: 'parent' y 'text_col'.")
  }
  mp[[col_parent]] <- trimws(as.character(mp[[col_parent]]))
  mp[[col_after]]  <- trimws(as.character(mp[[col_after]]))
  mp <- mp[mp[[col_parent]] != "" & !is.na(mp[[col_parent]]), , drop = FALSE]
  names(mp)[names(mp)==col_parent] <- "parent"
  names(mp)[names(mp)==col_after]  <- "text_col"
  mp
}

ppra_place_recods_by_map <- function(df, parents, path_map){
  mp <- .ppra_read_family_map(path_map)
  nms <- names(df)

  # helper: reordena columnas recod (padre primero, hijas en orden de bloque crudo, nuevas al final)
  .order_rec_cols <- function(rec_cols, pv, nms_all){
    parent_col <- paste0(pv, "_recod")
    child_rec  <- rec_cols[grepl(paste0("^", pv, "/[^/]+_recod$"), rec_cols, perl = TRUE)]
    other_rec  <- setdiff(rec_cols, c(parent_col, child_rec))  # por si aparece algo más

    # orden del bloque crudo actual (padre+slash), para derivar orden de códigos
    raw_idx <- which(grepl(paste0("^", pv, "($|/[^/]+$)"), nms_all, perl = TRUE) &
                       !grepl("_recod$", nms_all, perl = TRUE))
    raw_cols <- nms_all[raw_idx]
    raw_child <- raw_cols[grepl(paste0("^", pv, "/[^/]+$"), raw_cols, perl = TRUE)]
    raw_codes <- sub("^.+/", "", raw_child)              # códigos crudos en su orden actual

    # códigos de las hijas recod
    rec_codes <- sub("^.+/", "", child_rec)
    rec_codes <- sub("_recod$", "", rec_codes)

    # ordenar hijas recod según el orden crudo; los desconocidos al final
    ord <- match(rec_codes, raw_codes)
    ord_s <- order(ifelse(is.na(ord), 1e9L, ord), rec_codes)  # desconocidos al final, desempate por nombre
    child_rec_sorted <- child_rec[ord_s]

    c(parent_col, child_rec_sorted, other_rec)
  }

  grupos <- list()

  for (pv in parents){
    nms  <- names(df)

    # columnas recod de este parent en el df actual
    rx_rec  <- paste0("^", pv, "(_recod|/[^/]+_recod)$")
    rec_idx <- which(grepl(rx_rec, nms, perl = TRUE))
    if (!length(rec_idx)) next
    rec_cols <- nms[rec_idx]

    # ordenar internamente el bloque recod
    rec_cols <- .order_rec_cols(rec_cols, pv, nms_all = nms)

    # bloque crudo (padre + hijas + *_other) para calcular fin
    raw_idx <- which(grepl(paste0("^", pv, "($|/[^/]+$)"), nms, perl = TRUE) &
                       !grepl("_recod$", nms, perl = TRUE))
    raw_end <- if (length(raw_idx)) max(raw_idx) else if (pv %in% nms) match(pv, nms) else 0L

    # quitar recods para fijar posición de inserción
    cols_wo_rec <- setdiff(nms, rec_cols)
    nms2 <- cols_wo_rec
    raw_idx2 <- which(grepl(paste0("^", pv, "($|/[^/]+$)"), nms2, perl = TRUE) &
                        !grepl("_recod$", nms2, perl = TRUE))
    raw_end2 <- if (length(raw_idx2)) max(raw_idx2) else if (pv %in% nms2) match(pv, nms2) else 0L

    # ancla desde mapa (si existe)
    mpline <- mp %>% dplyr::filter(parent == pv) %>% dplyr::slice(1)
    anchor <- mpline$text_col %||% "__BLOCK_END__"
    anchor_idx <- if (anchor == "__BLOCK_END__") NA_integer_ else match(anchor, nms2)

    # si el ancla cae dentro del bloque crudo, ignoramos ancla y pegamos al final del crudo
    insert_pos <- if (!is.na(anchor_idx)) max(raw_end2, anchor_idx) else raw_end2

    left  <- if (insert_pos <= 0L) character(0) else nms2[seq_len(insert_pos)]
    right <- if (insert_pos <= 0L) nms2 else nms2[(insert_pos+1):length(nms2)]
    new_order <- c(left, rec_cols, right)

    df <- df[, new_order, drop = FALSE]
    grupos[[pv]] <- which(names(df) %in% rec_cols)  # para coloreo
  }

  list(df = df, grupos = grupos)
}

# =============================
# Exportar con color en TODA la columna
# =============================
ppra_export_xlsx <- function(df, path, grupos_cols, colores = TRUE) {
  if (is.null(path)) return(invisible(NULL))
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    warning("No se encontró 'openxlsx'. No se exporta Excel.")
    return(invisible(NULL))
  }

  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "DATA")
  openxlsx::writeData(wb, "DATA", df)

  if (colores && length(grupos_cols)) {
    morado_padre <- "#E6D7FF"  # un poco más suave
    morado_hijas <- "#F2EAFF"  # aún más suave
    for (pv in names(grupos_cols)) {
      idxs  <- sort(grupos_cols[[pv]])
      if (!length(idxs)) next
      # padre = la columna cuyo nombre EXACTO es "<Parent>_recod"
      padre_name <- paste0(pv, "_recod")
      padre <- idxs[match(TRUE, names(df)[idxs] == padre_name)]
      if (is.na(padre)) padre <- idxs[1]

      openxlsx::addStyle(wb, "DATA",
                         openxlsx::createStyle(fgFill = morado_padre),
                         rows = 1:(nrow(df)+1), cols = padre,
                         gridExpand = TRUE, stack = TRUE)

      hijas <- setdiff(idxs, padre)
      if (length(hijas)) {
        openxlsx::addStyle(wb, "DATA",
                           openxlsx::createStyle(fgFill = morado_hijas),
                           rows = 1:(nrow(df)+1), cols = hijas,
                           gridExpand = TRUE, stack = TRUE)
      }
    }
  }

  openxlsx::freezePane(wb, "DATA", firstActiveRow = 2)
  openxlsx::setColWidths(wb, "DATA", cols = 1:ncol(df), widths = "auto")
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
  invisible(path)
}

# =============================
# Plantilla para SELECT-ONE
# =============================
ppra_read_template_sheet_select_one <- function(path_plantilla, parent_var){
  sheet <- ppra_resolve_template_sheet(path_plantilla, parent_var, verbose = TRUE)
  if (is.na(sheet)) stop("Sheet para '", parent_var, "' no encontrado en la plantilla.")
  df <- readxl::read_excel(path_plantilla, sheet = sheet)
  names(df) <- trimws(names(df))
  df
}

# ========= COLOCADOR SELECT-ONE (a la derecha del parent, con mapa opcional) =========
ppra_place_uni_after_parent <- function(df, parent, new_cols, path_map = NULL){
  nms <- names(df)
  rec_cols <- intersect(new_cols, nms)
  if (!length(rec_cols)) return(df)

  # posición de inserción: justo a la derecha del parent (si existe)
  ppos <- match(parent, nms)
  if (is.na(ppos)) ppos <- 0L

  # Si hay mapa y anchor para este parent, úsalo como "mejor esfuerzo"
  if (!is.null(path_map) && file.exists(path_map)) {
    mp <- readxl::read_excel(path_map)
    nmsm <- tolower(trimws(names(mp)))
    col_parent <- names(mp)[which(nmsm == "parent")][1]
    col_after  <- names(mp)[which(nmsm == "text_col")][1]
    if (!is.na(col_parent) && !is.na(col_after)) {
      a <- mp %>% dplyr::filter(!!as.name(col_parent) == parent) %>% dplyr::slice(1)
      if (nrow(a)) {
        anchor <- trimws(as.character(a[[col_after]][1]))
        if (!is.na(anchor) && nzchar(anchor)) {
          aidx <- match(anchor, nms)
          if (!is.na(aidx)) ppos <- max(ppos, aidx)
        }
      }
    }
  }

  nms_wo <- setdiff(nms, rec_cols)
  left   <- if (ppos <= 0L) character(0) else nms_wo[seq_len(ppos)]
  right  <- if (ppos <= 0L) nms_wo else nms_wo[(ppos+1):length(nms_wo)]
  new_order <- c(left, rec_cols, right)
  df[, new_order, drop = FALSE]
}

# ========= CONSTRUCTOR SELECT-ONE (usa choices + plantilla + fallback Other) =========
ppra_build_select_one <- function(df, parent, choices_df, tpl_df){
  nms <- names(df)
  kd <- pick_join_key(df); kt <- pick_join_key(tpl_df)
  if (is.na(kd) || is.na(kt)) stop("No encontré claves comunes df/plantilla en ", parent)
  tpl2 <- tpl_df; names(tpl2)[names(tpl2) == kt] <- kd
  oth <- setdiff(names(tpl2), kd)
  names(tpl2)[names(tpl2) %in% oth] <- paste0(oth, "__tpl")
  df <- dplyr::left_join(df, tpl2, by = kd)
  nms <- names(df)

  # ---- NUEVO: resolver nombres con variantes (con/sin prefijo 'mand_') ----
  base_parent <- sub("^mand_","", parent)  # ej: Assistance_prefered_modality
  get_col <- function(cands) { cn <- cands[cands %in% nms]; if (length(cn)) cn[1] else NA_character_ }

  recod_tpl_col        <- get_col(c(
    paste0(parent,"_RECOD__tpl"),        paste0(parent,"_recod__tpl"),
    paste0(base_parent,"_RECOD__tpl"),   paste0(base_parent,"_recod__tpl")
  ))
  label_recod_tpl_col  <- get_col(c(
    paste0(parent,"_LABEL_RECOD__tpl"),      paste0(parent,"_label_recod__tpl"),
    paste0(base_parent,"_LABEL_RECOD__tpl"), paste0(base_parent,"_label_recod__tpl")
  ))
  label_tpl_col        <- get_col(c(
    paste0(parent,"_LABEL__tpl"),        paste0(parent,"_label__tpl"),
    paste0(base_parent,"_LABEL__tpl"),   paste0(base_parent,"_label__tpl")
  ))
  other_rec_tpl_col    <- get_col(c(
    paste0(parent,"_OTHER_RECOD__tpl"),      paste0(parent,"_other_recod__tpl"),
    paste0(base_parent,"_OTHER_RECOD__tpl"), paste0(base_parent,"_other_recod__tpl"),
    paste0(base_parent,"_other__tpl"),      # por si alguien usó este nombre
    # caso específico que reportaste:
    "Assistance_prefered_modality_other_recod__tpl"
  ))
  other_labrec_tpl_col <- get_col(c(
    paste0(parent,"_OTHER_RECOD_LABEL__tpl"),      paste0(parent,"_other_recod_label__tpl"),
    paste0(base_parent,"_OTHER_RECOD_LABEL__tpl"), paste0(base_parent,"_other_recod_label__tpl"),
    paste0(base_parent,"_other_label__tpl")
  ))
  # -------------------------------------------------------------------------

  base_code <- as.character(df[[parent]])
  base_code[base_code==""] <- NA_character_

  cat_codes <- as.character(choices_df$code)
  cat_labs  <- as.character(choices_df$label)
  lab_by_code <- function(code){ i <- match(code, cat_codes); ifelse(is.na(i), NA_character_, cat_labs[i]) }

  # 1) recod directo para el código
  code_final <- base_code
  if (!is.na(recod_tpl_col)) {
    x <- trimws(as.character(df[[recod_tpl_col]])); x[x==""] <- NA_character_
    repl <- which(!is.na(x)); if (length(repl)) code_final[repl] <- x[repl]
  }

  # 2) si es Other y hay *_other_recod -> usar ese código (puede ser “AlgunodelosDos”)
  is_other <- (!is.na(code_final) & tolower(code_final) == "other") |
    (!is.na(base_code)  & tolower(base_code)  == "other")
  if (any(is_other) && !is.na(other_rec_tpl_col)) {
    tok <- trimws(as.character(df[[other_rec_tpl_col]])); tok[tok==""] <- NA_character_
    fill <- which(is_other & !is.na(tok) & nzchar(tok))
    if (length(fill)) code_final[fill] <- tok[fill]
  }

  # 3) label final: prioridad label_recod, luego catálogo, luego label de plantilla, luego other_label_recod
  label_final <- rep(NA_character_, nrow(df))
  if (!is.na(label_recod_tpl_col)) {
    y <- trimws(as.character(df[[label_recod_tpl_col]])); y[y==""] <- NA_character_
    label_final[!is.na(y)] <- y[!is.na(y)]
  }
  need <- is.na(label_final) & !is.na(code_final)
  if (any(need)) label_final[need] <- lab_by_code(code_final[need])
  need <- is.na(label_final)
  if (any(need) && !is.na(label_tpl_col)) {
    z <- trimws(as.character(df[[label_tpl_col]])); z[z==""] <- NA_character_
    label_final[need & !is.na(z)] <- z[need & !is.na(z)]
  }
  need <- is.na(label_final)
  if (any(need) && !is.na(other_labrec_tpl_col)) {
    w <- trimws(as.character(df[[other_labrec_tpl_col]])); w[w==""] <- NA_character_
    label_final[need & !is.na(w)] <- w[need & !is.na(w)]
  }

  parent_recod_col <- paste0(parent, "_recod")
  parent_lab_col   <- paste0(parent, "_recod_label")

  df[[parent_recod_col]] <- code_final
  df[[parent_lab_col]]   <- label_final

  # limpia __tpl
  keep <- !grepl("__tpl($|\\.[xy]$)", names(df), perl = TRUE)
  df <- df[, keep, drop = FALSE]

  list(df = df, new_cols = c(parent_recod_col, parent_lab_col))
}


# =============================
# Plantilla para OPEN
# =============================

.ppra_norm_multi_tokens <- function(x) {
  x <- as.character(x)
  x <- gsub("[,;|/]+", " ", x, perl = TRUE)  # comas, ;, |, / -> espacio
  x <- gsub("\\s+", " ", x, perl = TRUE)     # espacios múltiples -> 1
  x <- trimws(x)
  x[x == ""] <- NA_character_
  x
}

# Coloca <detail>_recod (y opcional label) a la derecha de <detail>, con mapa como ancla opcional
ppra_place_open_after_details <- function(df, detail_parent, new_cols, path_map = NULL){
  nms <- names(df)
  rec_cols <- intersect(new_cols, nms)
  if (!length(rec_cols)) return(df)

  ppos <- match(detail_parent, nms); if (is.na(ppos)) ppos <- 0L

  # Soporte de anclaje por mapa (columna text_col si quieres empujar más a la derecha)
  if (!is.null(path_map) && file.exists(path_map)) {
    mp <- readxl::read_excel(path_map)
    nmsm <- tolower(trimws(names(mp)))
    col_parent <- names(mp)[which(nmsm == "parent")][1]
    col_after  <- names(mp)[which(nmsm == "text_col")][1]
    if (!is.na(col_parent) && !is.na(col_after)) {
      a <- mp %>% dplyr::filter(!!as.name(col_parent) == detail_parent) %>% dplyr::slice(1)
      if (nrow(a)) {
        anchor <- trimws(as.character(a[[col_after]][1]))
        if (!is.na(anchor) && nzchar(anchor)) {
          aidx <- match(anchor, nms); if (!is.na(aidx)) ppos <- max(ppos, aidx)
        }
      }
    }
  }

  nms_wo <- setdiff(nms, rec_cols)
  left   <- if (ppos <= 0L) character(0) else nms_wo[seq_len(ppos)]
  right  <- if (ppos <= 0L) nms_wo else nms_wo[(ppos+1):length(nms_wo)]
  new_order <- c(left, rec_cols, right)
  df[, new_order, drop = FALSE]
}


# Limpia y busca hoja de plantilla compatible con TEXT_<detail>
ppra_resolve_text_sheet <- function(path_plantilla, detail_var){
  if (!file.exists(path_plantilla)) return(NA_character_)
  sh <- readxl::excel_sheets(path_plantilla)
  if (!length(sh)) return(NA_character_)
  cand <- paste0("TEXT_", detail_var)
  # exacto / insensitive / truncado 31 / limpieza / similitud
  .cl <- function(x) gsub("[^A-Za-z0-9]+","", tolower(x))
  cands <- c(cand, substr(cand,1,31))
  hit <- which(tolower(sh) %in% tolower(cands))
  if (length(hit)) return(sh[hit[1]])
  cls <- .cl(sh); clc <- .cl(cand)
  hit <- which(cls == clc)
  if (length(hit)) return(sh[hit[1]])
  d <- adist(clc, cls); j <- which.min(d)
  if (length(j) && is.finite(d[j]) && d[j] <= 5) return(sh[j])
  NA_character_
}

ppra_build_open_text <- function(df, detail_var, path_plantilla){
  nms <- names(df)
  if (!detail_var %in% nms) stop("No encuentro la columna de texto: ", detail_var)

  # Merge plantilla TEXT_<detail>
  sheet <- ppra_resolve_text_sheet(path_plantilla, detail_var)
  if (!is.na(sheet)) {
    tpl <- readxl::read_excel(path_plantilla, sheet = sheet)
    names(tpl) <- trimws(names(tpl))
    kd <- pick_join_key(df); kt <- pick_join_key(tpl)
    if (!is.na(kd) && !is.na(kt)) {
      names(tpl)[names(tpl)==kt] <- kd
      oth <- setdiff(names(tpl), kd)
      names(tpl)[names(tpl) %in% oth] <- paste0(oth, "__tpl")
      df <- dplyr::left_join(df, tpl, by = kd)
      nms <- names(df)
    }
  }

  # Detectar columnas recod/label posibles en plantilla
  base <- detail_var
  get_col <- function(cands) { cn <- cands[cands %in% nms]; if (length(cn)) cn[1] else NA_character_ }
  rec_tpl   <- get_col(c(paste0(base, "_RECOD__tpl"), paste0(base, "_recod__tpl")))
  lab_tpl   <- get_col(c(paste0(base, "_LABEL__tpl"), paste0(base, "_label__tpl")))
  # (opcional) algunos definen *_recod_label__tpl
  lab_rec_tpl <- get_col(c(paste0(base, "_RECOD_LABEL__tpl"), paste0(base, "_recod_label__tpl")))

  # Construcción
  rec_col <- paste0(detail_var, "_recod")
  lab_col <- paste0(detail_var, "_recod_label")

  code_final <- rep(NA_character_, nrow(df))
  if (!is.na(rec_tpl)) {
    x <- .ppra_norm_multi_tokens(df[[rec_tpl]])
    x <- vapply(strsplit(ifelse(is.na(x), "", x), "\\s+"),
                function(tok){
                  tok <- tok[nzchar(tok)]
                  if (!length(tok)) return(NA_character_)
                  tok <- unique(tok)          # sin duplicados, respeta orden de aparición
                  paste(tok, collapse = " ")
                }, FUN.VALUE = character(1))
    code_final[!is.na(x) & nzchar(x)] <- x[!is.na(x) & nzchar(x)]
  }
  label_final <- rep(NA_character_, nrow(df))
  if (!is.na(lab_rec_tpl)) {
    y <- trimws(as.character(df[[lab_rec_tpl]])); y[y==""] <- NA_character_
    label_final[!is.na(y)] <- y[!is.na(y)]
  }
  need <- is.na(label_final)
  if (any(need) && !is.na(lab_tpl)) {
    z <- trimws(as.character(df[[lab_tpl]])); z[z==""] <- NA_character_
    label_final[need & !is.na(z)] <- z[need & !is.na(z)]
  }

  df[[rec_col]] <- code_final
  if (any(!is.na(label_final))) df[[lab_col]] <- label_final

  # Limpieza __tpl
  keep <- !grepl("__tpl($|\\.[xy]$)", names(df), perl = TRUE)
  df <- df[, keep, drop = FALSE]

  # ==> FIX: armar new_cols correctamente
  new_cols <- c(rec_col)
  if (lab_col %in% names(df)) new_cols <- c(new_cols, lab_col)
  list(df = df, new_cols = new_cols)
}

ppra_join_text_details <- function(df, path_plantilla, parent) {
  df
}

# ========= EXPORT con triple paleta =========
ppra_export_xlsx2 <- function(df, path,
                              grupos_multi = list(),
                              grupos_uni   = list(),
                              grupos_open  = list(),
                              colores      = TRUE){

  if (is.null(path)) return(invisible(NULL))
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    warning("No se encontró 'openxlsx'. No se exporta Excel.")
    return(invisible(NULL))
  }

  # Paletas (padre / hijas)
  verde_padre   <- "#DFF5DF"; verde_hijas   <- "#EFFAEF"  # multiple (SM)
  azul_padre    <- "#DCEBFF"; azul_hijas    <- "#EEF5FF"  # one (SO)
  naranja_padre <- "#FFE6CC"; naranja_hijas <- "#FFF1E0"  # open (TEXT_)

  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "DATA")
  openxlsx::writeData(wb, "DATA", df)

  # --- Resolver a índices, acepte nombres o índices ---
  nms <- names(df)
  .to_idx <- function(x) {
    if (is.null(x) || !length(x)) return(integer(0))
    if (is.numeric(x)) {
      ix <- x[is.finite(x) & x >= 1 & x <= length(nms)]
      unique(as.integer(ix))
    } else {
      # caracteres -> match a nombres
      ix <- match(x, nms)
      unique(as.integer(stats::na.omit(ix)))
    }
  }

  .paint_block <- function(block, parent_name = NULL, col_parent, col_child){
    idxs <- .to_idx(block)
    if (!length(idxs)) return(0L)
    # padre: si nos dieron parent_name, úsalo; si no, toma el primero del bloque
    if (!is.null(parent_name) && parent_name %in% nms) {
      pidx <- match(parent_name, nms)
    } else {
      pidx <- idxs[1]
    }
    # pinta padre
    openxlsx::addStyle(
      wb, "DATA",
      openxlsx::createStyle(fgFill = col_parent),
      rows = 1:(nrow(df)+1), cols = pidx,
      gridExpand = TRUE, stack = TRUE
    )
    # pinta hijas
    child_idxs <- setdiff(idxs, pidx)
    if (length(child_idxs)) {
      openxlsx::addStyle(
        wb, "DATA",
        openxlsx::createStyle(fgFill = col_child),
        rows = 1:(nrow(df)+1), cols = child_idxs,
        gridExpand = TRUE, stack = TRUE
      )
    }
    length(idxs)
  }

  if (colores) {
    # MULTIPLE (verde)
    for (pv in names(grupos_multi)) {
      block <- grupos_multi[[pv]]  # puede ser nombres o índices
      parent_nm <- paste0(pv, "_recod")
      coloured <- .paint_block(block, parent_nm, verde_padre, verde_hijas)
      # message opcional de diagnóstico:
      message(sprintf("[paint][SM] %s: %d col(s)", pv, coloured))
    }
    # ONE (azul)
    for (pv in names(grupos_uni)) {
      block <- grupos_uni[[pv]]
      parent_nm <- paste0(pv, "_recod")
      coloured <- .paint_block(block, parent_nm, azul_padre, azul_hijas)
      message(sprintf("[paint][SO] %s: %d col(s)", pv, coloured))
    }
    # OPEN (naranja)
    for (ov in names(grupos_open)) {
      block <- grupos_open[[ov]]
      parent_nm <- paste0(ov, "_recod")
      coloured <- .paint_block(block, parent_nm, naranja_padre, naranja_hijas)
      message(sprintf("[paint][OPEN] %s: %d col(s)", ov, coloured))
    }
  }

  openxlsx::freezePane(wb, "DATA", firstActiveRow = 2)
  openxlsx::setColWidths(wb, "DATA", cols = 1:ncol(df), widths = "auto")
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
  invisible(path)
}

# ======================================================
# === Paso final: dejar solo padres _recod (y quitar hijas y labels) ===
# =======================================================


ppra_final_slim <- function(df, out_path = NULL) {
  nms <- names(df)

  # Marcas de borrado
  drop_children <- grepl(".+/.+_recod$", nms, perl = TRUE)   # ej: Parent/Code_recod
  drop_labels   <- grepl("_recod_label$", nms, perl = TRUE)  # ej: <var>_recod_label

  keep_idx <- !(drop_children | drop_labels)
  df_slim  <- df[, keep_idx, drop = FALSE]

  dropped <- tibble::tibble(
    column = nms[!keep_idx],
    reason = dplyr::case_when(
      drop_children[!keep_idx] ~ "child_recod",
      drop_labels[!keep_idx]   ~ "recod_label",
      TRUE ~ "other"
    )
  )

  if (!is.null(out_path)) {
    if (!requireNamespace("openxlsx", quietly = TRUE)) {
      warning("No se encontró 'openxlsx'. No se exporta Excel.")
    } else {
      wb <- openxlsx::createWorkbook()
      openxlsx::addWorksheet(wb, "DATA")
      openxlsx::writeData(wb, "DATA", df_slim)
      openxlsx::freezePane(wb, "DATA", firstActiveRow = 2)
      openxlsx::setColWidths(wb, "DATA", cols = 1:ncol(df_slim), widths = "auto")
      openxlsx::saveWorkbook(wb, out_path, overwrite = TRUE)
    }
  }

  list(
    data    = df_slim,
    dropped = dropped
  )
}

# Devuelve los nombres de columnas padre a pintar (no índices),
# inferidos de los grupos construidos durante la fase “completa”.
ppra_collect_parent_targets <- function(df, grupos_multi, grupos_uni, grupos_open){
  nms <- names(df)

  vec_or_empty <- function(x) if (is.null(x)) character(0) else x

  # Basado en el nombre del parent en los grupos
  m_parents <- names(vec_or_empty(grupos_multi))
  u_parents <- names(vec_or_empty(grupos_uni))
  o_parents <- names(vec_or_empty(grupos_open))

  multi_cols <- paste0(m_parents, "_recod")
  multi_cols <- multi_cols[multi_cols %in% nms]
  names(multi_cols) <- sub("_recod$","", multi_cols)

  one_cols   <- paste0(u_parents, "_recod")
  one_cols   <- one_cols[one_cols %in% nms]
  names(one_cols) <- sub("_recod$","", one_cols)

  open_cols  <- paste0(o_parents, "_recod")
  open_cols  <- open_cols[open_cols %in% nms]
  names(open_cols) <- sub("_recod$","", open_cols)

  list(multi = multi_cols, one = one_cols, open = open_cols)
}


# ========= FUNCIÓN PRINCIPAL UNIFICADA (con slim integrado) =========
#' Aplica recodificación para select-multiple (parent_var), select-one (uni_var)
#' y abiertos (open_var). Opcionalmente aplica un “slim” final que elimina
#' todas las hijas *_recod y todas las *_recod_label, dejando solo los padres
#' <Parent>_recod (y todas las columnas originales).
#'
#' @param path_instrumento XLSForm (survey/choices)
#' @param path_datos Base original
#' @param path_plantilla Excel de plantilla con hojas por parent (31-chars safe)
#' @param parent_var vector de select-multiple a procesar
#' @param uni_var vector de select-one a procesar
#' @param open_var vector de “abiertas” a procesar (detalles de texto)
#' @param path_map Excel familias (columns: parent, text_col) para anclaje (opcional)
#' @param out_path ruta de salida XLSX “completo” (con bloques y color)
#' @param colores TRUE para colorear el out_path “completo”
#' @param slim si TRUE, ejecuta recorte final para dejar solo padres _recod
#' @param out_path_final ruta de salida XLSX “slim” (opcional, solo padres _recod)
#' @return list con data (completa), final_data (slim si aplica) y metadatos
#' @export
ppra_adaptar_data <- function(path_instrumento,
                         path_datos,
                         path_plantilla,
                         parent_var = character(0),
                         uni_var    = character(0),
                         open_var   = character(0),
                         path_map   = "familiasv3.xlsx",
                         out_path   = NULL,
                         colores    = TRUE,
                         slim       = TRUE,
                         out_path_final = NULL) {

  df <- leer_datos_generico(path_datos)

  grupos_multi <- list()
  grupos_uni   <- list()
  grupos_open  <- list()

  # ===== SELECT MULTIPLE =====
  if (length(parent_var)) {
    columnas_nuevas <- list()
    procesadas <- character(0)

    for (pv in parent_var) {
      ch  <- ppra_get_choices_parent(path_instrumento, pv)
      tpl <- ppra_read_template_sheet(path_plantilla, pv)
      df  <- ppra_merge_template(df, tpl$df)

      built <- ppra_build_parent_block(df, pv, choices_df = ch)
      df    <- built$df
      columnas_nuevas[[pv]] <- built$new_cols
      procesadas <- c(procesadas, pv)
    }

    df <- ppra_strip_tpl_cols(df)
    df <- ppra_enforce_all_recods(df)
    df <- ppra_parent_rebuild_fallback(df)

    placed <- ppra_place_recods_by_map(df, unique(procesadas), path_map = path_map)
    df <- placed$df
    # Reconstruir grupos_multi como NOMBRES:
    grupos_multi <- list()
    for (pv in unique(procesadas)) {
      parent_col <- paste0(pv, "_recod")
      child_rec  <- names(df)[grepl(paste0("^", pv, "/[^/]+_recod$"), names(df))]
      grupos_multi[[pv]] <- c(parent_col, child_rec)  # <-- NOMBRES
    }
  }

  # ===== SELECT ONE =====
  if (length(uni_var)) {
    for (pv in uni_var) {
      ch  <- ppra_get_choices_parent(path_instrumento, pv)

      # intentar hoja normal; si falla, fallback TEXT_<base>_other
      ok <- TRUE
      tpl <- try(ppra_read_template_sheet_select_one(path_plantilla, pv), silent = TRUE)
      if (inherits(tpl, "try-error")) {
        ok <- FALSE
        # fallback “other” basado en XLSForm
        msg <- conditionMessage(attr(tpl, "condition"))
        base <- sub("^mand_", "", pv)
        fb_sheet <- paste0("TEXT_", base, "_other")
        sheets <- try(readxl::excel_sheets(path_plantilla), silent = TRUE)
        has_fb <- !inherits(sheets, "try-error") && any(tolower(sheets) == tolower(fb_sheet))
        if (has_fb) {
          message("No hay hoja para '", pv, "'; usando fallback ", fb_sheet, ". Motivo: ", msg)
          # crea un tpl vacío (no forzará recods, pero permite el flujo)
          tpl <- list(df = tibble::tibble())
          ok <- TRUE
        } else {
          message("No hay hoja para '", pv, "'; sin fallback disponible. Motivo: ", msg)
        }
      }

      if (ok) {
        res <- ppra_build_select_one(df, pv, ch, tpl)
        df  <- res$df
        df  <- ppra_place_uni_after_parent(df, parent = pv, new_cols = res$new_cols, path_map = path_map)
        grupos_uni[[pv]] <- res$new_cols   # <-- NOMBRES, no índices
      }
    }
  }

  # ===== OPEN (TEXT_* detalles) =====
  if (length(open_var)) {
    for (ov in open_var) {
      res <- ppra_build_open_text(df, detail_var = ov, path_plantilla = path_plantilla)
      df  <- res$df
      df  <- ppra_place_open_after_details(df, detail_parent = ov, new_cols = res$new_cols, path_map = path_map)
      grupos_open[[ov]] <- res$new_cols  # <-- NOMBRES, no índices
    }
  }
  # ===== Export completo (con color) =====
  if (!is.null(out_path)) {
    ppra_export_xlsx2(df,
                      path = out_path,
                      grupos_multi = grupos_multi,
                      grupos_uni   = grupos_uni,
                      grupos_open  = grupos_open,
                      colores      = colores)
  }

  # <<< RECUERDA QUÉ PADRES PINTASTE (por nombre) PARA USARLOS TRAS EL SLIM >>>
  paint_targets <- ppra_collect_parent_targets(df, grupos_multi, grupos_uni, grupos_open)

  # ===== Slim final opcional =====
  final_data   <- NULL
  final_dropped <- NULL
  if (isTRUE(slim)) {
    fin <- ppra_final_slim(df, out_path = NULL)   # no exportes aquí dentro
    final_data    <- fin$data
    final_dropped <- fin$dropped

    # <<< pintar el Excel FINAL por NOMBRE de padres _recod >>>
    if (!is.null(out_path_final)) {
      nms_fin <- names(final_data)

      # cada entrada es un bloque de 1 col (el padre _recod)
      mk_block_list <- function(named_vec_cols) {
        # named_vec_cols: vector con nombres = parent base, valores = "<parent>_recod"
        keep <- named_vec_cols[named_vec_cols %in% nms_fin]
        setNames(lapply(as.vector(keep), function(col) col), names(keep))
      }

      slim_multi <- mk_block_list(paint_targets$multi)
      slim_one   <- mk_block_list(paint_targets$one)
      slim_open  <- mk_block_list(paint_targets$open)

      ppra_export_xlsx2(final_data,
                        path         = out_path_final,
                        grupos_multi = slim_multi,
                        grupos_uni   = slim_one,
                        grupos_open  = slim_open,
                        colores      = TRUE)  # pinta siempre el final
    }
  }
}


# =====================================================================
# Auditor unificado (select_multiple + select_one)
# Requiere: readxl, dplyr, tibble
# =====================================================================
suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(tibble)
})

# --- Helpers locales mínimos que usa el auditor ---
.ppra_clean  <- function(x) gsub("[^A-Za-z0-9]+","", tolower(trimws(as.character(x))))
.ppra_where_raw <- function(df, pv) which(grepl(paste0("^",pv,"($|/[^/]+$)"), names(df)) & !grepl("_recod$", names(df)))
.ppra_where_rec <- function(df, pv) which(grepl(paste0("^",pv,"(_recod|/[^/]+_recod)$"), names(df)))
.ppra_where_uni_raw <- function(df, pv) match(pv, names(df))
.ppra_where_uni_rec <- function(df, pv) which(names(df) %in% c(paste0(pv,"_recod"), paste0(pv,"_recod_label")))
.ppra_norm_multi_tokens <- function(x){
  x <- as.character(x)
  x <- gsub("[,;|/]+"," ", x, perl = TRUE)
  x <- gsub("\\s+"," ", x, perl = TRUE)
  x <- trimws(x); x[x==""] <- NA_character_; x
}

# --- Detección de hojas fallback TEXT_<parent>_other en la plantilla (para ONE) ---
.ppra_exists_sheet <- function(sh, name){
  cand <- c(name, substr(name,1,31))
  i <- which(tolower(sh) %in% tolower(cand))
  if (length(i)) return(sh[i[1]])
  cs <- .ppra_clean(sh); cn <- .ppra_clean(name)
  i <- which(cs == cn); if (length(i)) return(sh[i[1]])
  d <- adist(cn, cs); j <- which.min(d)
  if (length(j) && is.finite(d[j]) && d[j] <= 5) return(sh[j])
  NA_character_
}

# --- Faltantes: operador %||% y detectores de SM/SO desde el XLSForm ---

`%||%` <- function(a, b) if (is.null(a) || (length(a) == 1 && is.na(a))) b else a

ppra_detect_sm_parents <- function(path_instrumento){
  stopifnot(file.exists(path_instrumento))
  survey <- readxl::read_excel(path_instrumento, sheet = "survey")

  nms <- tolower(names(survey))
  type_col <- names(survey)[match("type", nms)]
  name_col <- names(survey)[match("name", nms)]
  if (is.na(type_col) || is.na(name_col))
    stop("No encuentro columnas 'type' y/o 'name' en la hoja 'survey'.")

  ty <- tolower(trimws(as.character(survey[[type_col]])))
  nm <- trimws(as.character(survey[[name_col]]))

  idx <- grepl("^select[ _]?multiple\\b", ty)
  unique(nm[idx & nzchar(nm)])
}

ppra_detect_so_parents <- function(path_instrumento){
  stopifnot(file.exists(path_instrumento))
  survey <- readxl::read_excel(path_instrumento, sheet = "survey")

  nms <- tolower(names(survey))
  type_col <- names(survey)[match("type", nms)]
  name_col <- names(survey)[match("name", nms)]
  if (is.na(type_col) || is.na(name_col))
    stop("No encuentro columnas 'type' y/o 'name' en la hoja 'survey'.")

  ty <- tolower(trimws(as.character(survey[[type_col]])))
  nm <- trimws(as.character(survey[[name_col]]))

  idx <- grepl("^select[ _]?one\\b", ty)
  unique(nm[idx & nzchar(nm)])
}

# ================= Auditar tomando directamente el Slim =================
suppressPackageStartupMessages({
  library(readxl); library(dplyr); library(tibble); library(stringr)
})

# -- Helpers de selección de entrada ------------------------------------
.ppra_pick_df <- function(x, prefer = c("final","data","auto")) {
  prefer <- match.arg(prefer)
  # 1) Si viene una lista tipo ppra_adaptar_data(...)
  if (is.list(x) && !is.data.frame(x)) {
    if (prefer == "final" && !is.null(x$final_data) && is.data.frame(x$final_data)) return(x$final_data)
    if (!is.null(x$data) && is.data.frame(x$data)) return(x$data)
    stop("La lista no contiene 'data' ni 'final_data' como data.frame.")
  }
  # 2) Si viene un path a xlsx
  if (is.character(x) && length(x) == 1 && grepl("\\.xlsx?$", x, ignore.case = TRUE)) {
    if (!file.exists(x)) stop("No encuentro el archivo: ", x)
    return(readxl::read_xlsx(x))
  }
  # 3) Si ya es un data.frame
  if (is.data.frame(x)) return(x)

  stop("Entrada no reconocida. Pase un data.frame, una lista de ppra_adaptar_data(), o una ruta .xlsx.")
}

.ppra_pick_expected <- function(x, parent_var, uni_var, open_var) {
  # Si x es la lista del orquestador y no pasaron vectores, usa los de la lista
  if (is.list(x) && !is.data.frame(x)) {
    if (is.null(parent_var) && !is.null(x$multi_procesadas)) parent_var <- x$multi_procesadas
    if (is.null(uni_var)    && !is.null(x$one_procesadas))   uni_var    <- x$one_procesadas
    if (is.null(open_var)   && !is.null(x$open_procesadas))  open_var   <- x$open_procesadas
  }
  list(parent_var = parent_var %||% character(0),
       uni_var    = uni_var    %||% character(0),
       open_var   = open_var   %||% character(0))
}

# -- Sus helpers mínimos del auditor (igual que antes) -------------------
`%||%` <- function(a,b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a
.ppra_clean  <- function(x) gsub("[^A-Za-z0-9]+","", tolower(trimws(as.character(x))))
.ppra_where_raw <- function(df, pv) which(grepl(paste0("^",pv,"($|/[^/]+$)"), names(df)) & !grepl("_recod$", names(df)))
.ppra_where_rec <- function(df, pv) which(grepl(paste0("^",pv,"(_recod|/[^/]+_recod)$"), names(df)))
.ppra_where_uni_raw <- function(df, pv) match(pv, names(df))
.ppra_where_uni_rec <- function(df, pv) which(names(df) %in% c(paste0(pv,"_recod"), paste0(pv,"_recod_label")))
.ppra_norm_multi_tokens <- function(x){
  x <- as.character(x); x <- gsub("[,;|/]+"," ", x); x <- gsub("\\s+"," ", x)
  x <- trimws(x); x[x==""] <- NA_character_; x
}

# -- Detectores desde XLSForm (usa los suyos existentes) -----------------
# Debe tener estas funciones definidas en su entorno como antes:
#   ppra_detect_sm_parents(path_instrumento)
#   ppra_detect_so_parents(path_instrumento)

# -- Auditor unificado que acepta lista / df / ruta ----------------------
#' Auditoría unificada de variables (SM, SO y OPEN) con soporte para Slim
#'
#' Esta función permite auditar la estructura de recodificación de un dataset
#' generado por \code{\link{ppra_adaptar_data}}. Es flexible en la entrada:
#' puede recibir directamente un \code{data.frame}, una lista retornada por
#' \code{ppra_adaptar_data()}, o una ruta a un archivo \code{.xlsx}.
#'
#' - Si se pasa una lista de \code{ppra_adaptar_data()}, por defecto se audita
#'   el producto \code{final_data} ("slim"). Si no existe, se usa \code{data}.
#' - Si se pasa un \code{.xlsx}, la función lo carga con \code{readxl}.
#' - Si se pasa un \code{data.frame}, se usa directamente.
#'
#' El reporte generado cubre:
#' - Select-multiple (SM): verifica presencia de columnas hijas, bloques recodificados,
#'   orden relativo, proporción de filas con datos, etc.
#' - Select-one (SO): idem, además de detectar recodificación vía \code{_recod} y
#'   \code{_recod_label}.
#' - Abiertas (OPEN): genera sugerencias de si deberían ser tratadas como
#'   \code{one} o \code{multiple}, en base a la distribución de tokens.
#'
#' @param df_or_res Entrada a auditar: un \code{data.frame}, una lista producida
#'   por \code{ppra_adaptar_data()}, o una ruta a un archivo Excel (.xlsx).
#' @param path_instrumento Ruta al XLSForm (survey/choices) para detectar padres
#'   de select-multiple y select-one.
#' @param parent_var Vector opcional de nombres de variables SM esperadas.
#'   Si no se pasa y la entrada es una lista, se usan los de \code{multi_procesadas}.
#' @param uni_var Vector opcional de variables SO esperadas.
#'   Si no se pasa y la entrada es lista, se usan los de \code{one_procesadas}.
#' @param open_var Vector opcional de variables abiertas a evaluar.
#'   Si no se pasa y la entrada es lista, se usan los de \code{open_procesadas}.
#' @param allowed_gap Número máximo de columnas permitidas entre el bloque crudo
#'   y el bloque recodificado (default = 2).
#' @param include_one_fallback ¿Anotar posibles fallback de ONE desde plantilla?
#'   (default = TRUE, sólo si se integra esa lógica).
#' @param include_open_suggestions ¿Generar bloque de sugerencias OPEN?
#'   (default = TRUE).
#' @param min_single Umbral mínimo de proporción de filas con 1 token para sugerir ONE
#'   (default = 0.80).
#' @param min_multi Umbral mínimo de proporción de filas con >1 token para sugerir MULTIPLE
#'   (default = 0.10).
#' @param max_unique_for_one Número máximo de categorías únicas aceptado para sugerir ONE
#'   (default = 20).
#' @param prefer Qué dataset usar si se pasa una lista: \code{"final"} (slim),
#'   \code{"data"} (completo) o \code{"auto"} (elige slim si existe).
#'
#' @return Un \code{tibble} con filas para cada variable auditada:
#'   \itemize{
#'     \item \code{type}: "multiple", "one" o "open_suggest".
#'     \item \code{parent}: nombre de la variable auditada.
#'     \item \code{esperado}: si estaba en los vectores declarados.
#'     \item \code{raw_present}, \code{rec_present}: indicadores de columnas presentes.
#'     \item \code{raw_range}, \code{rec_range}: rangos de posiciones de columnas.
#'     \item \code{gap}, \code{ok_order}: métricas de orden/bloque.
#'     \item \code{status}: estado ("codificado", "no_codificado", "missing_recod",
#'       "one", "multiple", "undetermined").
#'     \item \code{note}: anotaciones (ej. fallback, número de tokens únicos).
#'   }
#'
#' @examples
#' \dontrun{
#' # Usando lista de ppra_adaptar_data
#' res <- ppra_adaptar_data(...)
#' aud <- ppra_auditar(res, path_instrumento = "instrumento.xlsx")
#'
#' # Usando ruta a Excel final
#' aud <- ppra_auditar("DATA_ADAPTADA_SM_FINAL.xlsx", "instrumento.xlsx")
#'
#' # Usando un data.frame en memoria
#' df <- readxl::read_xlsx("DATA_ADAPTADA_SM_FINAL.xlsx")
#' aud <- ppra_auditar(df, "instrumento.xlsx")
#' }
#'
#' @export
ppra_auditar <- function(df_or_res,
                         path_instrumento,
                         parent_var = NULL,
                         uni_var    = NULL,
                         open_var   = NULL,
                         allowed_gap = 2,
                         include_one_fallback = TRUE,
                         include_open_suggestions = TRUE,
                         min_single = 0.80,
                         min_multi  = 0.10,
                         max_unique_for_one = 20,
                         prefer = c("final","data","auto")) {

  prefer <- match.arg(prefer)

  # 0) Resolver df según entrada (usa Slim por defecto)
  df <- .ppra_pick_df(df_or_res, prefer = if (prefer=="auto") "final" else prefer)

  # 0b) Si no pasaron expected, tomar los de la lista si df_or_res era lista
  ex <- .ppra_pick_expected(df_or_res, parent_var, uni_var, open_var)
  parent_var <- ex$parent_var; uni_var <- ex$uni_var; open_var <- ex$open_var

  # 1) Universo SM/SO desde el XLSForm
  sm_parents <- unique(ppra_detect_sm_parents(path_instrumento))
  so_parents <- unique(ppra_detect_so_parents(path_instrumento))

  expected_sm <- unique(parent_var %||% character(0))
  expected_so <- unique(uni_var    %||% character(0))
  nms <- names(df)

  # ===== SM =====
  sm_rows <- lapply(sm_parents, function(pv){
    raw_idx <- .ppra_where_raw(df, pv)
    rec_idx <- .ppra_where_rec(df, pv)
    raw_range <- if (length(raw_idx)) paste(range(raw_idx), collapse = "-") else NA_character_
    rec_range <- if (length(rec_idx)) paste(range(rec_idx), collapse = "-") else NA_character_
    gap <- if (length(raw_idx) && length(rec_idx)) min(rec_idx) - max(raw_idx) else NA_integer_
    ok_order <- (length(raw_idx) > 0 && length(rec_idx) > 0 && !is.na(gap) && gap >= 1 && gap <= allowed_gap)
    n_raw_children <- sum(grepl(paste0("^", pv, "/[^/]+$"), nms) & !grepl("_recod$", nms))
    n_rec_children <- sum(grepl(paste0("^", pv, "/[^/]+_recod$"), nms))
    parent_col <- paste0(pv, "_recod")
    has_parent <- parent_col %in% nms
    filled_prop <- if (has_parent) {
      v <- trimws(as.character(df[[parent_col]])); mean(!is.na(v) & nzchar(v))
    } else NA_real_
    status <- if (n_rec_children > 0 || has_parent) "codificado" else "no_codificado"

    tibble(
      type="multiple", parent=pv, esperado = pv %in% expected_sm,
      raw_present = length(raw_idx) > 0, rec_present = length(rec_idx) > 0,
      raw_range, rec_range, gap, ok_order,
      n_raw_children, n_rec_children,
      has_parent_recod = has_parent, parent_recod_fill_p = filled_prop,
      status, note = NA_character_
    )
  }) %>% bind_rows()

  # ===== SO =====
  # (omito fallback-per-plantilla aquí por brevedad; puede mantener su versión larga)
  so_rows <- lapply(so_parents, function(pv){
    raw_idx <- .ppra_where_uni_raw(df, pv)
    rec_idx <- .ppra_where_uni_rec(df, pv)
    raw_range <- if (!is.na(raw_idx)) paste(raw_idx, raw_idx, sep="-") else NA_character_
    rec_range <- if (length(rec_idx)) paste(range(rec_idx), collapse="-") else NA_character_
    gap <- if (!is.na(raw_idx) && length(rec_idx)) min(rec_idx) - raw_idx else NA_integer_
    ok_order <- (!is.na(raw_idx) && length(rec_idx) > 0 && !is.na(gap) && gap >= 1 && gap <= allowed_gap)
    recod_col <- paste0(pv, "_recod")
    lab_col   <- paste0(pv, "_recod_label")
    has_recod <- recod_col %in% nms
    filled_prop <- if (has_recod) { v <- trimws(as.character(df[[recod_col]])); mean(!is.na(v) & nzchar(v)) } else NA_real_
    status <- if (has_recod || (lab_col %in% nms)) "codificado" else "no_codificado"

    tibble(
      type="one", parent=pv, esperado = pv %in% expected_so,
      raw_present = !is.na(raw_idx), rec_present = length(rec_idx) > 0,
      raw_range, rec_range, gap, ok_order,
      n_raw_children = NA_integer_, n_rec_children = NA_integer_,
      has_parent_recod = has_recod, parent_recod_fill_p = filled_prop,
      status, note = NA_character_
    )
  }) %>% bind_rows()

  # ===== Sugerencias OPEN (si las pasó) =====
  open_rows <- tibble()
  if (length(open_var)) {
    open_rows <- lapply(open_var, function(ov){
      raw_idx <- match(ov, nms)
      rec_idx <- match(paste0(ov,"_recod"), nms)
      raw_range <- if (!is.na(raw_idx)) paste(raw_idx, raw_idx, sep="-") else NA_character_
      rec_range <- if (!is.na(rec_idx)) paste(rec_idx, rec_idx, sep="-") else NA_character_

      if (is.na(rec_idx)) {
        return(tibble(
          type="open_suggest", parent=ov, esperado = TRUE,
          raw_present = !is.na(raw_idx), rec_present = FALSE,
          raw_range, rec_range, gap = NA_integer_, ok_order = NA,
          n_raw_children = NA_integer_, n_rec_children = NA_integer_,
          has_parent_recod = FALSE, parent_recod_fill_p = NA_real_,
          status = "missing_recod", note = NA_character_
        ))
      }

      v <- .ppra_norm_multi_tokens(df[[rec_idx]])
      filled <- !is.na(v); filled_p <- mean(filled)
      toks <- strsplit(ifelse(is.na(v),"",v), "\\s+")
      n_tok <- vapply(toks, function(z) sum(nzchar(z)), integer(1))
      p_single <- ifelse(any(filled), mean(n_tok[filled]==1), NA_real_)
      p_multi  <- ifelse(any(filled), mean(n_tok[filled]> 1), NA_real_)
      all_tok  <- unique(unlist(toks)); all_tok <- all_tok[nzchar(all_tok)]
      uniq_n   <- length(all_tok)

      suggestion <- dplyr::case_when(
        is.na(p_single) ~ "no_data",
        p_multi >= min_multi ~ "multiple",
        p_single >= min_single & uniq_n <= max_unique_for_one ~ "one",
        TRUE ~ "undetermined"
      )

      tibble(
        type="open_suggest", parent=ov, esperado = TRUE,
        raw_present = !is.na(raw_idx), rec_present = TRUE,
        raw_range, rec_range, gap = NA_integer_, ok_order = NA,
        n_raw_children = NA_integer_, n_rec_children = NA_integer_,
        has_parent_recod = FALSE, parent_recod_fill_p = round(filled_p,3),
        status = suggestion, note = paste0("unique_tokens=", uniq_n)
      )
    }) %>% bind_rows()
  }

  bind_rows(sm_rows, so_rows, open_rows) %>%
    arrange(desc(esperado), type, parent)
}
