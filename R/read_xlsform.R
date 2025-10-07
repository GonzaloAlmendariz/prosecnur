# ============================================================
# LECTOR "SIN NORMALIZAR" PARA REGLAS / PLAN DE LIMPIEZA — ONE-FILE v2.2
# - Soporta begin_group/end_group y begin_repeat/end_repeat
# - Captura choice_filter y variables referenciadas
# - groups_detail con glabel (label del begin_group; cae en gname si falta)
# - section_map dentro de meta (prefix, is_conditional, is_repeat)
# - choice_cols_by_list depurado (solo columnas no vacías y no "decorativas")
# - choice_filter_summary (qué columnas de choices usa cada filtro)
# - Normalización mínima: trims + comillas tipográficas -> ASCII
# - Validaciones suaves (warnings)
# ============================================================

#' Leer XLSForm para limpieza (sin normalizar valores/labels)
#' @param path Ruta al archivo .xlsx
#' @param lang Código de idioma a preferir para label (p.ej., "es")
#' @param prefer_label Nombre exacto de la columna de label a priorizar (opcional)
#' @return lista: survey_raw, choices_raw, settings_raw, survey, choices, meta
#' @export
leer_xlsform_limpieza <- function(path, lang = "es", prefer_label = NULL) {
  stopifnot(file.exists(path))

  suppressPackageStartupMessages({
    requireNamespace("readxl", quietly = TRUE)
    requireNamespace("dplyr",  quietly = TRUE)
    requireNamespace("stringr", quietly = TRUE)
    requireNamespace("tibble",  quietly = TRUE)
    requireNamespace("purrr",   quietly = TRUE)
  })

  `%||%` <- function(a, b) if (is.null(a)) b else a

  .trim <- function(x) {
    x <- as.character(x)
    x <- stringr::str_replace_all(x, "\\r\\n|\\r|\\n", " ")
    stringr::str_squish(x)
  }
  .fix_names <- function(df) {
    if (is.null(df) || nrow(df) == 0) return(df)
    cn <- .trim(names(df))
    dupl <- duplicated(tolower(cn))
    if (any(dupl)) {
      idx <- ave(seq_along(cn), tolower(cn), FUN = seq_along)
      cn <- ifelse(idx > 1, paste0(cn, "_", idx), cn)
    }
    names(df) <- cn
    df
  }
  .to_char_trim_df <- function(df) {
    if (is.null(df) || nrow(df) == 0) return(df)
    dplyr::mutate(df, dplyr::across(dplyr::everything(), ~ .trim(.)))
  }
  .pick_label_col <- function(cols, lang = "es", prefer_label = NULL) {
    if (!is.null(prefer_label) && prefer_label %in% cols) return(prefer_label)
    low <- tolower(cols)
    exacts <- c(
      paste0("label::", lang),
      paste0("label::", lang, " (", toupper(lang), ")"),
      paste0("label::", toupper(lang), " (", toupper(lang), ")"),
      "label"
    )
    for (ex in exacts) {
      hit <- which(low == tolower(ex))
      if (length(hit)) return(cols[hit[1]])
    }
    hit <- grep(paste0("^label.*", lang), low)
    if (length(hit)) return(cols[hit[1]])
    hit <- grep("^label", low)
    if (length(hit)) return(cols[hit[1]])
    NA_character_
  }
  .parse_type_scalar <- function(type) {
    type <- .trim(type)
    if (!nzchar(type)) return(list(base = "", list_name = ""))
    if (stringr::str_detect(type, "^begin_group"))  return(list(base = "begin_group",  list_name = ""))
    if (stringr::str_detect(type, "^end_group"))    return(list(base = "end_group",    list_name = ""))
    if (stringr::str_detect(type, "^begin_repeat")) return(list(base = "begin_repeat", list_name = ""))
    if (stringr::str_detect(type, "^end_repeat"))   return(list(base = "end_repeat",   list_name = ""))
    m <- stringr::str_match(type, "^(select_one|select_multiple)\\s+(.+)$")
    if (!any(is.na(m))) return(list(base = m[,2], list_name = .trim(m[,3])))
    list(base = type, list_name = "")
  }
  .norm_list_name <- function(x){
    x <- tolower(trimws(as.character(x)))
    x <- gsub("\\s+", "_", x)
    gsub("[^a-z0-9_]", "_", x)
  }
  .lists_with_other <- function(choices_df) {
    if (is.null(choices_df) || !nrow(choices_df)) {
      return(tibble::tibble(list_name = character(), has_other = logical()))
    }
    nms <- tolower(names(choices_df))
    lab_col <- if ("label" %in% nms) names(choices_df)[which(nms=="label")[1]] else NA_character_
    has_other_vec <- purrr::map_lgl(unique(choices_df$list_name), function(ln){
      sub <- choices_df[choices_df$list_name == ln, , drop = FALSE]
      if (!nrow(sub)) return(FALSE)
      any(.trim(sub$name) %in% c("Other","Otro","Otra","Otro(a)")) ||
        (!is.na(lab_col) && any(.trim(sub[[lab_col]]) %in% c("Other","Otro","Otra","Otro(a)","Otro (especifique)")))
    })
    tibble::tibble(list_name = unique(choices_df$list_name), has_other = has_other_vec)
  }
  .vars_in_expr <- function(expr) {
    if (is.null(expr) || is.na(expr) || !nzchar(expr)) return(character(0))
    m <- stringr::str_match_all(expr, "\\$\\{([^}]+)\\}")
    unique(unlist(lapply(m, function(mm) mm[,2])))
  }
  .normalize_quotes <- function(x){
    x <- gsub("\u2018|\u2019", "'", x, perl = TRUE)
    x <- gsub("\u201C|\u201D", "\"", x, perl = TRUE)
    x
  }
  .nonempty_cols_by_list <- function(ch){
    core_cols   <- c("list_name","list_norm","name","label")
    ignore_like <- c("^label(::|:).*", "^media(::|:).*", "^name.*_label$", "^order$", "^choice.*_name$")
    by_list <- split(ch, ch$list_name)
    res <- lapply(by_list, function(df){
      cand <- setdiff(names(df), core_cols)
      if (!length(cand)) return(character(0))
      is_ignored <- Reduce(`|`, lapply(ignore_like, function(rx) grepl(rx, cand, ignore.case = TRUE)), init = FALSE)
      cand <- cand[!is_ignored]
      if (!length(cand)) return(character(0))
      keep <- vapply(cand, function(col){
        any(nchar(trimws(as.character(df[[col]]))) > 0, na.rm = TRUE)
      }, logical(1))
      cand[keep]
    })
    tibble::tibble(list_name = names(res), extra_cols = unname(res))
  }

  # ---- leer hojas (texto, guess_max alto)
  sheets <- readxl::excel_sheets(path)
  .get_sheet <- function(nm, guess_max = 10000L) {
    i <- which(tolower(sheets) == tolower(nm))
    if (!length(i)) return(NULL)
    readxl::read_excel(path, sheet = sheets[i[1]], .name_repair = "minimal",
                       guess_max = guess_max, col_types = "text")
  }
  survey_raw   <- .get_sheet("survey")
  choices_raw  <- .get_sheet("choices")
  settings_raw <- .get_sheet("settings")

  if (is.null(survey_raw)) stop("Hoja 'survey' no encontrada en el XLSForm.")
  if (is.null(choices_raw)) {
    choices_raw <- tibble::tibble(list_name = character(), name = character(), label = character())
  }

  survey_raw   <- .fix_names(survey_raw)   |> .to_char_trim_df()
  choices_raw  <- .fix_names(choices_raw)  |> .to_char_trim_df()
  settings_raw <- .fix_names(settings_raw) |> .to_char_trim_df()

  # normaliza comillas en expresiones (mínima normalización)
  for (cc in c("relevant","constraint","choice_filter","calculation")) {
    if (cc %in% names(survey_raw)) survey_raw[[cc]] <- .normalize_quotes(survey_raw[[cc]])
  }

  # columnas mínimas garantizadas en survey
  for (cc in c("type","name","required","relevant","constraint","calculation","appearance","hint","choice_filter")) {
    if (!cc %in% names(survey_raw)) survey_raw[[cc]] <- ""
  }

  survey_label_col  <- .pick_label_col(names(survey_raw),  lang = lang, prefer_label = prefer_label)
  choices_label_col <- .pick_label_col(names(choices_raw), lang = lang, prefer_label = prefer_label)

  parsed <- purrr::map(survey_raw$type, .parse_type_scalar)

  survey <- survey_raw |>
    dplyr::mutate(
      type_base    = purrr::map_chr(parsed, "base"),
      list_name    = purrr::map_chr(parsed, "list_name"),
      q_order      = dplyr::row_number(),
      list_norm    = .norm_list_name(list_name),
      is_begin     = type_base %in% c("begin_group","begin_repeat"),
      is_end       = type_base %in% c("end_group","end_repeat")
    )

  # init
  survey$group_name  <- ""
  survey$group_label <- ""

  # ---- detectar grupos / repeats
  grp_stack <- list()
  label_by_gname <- list()
  groups_detail <- list()
  depth <- 0

  for (i in seq_len(nrow(survey))) {
    tb <- survey$type_base[i]
    nm <- survey$name[i]

    # label del grupo (prefiere prefer_label, luego "label")
    lb <- NA_character_
    if (!is.null(survey_label_col) && !is.na(survey_label_col) && survey_label_col %in% names(survey_raw)) {
      lb <- survey_raw[[survey_label_col]][i]
    }
    if ((is.na(lb) || !nzchar(lb)) && "label" %in% names(survey_raw)) {
      lb <- survey_raw[["label"]][i]
    }

    if (tb %in% c("begin_group","begin_repeat")) {
      depth <- depth + 1L
      gname <- if (nzchar(nm)) nm else paste0("group_", i)
      glabel <- if (isTRUE(nzchar(lb))) lb else gname
      grp_stack[[length(grp_stack)+1]] <- gname
      label_by_gname[[gname]] <- glabel

      groups_detail[[length(groups_detail)+1]] <- tibble::tibble(
        gname       = gname,
        glabel      = glabel,
        begin_row   = i,
        end_row     = NA_integer_,
        depth       = depth,
        is_repeat   = identical(tb, "begin_repeat"),
        relevant    = survey$relevant[i] %||% "",
        appearance  = survey$appearance[i] %||% "",
        relevant_vars = list(.vars_in_expr(survey$relevant[i] %||% ""))
      )
    } else if (tb %in% c("end_group","end_repeat")) {
      if (length(groups_detail)) {
        idx_open <- which(vapply(groups_detail, function(x) is.na(x$end_row)[1], logical(1)))
        if (length(idx_open)) {
          j <- idx_open[length(idx_open)]
          groups_detail[[j]]$end_row <- i
        }
      }
      if (length(grp_stack)) grp_stack <- grp_stack[-length(grp_stack)]
      depth <- max(0L, depth - 1L)
    }

    survey$group_name[i]  <- if (length(grp_stack)) grp_stack[[length(grp_stack)]] else ""
    survey$group_label[i] <- if (nzchar(survey$group_name[i]) && survey$group_name[i] %in% names(label_by_gname)) {
      label_by_gname[[ survey$group_name[i] ]]
    } else survey$group_name[i]
  }
  if (length(groups_detail)) {
    open_idx <- which(vapply(groups_detail, function(x) is.na(x$end_row)[1], logical(1)))
    if (length(open_idx)) {
      for (k in open_idx) groups_detail[[k]]$end_row <- nrow(survey)
    }
  }

  groups_detail_df <- if (length(groups_detail)) dplyr::bind_rows(groups_detail) else
    tibble::tibble(gname=character(), glabel=character(), begin_row=integer(),
                   end_row=integer(), depth=integer(), is_repeat=logical(),
                   relevant=character(), appearance=character(), relevant_vars=list())

  # ---- survey preguntas (excluye begin/end) + vars referenciadas
  if (!"choice_filter" %in% names(survey)) survey$choice_filter <- ""
  survey_questions <- survey |>
    dplyr::filter(!(type_base %in% c("begin_group","end_group","begin_repeat","end_repeat"))) |>
    dplyr::mutate(
      vars_in_relevant     = lapply(relevant,     .vars_in_expr),
      vars_in_constraint   = lapply(constraint,   .vars_in_expr),
      vars_in_calc         = lapply(calculation,  .vars_in_expr),
      vars_in_choicefilter = lapply(choice_filter,.vars_in_expr)
    ) |>
    dplyr::select(
      type, type_base, list_name, list_norm,
      name, dplyr::any_of(c("label")),
      required, relevant, constraint, calculation,
      choice_filter,
      appearance, hint,
      group_name, group_label,
      q_order,
      vars_in_relevant, vars_in_constraint, vars_in_calc, vars_in_choicefilter,
      dplyr::everything()
    )

  # ---- choices (mínimo + list_norm)
  for (cc in c("list_name","name")) if (!cc %in% names(choices_raw)) choices_raw[[cc]] <- ""
  choices <- choices_raw |>
    dplyr::mutate(
      list_name = .data$list_name,
      name      = .data$name
    ) |>
    dplyr::mutate(list_norm = .norm_list_name(list_name)) |>
    dplyr::select(list_name, list_norm, name, dplyr::any_of("label"), dplyr::everything())

  # ---- validaciones suaves
  dup_ch <- choices |> dplyr::count(list_name, name) |> dplyr::filter(n > 1)
  if (nrow(dup_ch)) warning(sprintf(
    "Duplicados en choices por (list_name, name): %d filas. Ej.: %s",
    nrow(dup_ch), paste(utils::head(paste0(dup_ch$list_name, ":", dup_ch$name), 5), collapse=", ")
  ))
  vac_svy <- survey_questions |> dplyr::filter(!nzchar(name))
  if (nrow(vac_svy)) warning(sprintf("Preguntas en survey con 'name' vacío: %d", nrow(vac_svy)))
  vac_ch  <- choices |> dplyr::filter(!nzchar(list_name) | !nzchar(name))
  if (nrow(vac_ch)) warning(sprintf("Filas en choices con 'list_name' o 'name' vacío: %d", nrow(vac_ch)))
  ln_needed <- survey_questions |> dplyr::filter(type_base %in% c("select_one","select_multiple")) |> dplyr::pull(list_name) |> unique()
  ln_have   <- unique(choices$list_name)
  miss_ln   <- setdiff(ln_needed, ln_have)
  if (length(miss_ln)) warning(sprintf("Listas referidas en survey sin definir en choices: %s", paste(miss_ln, collapse=", ")))
  col_norm <- aggregate(list_name ~ list_norm, choices, function(v) length(unique(v)))
  col_norm <- subset(col_norm, list_name > 1)
  if (nrow(col_norm)) warning("Colisiones de list_norm: distintos list_name mapean al mismo list_norm (revisa acentos/espacios).")

  lists_with_other <- .lists_with_other(choices)

  # ---- meta: columnas no-vacías por lista (potenciales filtros)
  choice_cols_by_list <- .nonempty_cols_by_list(choices)

  # ---- section_map (prefijo, condicionalidad e is_repeat) dentro de la función
  if (nrow(groups_detail_df)) {
    base_map <- tibble::tibble(
      group_name     = as.character(groups_detail_df$gname),
      group_label    = as.character(groups_detail_df$glabel),
      group_relevant = {
        r <- groups_detail_df$relevant; r <- ifelse(is.na(r), NA_character_, as.character(r))
        r <- gsub("\\s+", " ", trimws(r))
        ifelse(nzchar(r), r, NA_character_)
      },
      is_conditional = {
        r <- groups_detail_df$relevant
        nzchar(gsub("\\s+", " ", ifelse(is.na(r), "", as.character(r))))
      },
      is_repeat      = as.logical(groups_detail_df$is_repeat),
      .gord          = rank(groups_detail_df$begin_row, ties.method = "first") |> as.integer()
    )
    # prefijos heurísticos (respetar previos si hubiera)
    pref <- rep(NA_character_, nrow(base_map))
    need_pref <- is.na(pref) | !nzchar(pref)
    if (any(need_pref)) {
      gn  <- base_map$group_name[need_pref]
      src <- base_map$group_label[need_pref]
      idx_empty <- is.na(src) | !nzchar(src)
      if (any(idx_empty)) src[idx_empty] <- gn[idx_empty]
      heur <- toupper(gsub("[^A-Za-z]", "", substr(src, 1, 3)))
      heur[!nzchar(heur)] <- "GEN"
      pref[need_pref] <- paste0(heur, "_")
    }
    base_map$prefix <- pref
    base_map <- dplyr::distinct(base_map, .data$group_name, .keep_all = TRUE)
    section_map <- base_map
  } else {
    section_map <- tibble::tibble(
      group_name     = character(),
      group_label    = character(),
      prefix         = character(),
      is_conditional = logical(),
      is_repeat      = logical(),
      group_relevant = character(),
      .gord          = integer()
    )
  }

  # ---- resumen de choice_filter (qué columnas de choices usa)
  has_label <- "label" %in% names(survey_questions)
  cols_for_list <- function(ln){
    row <- choice_cols_by_list[choice_cols_by_list$list_name == ln, , drop = FALSE]
    if (!nrow(row)) character(0) else unlist(row$extra_cols)
  }
  tokens_in_cf <- function(cf){
    if (is.na(cf) || !nzchar(cf)) return(character(0))
    toks <- unlist(strsplit(cf, "[^A-Za-z0-9_]+"))
    toks[nzchar(toks)]
  }
  csum0 <- survey_questions |>
    dplyr::filter(nzchar(choice_filter)) |>
    dplyr::rowwise() |>
    dplyr::mutate(
      cf_tokens    = list(tokens_in_cf(choice_filter)),
      cf_uses_cols = list(intersect(cf_tokens, cols_for_list(list_name))),
      cf_vars      = list(vars_in_choicefilter)
    ) |>
    dplyr::ungroup()
  if (has_label) {
    choice_filter_summary <- csum0 |>
      dplyr::transmute(
        q_order, name, label,
        list_name, choice_filter,
        cf_vars      = vapply(cf_vars,      function(v) paste(v, collapse=", "), character(1)),
        cf_uses_cols = vapply(cf_uses_cols, function(v) paste(v, collapse=", "), character(1))
      )
  } else {
    choice_filter_summary <- csum0 |>
      dplyr::transmute(
        q_order, name,
        label = NA_character_,
        list_name, choice_filter,
        cf_vars      = vapply(cf_vars,      function(v) paste(v, collapse=", "), character(1)),
        cf_uses_cols = vapply(cf_uses_cols, function(v) paste(v, collapse=", "), character(1))
      )
  }

  # ---- salida
  meta <- list(
    label_col_survey   = survey_label_col,
    label_col_choices  = choices_label_col,
    groups_detail      = groups_detail_df,
    lists_with_other   = lists_with_other,
    choice_cols_by_list= choice_cols_by_list,
    section_map        = section_map,
    choice_filter_summary = choice_filter_summary
  )

  out <- list(
    survey_raw   = survey_raw,
    choices_raw  = choices_raw,
    settings_raw = settings_raw,
    survey       = tibble::as_tibble(survey_questions),
    choices      = tibble::as_tibble(choices),
    meta         = meta
  )

  # advertencia opcional de balance begin/end
  b <- sum(grepl("^begin_", survey$type_base))
  e <- sum(grepl("^end_",   survey$type_base))
  if (b != e) warning(sprintf("Desbalance begin/end: begin=%d, end=%d (revisa la hoja 'survey')", b, e))

  out
}
