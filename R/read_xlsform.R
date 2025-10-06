
# ============================================================
# LECTOR "SIN NORMALIZAR" PARA REGLAS / PLAN DE LIMPIEZA
# Lee un XLSForm (KoBo/ODK) y agrega solo auxiliares mínimos.
# Incluye detección robusta de grupos y sus condiciones de apertura.
# ============================================================

#' Leer XLSForm para limpieza (sin normalizar valores/labels)
#'
#' Lee las hojas \code{survey}, \code{choices} y \code{settings} de un XLSForm
#' y agrega únicamente auxiliares mínimos para fabricar reglas de limpieza, sin
#' renombrar columnas visibles ni reescribir labels. Detecta grupos
#' (\code{begin_group}/\code{end_group}), su anidación, \code{appearance}
#' (p.ej., \code{field-list}), y la condición \code{relevant} que los abre
#' (p.ej. \verb{${Consent} = 'Yes'} o expresiones compuestas con \code{selected()}).
#'
#' @param path Ruta al archivo XLSForm (.xlsx).
#' @param lang Idioma preferido para detectar columnas de label (p.ej., "es").
#'   Sólo se usa para señalar cuál columna existe; no se reescriben labels.
#' @param prefer_label Nombre exacto de la columna de label a priorizar (opcional).
#'   Si existe, se registra como la elegida en meta, pero no se modifica el contenido.
#'
#' @return Lista con:
#' \itemize{
#'   \item \code{survey_raw}, \code{choices_raw}, \code{settings_raw}: hojas crudas (sólo saneo de encabezados).
#'   \item \code{survey}: preguntas (sin begin/end), con auxiliares:
#'         \code{type_base}, \code{list_name}, \code{list_norm}, \code{q_order},
#'         \code{group_name}, \code{group_label}, \code{relevant}, \code{constraint},
#'         \code{calculation}, \code{appearance}.
#'   \item \code{choices}: opciones (list_name, name, label si existe, list_norm).
#'   \item \code{meta}: lista con
#'         \code{label_col_survey}, \code{label_col_choices},
#'         \code{groups_detail} (rango begin/end, nivel, relevant, vars referenciadas),
#'         \code{lists_with_other}, \code{section_map}.
#' }
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
    rx_lang <- paste0("^(label(_|\\:|::)?", lang, ")$|^(label.*", lang, ".*)$")
    cand_lang <- cols[stringr::str_detect(tolower(cols), rx_lang)]
    if (length(cand_lang) > 0) return(cand_lang[order(nchar(cand_lang))][1])
    if ("label" %in% tolower(cols)) return(cols[tolower(cols) == "label"][1])
    cand <- cols[stringr::str_detect(tolower(cols), "^label")]
    if (length(cand) > 0) return(cand[1])
    NA_character_
  }
  .parse_type_scalar <- function(type) {
    type <- .trim(type)
    if (!nzchar(type)) return(list(base = "", list_name = ""))
    if (stringr::str_detect(type, "^begin_group")) return(list(base = "begin_group", list_name = ""))
    if (stringr::str_detect(type, "^end_group"))   return(list(base = "end_group",   list_name = ""))
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
      any(.trim(sub$name) %in% c("Other","Otro","Otra")) ||
        (!is.na(lab_col) && any(.trim(sub[[lab_col]]) %in% c("Other","Otro","Otra")))
    })
    tibble::tibble(list_name = unique(choices_df$list_name), has_other = has_other_vec)
  }
  .vars_in_expr <- function(expr) {
    if (is.null(expr) || is.na(expr) || !nzchar(expr)) return(character(0))
    m <- stringr::str_match_all(expr, "\\$\\{([^}]+)\\}")
    unique(unlist(lapply(m, function(mm) mm[,2])))
  }

  # -------- leer hojas
  sheets <- readxl::excel_sheets(path)
  .get_sheet <- function(nm) {
    i <- which(tolower(sheets) == tolower(nm))
    if (!length(i)) return(NULL)
    readxl::read_excel(path, sheet = sheets[i[1]], .name_repair = "minimal")
  }
  survey_raw   <- .get_sheet("survey")
  choices_raw  <- .get_sheet("choices")
  settings_raw <- .get_sheet("settings")

  if (is.null(survey_raw)) {
    stop("Hoja 'survey' no encontrada en el XLSForm.")
  }
  if (is.null(choices_raw)) {
    choices_raw <- tibble::tibble(list_name = character(), name = character(), label = character())
  }

  survey_raw   <- .fix_names(survey_raw)   |> .to_char_trim_df()
  choices_raw  <- .fix_names(choices_raw)  |> .to_char_trim_df()
  settings_raw <- .fix_names(settings_raw) |> .to_char_trim_df()

  for (cc in c("type","name","required","relevant","constraint","calculation","appearance","hint")) {
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
      is_begin     = type_base == "begin_group",
      is_end       = type_base == "end_group"
    )

  # Inicializar para evitar warnings
  survey$group_name  <- ""
  survey$group_label <- ""

  # -------- detectar grupos / anidación
  grp_stack <- list()
  label_by_gname <- list()
  groups_detail <- list()
  depth <- 0

  for (i in seq_len(nrow(survey))) {
    tb <- survey$type_base[i]
    nm <- survey$name[i]
    lb <- if (!is.null(survey_label_col) && !is.na(survey_label_col) && survey_label_col %in% names(survey_raw)) {
      survey_raw[[survey_label_col]][i] %||% ""
    } else survey_raw[["label"]][i] %||% ""

    if (identical(tb, "begin_group")) {
      depth <- depth + 1L
      gname <- if (nzchar(nm)) nm else paste0("group_", i)
      glabel <- if (nzchar(lb)) lb else gname
      grp_stack[[length(grp_stack)+1]] <- gname
      label_by_gname[[gname]] <- glabel

      groups_detail[[length(groups_detail)+1]] <- tibble::tibble(
        gname       = gname,
        glabel      = glabel,
        begin_row   = i,
        end_row     = NA_integer_,
        depth       = depth,
        relevant    = survey$relevant[i] %||% "",
        appearance  = survey$appearance[i] %||% "",
        relevant_vars = list(.vars_in_expr(survey$relevant[i] %||% ""))
      )
    } else if (identical(tb, "end_group")) {
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
                   end_row=integer(), depth=integer(), relevant=character(),
                   appearance=character(), relevant_vars=list())

  survey_questions <- survey |>
    dplyr::filter(!(type_base %in% c("begin_group","end_group"))) |>
    dplyr::select(
      type, type_base, list_name, list_norm,
      name,
      dplyr::any_of(c("label")),
      required, relevant, constraint, calculation,
      appearance, hint,
      group_name, group_label,
      q_order,
      dplyr::everything()
    )

  # -------- choices
  for (cc in c("list_name","name")) if (!cc %in% names(choices_raw)) choices_raw[[cc]] <- ""
  choices <- choices_raw |>
    dplyr::mutate(
      list_name = .data$list_name,
      name      = .data$name
    ) |>
    dplyr::mutate(list_norm = .norm_list_name(list_name)) |>
    dplyr::select(list_name, list_norm, name, dplyr::any_of("label"), dplyr::everything())

  lists_with_other <- .lists_with_other(choices)

  # Advertencia: preguntas fuera de grupo
  no_group <- survey_questions |> dplyr::filter(!nzchar(group_name))
  if (nrow(no_group) > 0) {
    warning(sprintf(
      "Hay %d preguntas fuera de cualquier begin_group/end_group (group_name vacío). Ejemplos: %s",
      nrow(no_group), paste(utils::head(no_group$name, 5), collapse = ", ")
    ))
  }

  # -------- meta básica (section_map se enriquece después)
  meta <- list(
    label_col_survey   = survey_label_col,
    label_col_choices  = choices_label_col,
    groups_detail      = groups_detail_df,
    lists_with_other   = lists_with_other,
    section_map        = tibble::tibble()   # se llenará con enriquecer_section_map()
  )

  # -------- salida intermedia
  out <- list(
    survey_raw   = survey_raw,
    choices_raw  = choices_raw,
    settings_raw = settings_raw,
    survey       = tibble::as_tibble(survey_questions),
    choices      = tibble::as_tibble(choices),
    meta         = meta
  )

  # -------- enriquecer section_map (G, relevant por grupo, prefix, .gord)
  out <- enriquecer_section_map(out)
  return(out)
}


#------G condiciones -------------

#' Enriquecer el section_map con banderas de condicionalidad (G) usando groups_detail
#' @export
enriquecer_section_map <- function(x) {
  stopifnot(is.list(x), "meta" %in% names(x))
  gd <- x$meta$groups_detail
  # Si por algún motivo no existe, deja un mapa mínimo
  if (is.null(gd) || !is.data.frame(gd) || nrow(gd) == 0) {
    x$meta$section_map <- tibble::tibble(
      group_name     = character(),
      group_label    = character(),
      prefix         = character(),
      is_conditional = logical(),
      group_relevant = character(),
      .gord          = integer()
    )
    return(x)
  }

  # Columnas esperadas en groups_detail: gname, glabel, relevant, begin_row (para ordenar)
  need <- c("gname","glabel","relevant","begin_row")
  if (!all(need %in% names(gd))) {
    stop("meta$groups_detail no tiene las columnas requeridas: ",
         paste(setdiff(need, names(gd)), collapse = ", "))
  }

  base_map <- tibble::tibble(
    group_name     = as.character(gd$gname),
    group_label    = as.character(gd$glabel),
    group_relevant = {
      r <- gd$relevant
      r <- ifelse(is.na(r), NA_character_, as.character(r))
      r <- gsub("\\s+", " ", trimws(r))
      ifelse(nzchar(r), r, NA_character_)
    },
    is_conditional = {
      r <- gd$relevant
      nzchar(gsub("\\s+", " ", ifelse(is.na(r), "", as.character(r))))
    },
    .gord          = rank(gd$begin_row, ties.method = "first") |> as.integer()
  )

  # Respetar prefijos previos si existen
  prior <- x$meta$section_map
  if (is.data.frame(prior) && nrow(prior) && all(c("group_name","prefix") %in% names(prior))) {
    pref <- prior$prefix[match(base_map$group_name, prior$group_name)]
  } else {
    # ANTES: pref <- NA_character_
    pref <- rep(NA_character_, nrow(base_map))  # <-- FIX: vector del tamaño correcto
  }

  # Asignar prefijo heurístico donde falte
  need_pref <- is.na(pref) | !nzchar(pref)

  if (any(need_pref)) {
    gn  <- base_map$group_name[need_pref]
    src <- base_map$group_label[need_pref]

    # tratar NAs y vacíos de forma segura
    idx_empty <- is.na(src) | !nzchar(src)
    if (any(idx_empty)) src[idx_empty] <- gn[idx_empty]

    heur <- toupper(gsub("[^A-Za-z]", "", substr(src, 1, 3)))
    heur[!nzchar(heur)] <- "GEN"
    pref[need_pref] <- paste0(heur, "_")
  }

  base_map$prefix <- pref

  # Sin duplicados por group_name
  base_map <- dplyr::distinct(base_map, .data$group_name, .keep_all = TRUE)

  x$meta$section_map <- base_map
  x
}
