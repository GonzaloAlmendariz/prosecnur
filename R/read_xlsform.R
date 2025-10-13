# ============================================================
# LECTOR "SIN NORMALIZAR" PARA REGLAS / PLAN DE LIMPIEZA — ONE-FILE v2.3
# - Soporta begin_group/end_group y begin_repeat/end_repeat
# - Captura choice_filter y variables referenciadas
# - groups_detail con glabel (label del begin_group; cae en gname si falta)
# - section_map dentro de meta (prefix, is_conditional, is_repeat)
# - choice_cols_by_list depurado (solo columnas no vacías y no "decorativas")
# - choice_filter_summary (qué columnas de choices usa cada filtro)
# - Normalización mínima: trims + comillas tipográficas -> ASCII
# - Validaciones suaves (warnings)
# ============================================================

# Helpers:

# Simplifica texto: quita acentos y deja solo [A-Za-z0-9_], espacios -> _
.slugify_es <- function(x) {
  x <- as.character(x)
  x <- iconv(x, to = "ASCII//TRANSLIT")        # quita acentos
  x <- gsub("[^A-Za-z0-9]+", "_", x)           # no alfanumérico -> _
  x <- gsub("_+", "_", x)                      # colapsa
  x <- gsub("^_|_$", "", x)                    # recorta _ extremos
  x
}

# Abrevia en modo "heurístico" (3-4 letras): toma la primera palabra significativa
.abreviar_heuristico <- function(lbl, max_len = 4) {
  if (is.na(lbl) || !nzchar(lbl)) return("GEN")
  # Limpia prefijos comunes tipo "S1:", "Parte 2.", etc.
  txt <- gsub("^\\s*(S\\d+\\s*[:.-]|Parte\\s*\\d+\\s*[:.-])\\s*", "", lbl, ignore.case = TRUE)
  # Toma palabras y omite stopwords simples
  toks <- unlist(strsplit(txt, "\\s+"))
  stopw <- c("de","del","la","el","los","las","y","en","para","por","a","al","lo","un","una","uno")
  toks <- toks[!tolower(toks) %in% stopw]
  base <- if (length(toks)) toks[1] else txt
  base <- gsub("[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ]", "", base)
  if (!nzchar(base)) base <- "GEN"
  # Normaliza a ASCII y toma primeras letras
  base <- .slugify_es(base)
  base <- toupper(substr(base, 1, max(1, max_len)))
  base
}

# Aplica transformación solicitada al prefijo base
.transformar_según_modo <- function(x, modo = c("mayúsculas","tal_cual","simplificar")) {
  modo <- match.arg(modo)
  if (modo == "mayúsculas") return(toupper(x))
  if (modo == "simplificar") return(toupper(.slugify_es(x)))
  x
}




#' Leer XLSForm para limpieza (sin normalizar valores/labels)
#'
#' Lee un XLSForm (hojas `survey`, `choices`, `settings`) y devuelve
#' objetos listos para generar el Plan de Limpieza (ACNUR), sin alterar
#' los valores originales. Detecta grupos y repeticiones, choice filters,
#' listas dinámicas (select_* ${var}), y arma un mapa de secciones con
#' prefijos configurables.
#'
#' @param path Ruta al archivo `.xlsx`.
#' @param lang Código de idioma a preferir para la etiqueta (p. ej., `"es"`).
#' @param prefer_label Nombre exacto de la columna de label a priorizar (opcional),
#'   p. ej. `"label::Spanish (ES)"`. Si no se encuentra, se intentan variantes
#'   comunes y finalmente `label`.
#' @param origen_prefijo Estrategia para generar el prefijo de sección en `meta$section_map$prefix`.
#'   Valores: \itemize{
#'     \item `"deducir"`: (por defecto) deduce un prefijo corto a partir de la etiqueta
#'       del grupo (o del nombre si no hay etiqueta). Útil cuando se buscan siglas
#'       cortas tipo `DOC_`, `MOV_`, `PAR_`.
#'     \item `"nombre_grupo"`: usa exactamente el `name` del `begin_group` / `begin_repeat`
#'       como base del prefijo (p. ej., `group_consent_`).
#'     \item `"etiqueta_grupo"`: usa el `label` del grupo como base del prefijo
#'       (p. ej., `CONSENTIMIENTO_`).
#'   }
#' @param transformar_prefijo Transformación a aplicar al texto base del prefijo.
#'   Valores: \itemize{
#'     \item `"mayúsculas"` (por defecto): convierte a mayúsculas.
#'     \item `"tal_cual"`: no modifica el texto.
#'     \item `"simplificar"`: quita acentos y deja solo \code{[A-Za-z0-9_]}.
#'   }
#' @param sufijo_prefijo Cadena a añadir al final del prefijo (por defecto, `"_”`).
#' @param max_longitud_prefijo Longitud máxima del prefijo \strong{solo cuando}
#'   \code{origen_prefijo = "deducir"} (por defecto, 4). No se aplica en
#'   `"nombre_grupo"` ni `"etiqueta_grupo"`.
#' @param asegurar_unicidad Si \code{TRUE} (por defecto), garantiza prefijos únicos
#'   añadiendo sufijos numéricos estables en caso de colisiones.
#'
#' @return Una \strong{lista} con elementos:
#'   \itemize{
#'     \item \code{survey_raw}: hoja \code{survey} tal cual (nombres saneados).
#'     \item \code{choices_raw}: hoja \code{choices} tal cual (nombres saneados).
#'     \item \code{settings_raw}: hoja \code{settings} (si existe).
#'     \item \code{survey}: preguntas (excluye begin/end) con metadatos útiles:
#'       \itemize{
#'         \item \code{type_base}, \code{list_name}, \code{list_norm}
#'         \item \code{group_name}, \code{group_label}
#'         \item \code{required}, \code{relevant}, \code{constraint}, \code{calculation}
#'         \item \code{choice_filter}
#'         \item \code{dyn_ref}: referencia de \emph{lista dinámica} si el tipo es \code{select_* ${var}}
#'         \item \code{vars_in_*}: variables referenciadas en \code{relevant/constraint/choice_filter/calculation}
#'       }
#'     \item \code{choices}: catálogo con \code{list_name}, \code{list_norm}, \code{name}, \code{label} (columna elegida).
#'     \item \code{meta}: lista con resúmenes y mapas:
#'       \itemize{
#'         \item \code{label_col_survey}, \code{label_col_choices}: columnas de label efectivamente usadas.
#'         \item \code{groups_detail}: detalle de grupos/repeats (incluye \code{is_repeat}, \code{relevant} y, si existe, \code{repeat_count}).
#'         \item \code{section_map}: mapa de secciones con \code{prefix}, \code{is_conditional}, \code{is_repeat}, \code{group_relevant}.
#'         \item \code{lists_with_other}: listas con opción “Otro/Other”.
#'         \item \code{choice_cols_by_list}: columnas no vacías por lista (útiles para \code{choice_filter}).
#'         \item \code{choice_filter_summary}: por pregunta con filtro, qué columnas del catálogo y variables del formulario intervienen.
#'       }
#'   }
#'
#' @details
#' \itemize{
#'   \item Los nombres de columnas vacíos/NA se renombran como \code{generico_#};
#'         las columnas completamente vacías y sin nombre se eliminan con aviso.
#'   \item \code{origen_prefijo = "deducir"} usa una abreviación corta (longitud controlada
#'         por \code{max_longitud_prefijo}). En los otros modos no se recorta.
#'   \item Las preguntas \emph{select} con \emph{lista dinámica} (\code{select_* ${var}})
#'         no exigen \code{choices$list_name} literal; por eso no generan falsos
#'         avisos de “lista faltante”.
#' }
#'
#' @examples
#' \dontrun{
#' # Prefijos deducidos (comportamiento por defecto)
#' inst <- leer_xlsform_limpieza(
#'   path = "RMS_instrumento.xlsx",
#'   lang = "es",
#'   prefer_label = "label::Español (es)"
#' )
#'
#' # Prefijo igual al nombre del grupo, sin modificar
#' inst2 <- leer_xlsform_limpieza(
#'   path = "RMS_instrumento.xlsx",
#'   origen_prefijo = "nombre_grupo",
#'   transformar_prefijo = "tal_cual",
#'   sufijo_prefijo = "_"
#' )
#'
#' # Prefijo desde la etiqueta del grupo, simplificado (sin acentos, A-Z0-9_)
#' inst3 <- leer_xlsform_limpieza(
#'   path = "RMS_instrumento.xlsx",
#'   origen_prefijo = "etiqueta_grupo",
#'   transformar_prefijo = "simplificar"
#' )
#' }
#'
#' @export
leer_xlsform_limpieza <- function(path,
                                  lang = "es",
                                  prefer_label = NULL,
                                  origen_prefijo = c("deducir", "nombre_grupo", "etiqueta_grupo"),
                                  transformar_prefijo = c("mayúsculas", "tal_cual", "simplificar"),
                                  sufijo_prefijo = "_",
                                  max_longitud_prefijo = 4,
                                  asegurar_unicidad = TRUE) {
  origen_prefijo      <- match.arg(origen_prefijo)
  transformar_prefijo <- match.arg(transformar_prefijo)
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

    cn <- names(df)
    cn[is.na(cn)] <- ""
    cn <- .trim(cn)

    # Detecta columnas totalmente vacías (todas las celdas NA o "")
    is_all_empty <- function(v) {
      vv <- as.character(v)
      all(is.na(vv) | !nzchar(.trim(vv)))
    }

    drop_cols <- which(!nzchar(cn) & vapply(df, is_all_empty, logical(1)))
    if (length(drop_cols)) {
      df <- df[, -drop_cols, drop = FALSE]
      cn <- names(df)
      cn[is.na(cn)] <- ""
      cn <- .trim(cn)
      message(sprintf("Aviso: eliminadas %d columnas sin nombre y vacías (survey/choices).", length(drop_cols)))
    }

    # Asigna nombres a encabezados vacíos restantes
    if (any(!nzchar(cn))) {
      idx_empty <- which(!nzchar(cn))
      cn[idx_empty] <- paste0("generico_", seq_along(idx_empty))
    }

    # Resuelve duplicados (case-insensitive) sin perder grafía original
    low <- tolower(cn)
    dupl <- duplicated(low)
    if (any(dupl)) {
      cn <- make.unique(cn, sep = "_")
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
      "label::spanish (es)", "label::spanish(es)", "label::spanish_es",
      "label_spanish_es", "label::spanish", "label::es",
      "label::español (es)", "label::espanol (es)",
      "label::español", "label::espanol",
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
    if (!nzchar(type)) return(list(base = "", list_name = "", dyn_ref = ""))
    if (stringr::str_detect(type, "^begin_group"))  return(list(base = "begin_group",  list_name = "", dyn_ref = ""))
    if (stringr::str_detect(type, "^end_group"))    return(list(base = "end_group",    list_name = "", dyn_ref = ""))
    if (stringr::str_detect(type, "^begin_repeat")) return(list(base = "begin_repeat", list_name = "", dyn_ref = ""))
    if (stringr::str_detect(type, "^end_repeat"))   return(list(base = "end_repeat",   list_name = "", dyn_ref = ""))
    m1 <- stringr::str_match(type, "^(select_one|select_multiple)\\s+([^\\s]+)$")
    if (!any(is.na(m1))) {
      base <- m1[,2]; rhs <- .trim(m1[,3])
      m2 <- stringr::str_match(rhs, "^\\$\\{([^}]+)\\}$")
      if (!any(is.na(m2))) return(list(base = base, list_name = "", dyn_ref = .trim(m2[,2])))
      return(list(base = base, list_name = rhs, dyn_ref = ""))
    }
    list(base = type, list_name = "", dyn_ref = "")
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
      type_base = purrr::map_chr(parsed, "base"),
      list_name = purrr::map_chr(parsed, "list_name"),
      dyn_ref   = purrr::map_chr(parsed, "dyn_ref"),
      q_order   = dplyr::row_number(),
      list_norm = .norm_list_name(list_name),
      is_begin  = type_base %in% c("begin_group","begin_repeat"),
      is_end    = type_base %in% c("end_group","end_repeat")
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

  repeat_count_vec <- if ("repeat_count" %in% names(survey_raw)) survey_raw$repeat_count else rep(NA_character_, nrow(survey_raw))

  if (nrow(groups_detail_df)) {
    groups_detail_df$repeat_count <- NA_character_
    groups_detail_df$repeat_count_vars <- vector("list", nrow(groups_detail_df))
    for (i in seq_len(nrow(groups_detail_df))) {
      if (isTRUE(groups_detail_df$is_repeat[i])) {
        br <- groups_detail_df$begin_row[i]
        rc <- repeat_count_vec[br]
        groups_detail_df$repeat_count[i] <- ifelse(is.na(rc) || !nzchar(rc), NA_character_, rc)
        groups_detail_df$repeat_count_vars[[i]] <- .vars_in_expr(groups_detail_df$repeat_count[i] %||% "")
      } else {
        groups_detail_df$repeat_count_vars[[i]] <- character(0)
      }
    }
  }

  survey_questions <- survey |>
    dplyr::filter(!(type_base %in% c("begin_group","end_group","begin_repeat","end_repeat"))) |>
    dplyr::mutate(
      vars_in_relevant     = lapply(relevant,     .vars_in_expr),
      vars_in_constraint   = lapply(constraint,   .vars_in_expr),
      vars_in_calc         = lapply(calculation,  .vars_in_expr),
      vars_in_choicefilter = lapply(choice_filter,.vars_in_expr)
    ) |>
    dplyr::select(
      type, type_base, list_name, list_norm, dyn_ref,
      name, dplyr::any_of(c("label")),
      required, relevant, constraint, calculation,
      choice_filter, appearance, hint,
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
  ln_needed <- survey_questions |>
    dplyr::filter(
      type_base %in% c("select_one","select_multiple"),
      (is.na(dyn_ref) | !nzchar(dyn_ref)),              # no dinámicas
      nzchar(list_name)                                  # con nombre real
    ) |>
    dplyr::pull(list_name) |> unique()

  ln_have <- unique(choices$list_name)
  miss_ln <- setdiff(ln_needed, ln_have)
  if (length(miss_ln)) {
    warning(sprintf("Listas referidas en survey sin definir en choices: %s",
                    paste(miss_ln, collapse = ", ")))
  }
  miss_ln   <- setdiff(ln_needed, ln_have)
  if (length(miss_ln)) warning(sprintf("Listas referidas en survey sin definir en choices: %s", paste(miss_ln, collapse=", ")))
  col_norm <- aggregate(list_name ~ list_norm, choices, function(v) length(unique(v)))
  col_norm <- subset(col_norm, list_name > 1)
  if (nrow(col_norm)) warning("Colisiones de list_norm: distintos list_name mapean al mismo list_norm (revisa acentos/espacios).")

  lists_with_other <- .lists_with_other(choices)

  # ---- meta: columnas no-vacías por lista (potenciales filtros)
  choice_cols_by_list <- .nonempty_cols_by_list(choices)

  # ---- section_map (prefijo configurable, condicionalidad e is_repeat)
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

    # Generación del prefijo según parámetros
    pref <- character(nrow(base_map))

    if (origen_prefijo == "deducir") {
      # Usa etiqueta si existe; si no, cae a nombre de grupo
      fuente <- ifelse(is.na(base_map$group_label) | !nzchar(base_map$group_label),
                       base_map$group_name, base_map$group_label)
      pref <- vapply(fuente, .abreviar_heuristico, character(1), max_len = max_longitud_prefijo)
      pref <- .transformar_según_modo(pref, transformar_prefijo)
    }

    if (origen_prefijo == "nombre_grupo") {
      pref <- base_map$group_name
      pref <- .transformar_según_modo(pref, transformar_prefijo)
      # En "nombre_grupo" NO aplicamos recorte por longitud (se mantiene tal cual)
    }

    if (origen_prefijo == "etiqueta_grupo") {
      # Si no hay etiqueta, cae a nombre de grupo
      fuente <- ifelse(is.na(base_map$group_label) | !nzchar(base_map$group_label),
                       base_map$group_name, base_map$group_label)
      pref <- .transformar_según_modo(fuente, transformar_prefijo)
      # En "etiqueta_grupo" NO recortamos por longitud (mantén lectura amigable)
    }

    # Asegurar que el prefijo no quede vacío
    pref[!nzchar(pref)] <- "GEN"

    # Añadir sufijo (p. ej., "_")
    pref <- paste0(pref, sufijo_prefijo %||% "")

    # Asegurar unicidad si se solicita
    if (isTRUE(asegurar_unicidad)) {
      # make.unique añade sufijos .1, .2; los cambiamos a _2, _3 manteniendo sufijo_prefijo final
      tmp <- make.unique(pref, sep = "")
      if (!identical(tmp, pref)) {
        # Detecta los que se alteraron y formatea sufijos como _2, _3...
        pattern <- paste0("^(", gsub("([\\W])", "\\\\\\1", pref), ")(\\.)(\\d+)$")
        for (i in seq_along(tmp)) {
          if (grepl("\\.\\d+$", tmp[i])) {
            # extrae base sin .n
            base <- sub("\\.\\d+$", "", tmp[i])
            num  <- sub("^.*\\.(\\d+)$", "\\1", tmp[i])
            # Si ya termina con sufijo_prefijo, añade número; si no, respeta el sufijo existente
            if (endsWith(base, sufijo_prefijo)) {
              tmp[i] <- paste0(base, num)
            } else {
              tmp[i] <- paste0(base, sufijo_prefijo, num)
            }
          }
        }
        pref <- tmp
      }
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
