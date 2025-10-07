# rule_factory.R ---------------------------------------------------------------
# Generador del Plan de Limpieza (G-aware) desde leer_xlsform_limpieza()

suppressPackageStartupMessages({
  library(dplyr)
  library(purrr)
  library(stringr)
  library(tidyr)
  library(tibble)
})

# =============================================================================
# Helpers básicos y robustos
# =============================================================================

`%||%` <- function(a, b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a

# Siempre devuelve character(1)
as_chr1 <- function(x) {
  if (is.null(x) || length(x) == 0) return("")
  x <- suppressWarnings(as.character(x))
  if (!length(x) || is.na(x[1])) "" else x[1]
}
# "nzchar" seguro para character(1)
nz1 <- function(x) is.character(x) && length(x) == 1 && nzchar(x)

.norm_token <- function(x){
  x <- trimws(as.character(x))
  x <- iconv(x, from = "", to = "ASCII//TRANSLIT")   # quita acentos (Sí -> Si)
  tolower(x)
}
is_yes <- function(x) {
  v <- .norm_token(x)
  v %in% c("yes","si","y","s","1","true","verdadero")
}
is_no <- function(x) {
  v <- .norm_token(x)
  v %in% c("no","n","0","false","falso")
}

regex_escape <- function(s) gsub("([\\^$.|?*+(){}\\[\\]\\\\])", "\\\\\\1", s, perl = TRUE)
expr_no_vacio <- function(v) sprintf("(!is.na(%s) & trimws(%s) != \"\")", v, v)

YES_SET_R <- "c('Yes','yes','Si','si','Sí','sí')"
NO_SET_R  <- "c('No','no')"


# --- Helpers requeridos por nombres de regla -------------------------------
`%||%` <- function(a, b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a
nz1 <- function(x) is.character(x) && length(x)==1 && !is.na(x) && nzchar(x)

sanitize_id <- function(x){
  x <- as.character(x %||% "")
  x <- gsub("\\s+", "_", x)           # espacios -> _
  x <- gsub("[^A-Za-z0-9_]", "_", x)  # cualquier cosa no [A-Za-z0-9_] -> _
  x
}


nombre_regla_consistencia_g <- function(var1, var2 = NULL, pref = "cruce"){
  v1 <- sanitize_id(var1 %||% "var")
  v2 <- if (is.null(var2) || !nzchar(var2) || is.na(var2)) "condicion" else sanitize_id(var2)
  paste0(pref, "_", v1, "x", v2, "_G")
}


# -------------------------------------------------------------------
# Wrapper compatible: reescribir_odk_fix() -> usa tu constraint_a_r()
# Soporta el caso count-selected(.) reemplazando "." por var1
# -------------------------------------------------------------------
reescribir_odk_fix <- function(txt, var1 = NULL) {
  if (is.null(txt) || is.na(txt) || !nzchar(txt)) return(NA_character_)
  x <- as.character(txt)

  # Reemplaza count-selected(.) por count-selected(var1) si corresponde
  if (!is.null(var1) && nzchar(var1)) {
    x <- gsub(
      "count\\s*-\\s*selected\\s*\\(\\s*\\.\\s*\\)",
      paste0("count-selected(", var1, ")"),
      x, perl = TRUE
    )
  }

  # Reutiliza tu traductor principal
  constraint_a_r(x, var_name = var1)
}


# ------------------------------------------------------------------
# Helper para armar gmap (G-aware) de forma robusta
# ------------------------------------------------------------------

# --- Parser robusto para group_relevant (section_map) -----------------
.parse_group_rel <- function(txt, survey, choices, label_col){
  YES_SET_R <- "c('Yes','yes','Si','si','Sí','sí')"
  NO_SET_R  <- "c('No','no')"

  x <- as.character(txt %||% "")
  x <- gsub("[\u00A0\u2007\u202F]", " ", x)
  x <- gsub("\u201C|\u201D", "\"", x)
  x <- gsub("\u2018|\u2019", "'",  x)
  x <- gsub("\\s+", " ", trimws(x))
  if (!nzchar(x)) return(list(expr_r = "", human = "", vars = character(0)))

  expr_out  <- x
  human_out <- x

  # selected(${var}, 'opt')
  m <- stringr::str_match_all(x, "selected\\(\\$\\{([A-Za-z0-9_]+)\\},\\s*'([^']+)'\\)")
  if (length(m) > 0 && length(m[[1]]) > 0) {
    mm <- m[[1]]
    for (k in seq_len(nrow(mm))) {
      var <- mm[k, 2]; opt <- mm[k, 3]
      pat <- paste0("(^|\\s)", gsub("([\\^$.|?*+(){}\\[\\]\\\\])","\\\\\\1", opt), "(\\s|$)")
      expr_r <- paste0('grepl("', pat, '", ', var, ', perl = TRUE)')

      lt <- survey$type[survey$name == var][1]
      ln <- if (!is.na(lt)) sub("^select_(one|multiple)\\s+", "", tolower(lt)) else ""
      var_lab <- etq_pregunta(survey, label_col, var)
      opt_lab <- etq_choice(choices, ln, opt, label_col)
      human   <- paste0("marcó «", opt_lab, "» en «", var_lab, "»")

      token <- paste0("selected(${", var, "}, '", opt, "')")
      expr_out  <- gsub(token, expr_r, expr_out,  fixed = TRUE)
      human_out <- gsub(token, human,  human_out, fixed = TRUE)
    }
  }

  # ${var} -> var
  expr_out  <- gsub("\\$\\{([A-Za-z0-9_]+)\\}", "\\1",  expr_out,  perl = TRUE)
  human_out <- gsub("\\$\\{([A-Za-z0-9_]+)\\}", "«\\1»", human_out, perl = TRUE)

  # and/or/not
  expr_out  <- gsub("(?i)\\band\\b", "&", expr_out,  perl = TRUE)
  expr_out  <- gsub("(?i)\\bor\\b",  "|", expr_out,  perl = TRUE)
  expr_out  <- gsub("(?i)\\bnot\\b", "!", expr_out,  perl = TRUE)
  human_out <- gsub("(?i)\\band\\b", "y",  human_out, perl = TRUE)
  human_out <- gsub("(?i)\\bor\\b",  "o",  human_out, perl = TRUE)
  human_out <- gsub("(?i)\\bnot\\b", "no", human_out, perl = TRUE)

  # "=" suelto -> "=="
  expr_out <- gsub("(?<!<|>|!|<-|=)=(?!=)", "==", expr_out, perl = TRUE)

  # == Yes / No → sets
  expr_out <- gsub("(?i)\\b([A-Za-z0-9_]+)\\s*==?\\s*(['\"])(yes|si|s[ií]|sí)\\2",
                   paste0("\\1 %in% ", YES_SET_R), expr_out, perl = TRUE)
  expr_out <- gsub("(?i)\\b([A-Za-z0-9_]+)\\s*==?\\s*(['\"])(no)\\2",
                   paste0("\\1 %in% ", NO_SET_R),  expr_out, perl = TRUE)

  expr_out  <- gsub("\\s+", " ", trimws(expr_out))
  human_out <- gsub("\\s+", " ", trimws(human_out))

  # drivers desde la cadena original
  mdrv <- stringr::str_match_all(x, "\\$\\{([A-Za-z0-9_]+)\\}")
  drivers <- if (length(mdrv) && nrow(mdrv[[1]])>0) unique(mdrv[[1]][,2]) else character(0)

  list(expr_r = expr_out, human = human_out, vars = drivers)
}


# --- Fallback permisivo (por si el parser formal no detecta nada) ------------
.g_relax <- function(txt){
  if (is.null(txt)) return("")
  x <- as.character(txt)
  x <- gsub("[\u00A0\u2007\u202F]", " ", x)
  x <- gsub("\u201C|\u201D", "\"", x)
  x <- gsub("\u2018|\u2019", "'",  x)
  x <- gsub("\\$\\{([A-Za-z0-9_]+)\\}", "\\1", x, perl = TRUE)  # ${var} -> var
  x <- gsub("(?i)\\band\\b", "&", x, perl = TRUE)
  x <- gsub("(?i)\\bor\\b",  "|", x, perl = TRUE)
  x <- gsub("(?i)\\bnot\\b", "!", x,  perl = TRUE)
  x <- gsub("(?<!<|>|!|<-|=)=(?!=)", "==", x, perl = TRUE)      # = suelto -> ==
  gsub("\\s+", " ", trimws(x))
}

# --- Parser "bonito" (lo que ya tienes) envolviéndolo en try ------------------
.parse_group_rel <- function(txt, survey, choices, label_col){
  out <- tryCatch(
    relevant_a_r_y_humano(txt, survey, choices, label_col),
    error = function(e) list(expr_r = "", human = "", vars = character(0))
  )
  out$expr_r  <- as_chr1(out$expr_r %||% "")
  out$human   <- as_chr1(out$human  %||% "")
  out$vars    <- out$vars %||% character(0)

  # Fallback: si había texto pero el parser dejó expr vacío, usa la versión permisiva
  if (nzchar(as.character(txt %||% "")) && !nzchar(out$expr_r)) {
    out$expr_r <- .g_relax(txt)
    # human: deja tal cual el raw (limpio)
    out$human  <- .g_relax(txt)
  }
  out
}

# --- GMAP “compat” solo desde section_map (sin warnings) ----------------------
.make_gmap <- function(x){
  survey   <- x$survey
  choices  <- x$choices %||% tibble::tibble()
  labelcol <- x$meta$label_col_survey %||% "_label_"

  sm <- x$meta$section_map %||% tibble::tibble()
  if (!nrow(sm)) {
    return(tibble::tibble(group_name=character(), G_expr=character(), G_humano=character(), G_vars=list()))
  }
  if (!"group_relevant" %in% names(sm)) sm$group_relevant <- ""

  sm$group_name     <- trimws(as.character(sm$group_name))
  sm$group_relevant <- as.character(sm$group_relevant)
  sm$group_relevant[is.na(sm$group_relevant)] <- ""

  parsed <- purrr::map(
    sm$group_relevant,
    ~ .parse_group_rel(.x, survey, choices, labelcol)
  )

  tibble::tibble(
    group_name = sm$group_name,
    G_expr     = vapply(parsed, function(p) as_chr1(p$expr_r), ""),
    G_humano   = vapply(parsed, function(p) as_chr1(p$human),  ""),
    G_vars     = lapply(parsed,  function(p) p$vars)
  )
}


# Asigna group_name a todas las filas del survey recorriendo begin_group / end_group
.ensure_group_names <- function(survey) {
  if (!"type" %in% names(survey)) return(survey)
  n <- nrow(survey); if (!n) return(survey)

  type_base <- tolower(trimws(sub("\\s.*$", "", as.character(survey$type))))
  nm        <- as.character(survey$name %||% "")
  grp_stack <- character(0)
  out       <- character(n)

  for (i in seq_len(n)) {
    tb <- type_base[i]
    if (identical(tb, "begin_group")) {
      # push
      gname <- if (!is.na(nm[i]) && nzchar(nm[i])) nm[i] else paste0("group_", i)
      grp_stack <- c(grp_stack, gname)
      out[i] <- tail(grp_stack, 1)
    } else if (identical(tb, "end_group")) {
      # asigna el nombre del grupo que cierra y luego pop
      out[i] <- if (length(grp_stack)) tail(grp_stack, 1) else NA_character_
      if (length(grp_stack)) grp_stack <- grp_stack[-length(grp_stack)]
    } else {
      out[i] <- if (length(grp_stack)) tail(grp_stack, 1) else NA_character_
    }
  }

  survey$group_name <- dplyr::coalesce(as.character(survey$group_name %||% NA_character_), out)
  survey
}



# Normaliza choice_filter a "" (vacío) en vez de NA, tanto en survey como en el summary
coalesce_choice_filter <- function(x) {
  if ("survey" %in% names(x) && is.data.frame(x$survey)) {
    if (!"choice_filter" %in% names(x$survey)) {
      x$survey$choice_filter <- ""
    } else {
      x$survey$choice_filter <- ifelse(is.na(x$survey$choice_filter), "", x$survey$choice_filter)
    }
  }
  if ("meta" %in% names(x) && is.list(x$meta) &&
      "choice_filter_summary" %in% names(x$meta) &&
      is.data.frame(x$meta$choice_filter_summary)) {
    cfs <- x$meta$choice_filter_summary
    if (!"choice_filter" %in% names(cfs)) {
      cfs$choice_filter <- ""
    } else {
      cfs$choice_filter <- ifelse(is.na(cfs$choice_filter), "", cfs$choice_filter)
    }
    x$meta$choice_filter_summary <- cfs
  }
  x
}


# --- Helpers CF ---------------------------------------------------------------

# Normaliza texto ligero (comillas y espacios)
.cf_norm <- function(x){
  x <- as.character(x %||% "")
  x <- gsub("[\u00A0\u2007\u202F]", " ", x)
  x <- gsub("\u201C|\u201D", "\"", x)
  x <- gsub("\u2018|\u2019", "'",  x)
  gsub("\\s+", " ", trimws(x))
}

# Traduce el choice_filter de una pregunta a una expresión R por Opción,
# sustituyendo:
#  - ${var}           -> as.character(var)
#  - name             -> '<optname>'
#  - columnas extra   -> '<valor_en_choices_para_esa_opción>'
#  - and/or/not/=/!=  -> & | !  y  == / !=
# Reemplaza la versión actual por esta
.cf_expr_por_opcion <- function(cf_raw, opt_row, extra_cols_for_list){
  if (is.null(cf_raw) || is.na(cf_raw) || !nzchar(cf_raw)) return("")

  cf <- gsub("(?<!<|>|!|<-|=)=(?!=)", "==", cf_raw, perl = TRUE)
  cf <- gsub("(?i)\\band\\b", "&", cf, perl = TRUE)
  cf <- gsub("(?i)\\bor\\b",  "|", cf, perl = TRUE)

  nm <- as.character(opt_row$name[[1]])
  if (grepl("name\\s*==?\\s*'0'", cf, perl = TRUE) && identical(nm, "0")) return("TRUE")

  used_cols <- unique({
    m <- stringr::str_match_all(cf, "\\bfilter_([A-Za-z0-9_]+)\\b")[[1]]
    if (nrow(m)) paste0("filter_", m[,2]) else character(0)
  })

  pieces <- character(0)
  for (col in extra_cols_for_list) {
    if (!length(used_cols) || !(col %in% used_cols)) next
    if (!col %in% names(opt_row)) next
    val <- as.character(opt_row[[col]][[1]])
    if (!length(val) || is.na(val) || !nzchar(val)) next  # ← ignora vacíos
    var_resp <- sub("^filter_", "", col)                  # filter_P14 → P14
    val_q    <- paste0("'", gsub("'", "\\\\'", val), "'") # siempre texto
    pieces   <- c(pieces, paste0("as.character(", var_resp, ") == ", val_q))
  }
  if (!length(pieces)) "" else paste(pieces, collapse = " | ") # OR entre columnas
}

# Une OR seguro ignorando vacíos
.or_join <- function(parts){
  parts <- parts[nzchar(parts)]
  if (!length(parts)) "" else paste(parts, collapse = " | ")
}



# =============================================================================
# Utils para columnas y etiquetas
# =============================================================================
`%||%` <- `%||%`

.columna_label_segura <- function(x) {
  labelcol <- x$meta$label_col_survey
  if (!is.character(labelcol) || length(labelcol)!=1 || is.na(labelcol)) labelcol <- "_label_"
  survey <- x$survey
  if (!labelcol %in% names(survey)) {
    nms <- tolower(names(survey))
    cand <- c(grep("^label(::|:|_)?(es|spanish).*", nms, value = TRUE),
              grep("^label$", nms, value = TRUE))
    if (length(cand)) {
      survey[["_label_"]] <- as.character(survey[[ names(survey)[match(cand[1], nms)] ]])
    } else survey[["_label_"]] <- rep("", nrow(survey))
    labelcol <- "_label_"
  } else {
    survey[[labelcol]] <- as.character(survey[[labelcol]])
  }
  x$survey <- survey
  list(labelcol = labelcol, survey = survey)
}

prefijo_para <- function(group_name, section_map) {
  if (!is.null(group_name) && nzchar(group_name)) {
    hit <- section_map$prefix[match(group_name, section_map$group_name)]
    if (!is.na(hit)) return(hit)
  }
  "GEN_"
}

etq_pregunta <- function(survey, label_col, var) {
  if (!("name" %in% names(survey))) return(as.character(var))
  row <- survey[survey$name == var, , drop = FALSE]
  if (nrow(row) == 0L) return(as.character(var))
  lab <- row[[label_col]][[1]]
  if (is.function(lab)) return(as.character(var))
  if (is.list(lab)) lab <- unlist(lab, use.names = FALSE)
  lab <- suppressWarnings(as.character(lab))
  out <- if (!length(lab) || is.na(lab) || !nzchar(lab)) var else lab
  gsub("\\s+", " ", trimws(out))
}

list_name_from_type <- function(type_txt) {
  tb <- tolower(trimws(type_txt %||% ""))
  sub("^select_(one|multiple)\\s+", "", tb)
}

etq_choice <- function(choices, list_name, opt_name, label_col) {
  if (is.null(choices) || !is.data.frame(choices)) return(as.character(opt_name))
  req <- c("list_name","name", label_col)
  if (!all(req %in% names(choices))) return(as.character(opt_name))
  row <- choices[choices$list_name == list_name & choices$name == opt_name, , drop = FALSE]
  if (nrow(row) == 0L) return(as.character(opt_name))
  lab <- row[[label_col]][[1]]
  if (is.function(lab)) return(as.character(opt_name))
  if (is.list(lab)) lab <- unlist(lab, use.names = FALSE)
  lab <- suppressWarnings(as.character(lab))
  out <- if (!length(lab) || is.na(lab) || !nzchar(lab)) opt_name else lab
  gsub("\\s+", " ", trimws(out))
}

# =============================================================================
# Parseos (relevant/constraint/calculate) y helpers G
# =============================================================================

`%||%` <- function(a, b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a
as_chr1 <- function(x) if (length(x)==0) "" else as.character(x[[1]])

# extrae posibles drivers (nombres de variables) desde una expresión en R
drivers_from_expr <- function(expr_r, survey_names) {
  if (is.null(expr_r) || !nzchar(expr_r)) return(character(0))
  m <- gregexpr("\\b[A-Za-z_][A-Za-z0-9_]*\\b", expr_r, perl = TRUE)[[1]]
  if (identical(m, -1L)) return(character(0))
  words <- substring(expr_r, m, m + attr(m, "match.length") - 1L)
  blacklist <- c("is_na","is","na","TRUE","FALSE","T","F","if","else",
                 "as","numeric","character","integer","logical","trimws",
                 "suppressWarnings","grepl","str_count","stringr",
                 "count","selected","is_yes","is_no")
  words <- setdiff(unique(words), blacklist)
  words[words %in% survey_names]
}



# selected(${var}, 'opt') y YES/NO -> (expr R, human, vars)
relevant_a_r_y_humano <- function(rel_raw, survey, choices, label_col) {
  label_col <- as.character(label_col)[1]
  x <- if (is.null(rel_raw)) NA_character_ else suppressWarnings(as.character(rel_raw))
  if (!length(x) || is.na(x) || !nzchar(x)) {
    return(list(expr_r = "", human = "", vars = character(0)))
  }
  x <- gsub("[\u00A0\u2007\u202F]", " ", x)
  x <- gsub("\u201C|\u201D", "\"", x)
  x <- gsub("\u2018|\u2019", "'",  x)
  x <- gsub("\\s+", " ", trimws(x))
  original <- x

  expr_out  <- x
  human_out <- x

  # selected(${var}, 'opt')
  m <- stringr::str_match_all(x, "selected\\(\\$\\{([A-Za-z0-9_]+)\\},\\s*'([^']+)'\\)")
  if (length(m) > 0 && length(m[[1]]) > 0) {
    mm <- m[[1]]
    for (k in seq_len(nrow(mm))) {
      var <- mm[k, 2]; opt <- mm[k, 3]
      opt_rx <- regex_escape(opt)
      pat    <- paste0("(^|\\s)", opt_rx, "(\\s|$)")
      expr_r <- paste0('grepl("', pat, '", ', var, ')')
      lt <- survey$type[survey$name == var]
      ln <- if (length(lt) && !is.na(lt[1])) list_name_from_type(lt[1]) else ""
      var_lab <- etq_pregunta(survey, label_col, var)
      opt_lab <- etq_choice(choices, ln, opt, label_col)
      human   <- paste0("marcó «", opt_lab, "» en «", var_lab, "»")
      token <- paste0("selected(${", var, "}, '", opt, "')")
      expr_out  <- gsub(token, expr_r, expr_out,  fixed = TRUE)
      human_out <- gsub(token, human,  human_out, fixed = TRUE)
    }
  }

  # Sustituir ${var} -> var (para el resto)
  expr_out <- gsub("\\$\\{([A-Za-z0-9_]+)\\}", "\\1", expr_out, perl = TRUE)
  human_out<- gsub("\\$\\{([A-Za-z0-9_]+)\\}", "«\\1»", human_out, perl = TRUE)

  # Operadores lógicos
  expr_out  <- gsub("(?i)\\band\\b", "&", expr_out,  perl = TRUE)
  expr_out  <- gsub("(?i)\\bor\\b",  "|", expr_out,  perl = TRUE)
  expr_out  <- gsub("(?i)\\bnot\\b", "!", expr_out,  perl = TRUE)
  human_out <- gsub("(?i)\\band\\b", "y",  human_out, perl = TRUE)
  human_out <- gsub("(?i)\\bor\\b",  "o",  human_out, perl = TRUE)
  human_out <- gsub("(?i)\\bnot\\b", "no", human_out, perl = TRUE)

  # Igualdades/Desigualdades → YES/NO sets
  # == YES/SI/SÍ
  expr_out <- gsub("(?i)\\b([A-Za-z0-9_]+)\\s*==?\\s*(['\"])(yes|si|s[ií]|sí)\\2",
                   paste0("\\1 %in% ", YES_SET_R), expr_out, perl = TRUE)
  # YES == var
  expr_out <- gsub("(?i)(['\"])(yes|si|s[ií]|sí)\\1\\s*==?\\s*\\b([A-Za-z0-9_]+)\\b",
                   paste0("\\3 %in% ", YES_SET_R), expr_out, perl = TRUE)

  # != YES -> !(var %in% YES)
  expr_out <- gsub("(?i)\\b([A-Za-z0-9_]+)\\s*!?=\\s*(['\"])(yes|si|s[ií]|sí)\\2",
                   paste0("!(\\1 %in% ", YES_SET_R, ")"), expr_out, perl = TRUE)
  expr_out <- gsub("(?i)(['\"])(yes|si|s[ií]|sí)\\1\\s*!?=\\s*\\b([A-Za-z0-9_]+)\\b",
                   paste0("!(\\3 %in% ", YES_SET_R, ")"), expr_out, perl = TRUE)

  # == NO
  expr_out <- gsub("(?i)\\b([A-Za-z0-9_]+)\\s*==?\\s*(['\"])(no)\\2",
                   paste0("\\1 %in% ", NO_SET_R), expr_out, perl = TRUE)
  expr_out <- gsub("(?i)(['\"])(no)\\1\\s*==?\\s*\\b([A-Za-z0-9_]+)\\b",
                   paste0("\\3 %in% ", NO_SET_R), expr_out, perl = TRUE)

  # != NO -> !(var %in% NO)
  expr_out <- gsub("(?i)\\b([A-Za-z0-9_]+)\\s*!?=\\s*(['\"])(no)\\2",
                   paste0("!(\\1 %in% ", NO_SET_R, ")"), expr_out, perl = TRUE)
  expr_out <- gsub("(?i)(['\"])(no)\\1\\s*!?=\\s*\\b([A-Za-z0-9_]+)\\b",
                   paste0("!(\\3 %in% ", NO_SET_R, ")"), expr_out, perl = TRUE)

  expr_out  <- gsub("(?<![!<>=])=(?!=)", "==", expr_out, perl = TRUE)
  expr_out  <- gsub("\\s+", " ", trimws(expr_out))
  human_out <- gsub("\\s+", " ", trimws(human_out))

  # Drivers (variables implicadas)
  drivers <- {
    mdrv <- stringr::str_match_all(original, "\\$\\{([A-Za-z0-9_]+)\\}")
    if (length(mdrv) > 0 && nrow(mdrv[[1]]) > 0) unique(mdrv[[1]][,2]) else character(0)
  }

  # Forzar a character(1)
  expr_out  <- as_chr1(expr_out)
  human_out <- as_chr1(human_out)

  list(expr_r = expr_out, human = human_out, vars = drivers)
}

# -------------------------------------------------------------------
# Constraint ODK -> R (con soporte count-selected para select_multiple)
# -------------------------------------------------------------------



# --- FIX: ODK -> R, respetando "." como variable actual -----------------------
constraint_a_r <- function(txt, var_name = NULL) {
  if (is.null(txt) || !nzchar(trimws(txt))) return(NA_character_)
  x <- as.character(txt)

  # 0) Sustituye "." por la variable actual
  if (!is.null(var_name) && nzchar(var_name)) {
    x <- gsub("(?<=\\W)\\.\\b", var_name, x, perl = TRUE)
    x <- gsub("^\\.", var_name, x, perl = TRUE)
  }

  # 1) ${var} -> var
  x <- gsub("\\$\\{([A-Za-z0-9_]+)\\}", "\\1", x, perl = TRUE)

  # 2) Operadores lógicos
  x <- gsub("(?i)\\band\\b", "&", x, perl = TRUE)
  x <- gsub("(?i)\\bor\\b",  "|", x, perl = TRUE)
  x <- gsub("(?i)\\bnot\\b", "!", x,  perl = TRUE)

  # 3) selected(var,'opt') -> grepl("(^|\\s)opt(\\s|$)", var)
  x <- gsub(
    "selected\\s*\\(\\s*([A-Za-z0-9_]+)\\s*,\\s*'([^']+)'\\s*\\)",
    'grepl("(^|\\\\s)\\2(\\\\s|$)", \\1)', x, perl = TRUE
  )

  # 4) regex(var,'pat') -> stringr::str_detect(var, 'pat')
  x <- gsub(
    "regex\\s*\\(\\s*([A-Za-z0-9_]+)\\s*,\\s*'([^']+)'\\s*\\)",
    "stringr::str_detect(\\1, '\\2')", x, perl = TRUE
  )

  # 5) count-selected(var)
  x <- gsub(
    "count-selected\\s*\\(\\s*([A-Za-z0-9_]+)\\s*\\)",
    "ifelse(is.na(\\1)|trimws(\\1)==\"\",0,1+stringr::str_count(\\1,\"\\\\s+\"))",
    x, perl = TRUE
  )

  # 6) Limpieza
  gsub("\\s+", " ", trimws(x))
}

# Wrapper: reescribir_odk_fix() → ahora normaliza "." y luego traduce a R
reescribir_odk_fix <- function(txt, var1 = NULL) {
  if (is.null(txt) || is.na(txt) || !nzchar(txt)) return(NA_character_)
  x <- as.character(txt)

  # Primero: sustituir el "." por var1 en funciones y comparaciones comunes
  x <- ._swap_dot_with_var1(x, var1)

  # Luego: tu traductor principal a R
  constraint_a_r(x, var_name = var1)
}

# --- Helper: detecta rangos de año a partir de patrón regex clásico -----------
.range_from_year_regex <- function(pat) {
  # Busca subpatrones como 19[8-9][0-9], 200[0-9], 201[0-6], 202[0-4], etc.
  frags <- unlist(regmatches(pat, gregexpr("(19\\[\\d-\\d\\]\\[\\d\\]|19\\d\\[\\d\\]|200\\[\\d\\]|201\\[\\d\\]|202\\[\\d\\])", pat)))
  # fallback: también acepta bloques '19[8-9][0-9]' completos o '201[0-6]'
  if (!length(frags)) frags <- unlist(regmatches(pat, gregexpr("(19\\[\\d-\\d\\]\\d|20\\d\\[\\d\\])", pat)))

  # Método práctico: prueba todos los años razonables y qué años acepta el patrón
  years <- 1900:2100
  ok <- grepl(paste0("^", pat, "$"), as.character(years))
  if (any(ok)) {
    rng <- range(years[ok])
    return(rng)
  }
  NULL
}

# --- NLG: Explicación en español a partir de la expresión R -------------------
constraint_a_es <- function(expr_r, var_lab) {
  if (is.null(expr_r) || is.na(expr_r) || !nzchar(expr_r)) {
    return(paste0("El valor registrado en «", var_lab, "» respeta la regla del formulario."))
  }
  x <- expr_r

  # 1) Rango numérico: a <= var <= b
  m <- stringr::str_match(x,
                          "as\\.numeric\\([^\\)]+\\)\\s*>=\\s*([0-9\\.]+)\\s*&\\s*as\\.numeric\\([^\\)]+\\)\\s*<=\\s*([0-9\\.]+)"
  )
  if (!any(is.na(m))) {
    return(paste0("El valor de «", var_lab, "» debe estar entre ", m[2], " y ", m[3], "."))
  }

  # 2) Cotas simples
  m <- stringr::str_match(x, "as\\.numeric\\([^\\)]+\\)\\s*<=\\s*([0-9\\.]+)")
  if (!any(is.na(m))) return(paste0("El valor de «", var_lab, "» no debe superar ", m[2], "."))
  m <- stringr::str_match(x, "as\\.numeric\\([^\\)]+\\)\\s*>=\\s*([0-9\\.]+)")
  if (!any(is.na(m))) return(paste0("El valor de «", var_lab, "» no debe ser menor que ", m[2], "."))

  # 3) Igualdad textual simple: var == 'X'
  m <- stringr::str_match(x, "(\\b[A-Za-z0-9_]+\\b)\\s*==\\s*'([^']+)'")
  if (!any(is.na(m))) {
    return(paste0("El valor de «", var_lab, "» debe ser «", m[3], "»."))
  }

  # 4) Pertenencia a conjunto: var %in% c('A','B',...)
  m <- stringr::str_match(x, "(\\b[A-Za-z0-9_]+\\b)\\s*%in%\\s*c\\(([^\\)]*)\\)")
  if (!any(is.na(m))) {
    elems <- strsplit(m[3], ",")[[1]] |> trimws() |> gsub("^['\"]|['\"]$", "", x = _)
    elems <- elems[nzchar(elems)]
    if (length(elems)) {
      return(paste0("El valor de «", var_lab, "» debe pertenecer a {", paste(elems, collapse = ", "), "}."))
    }
  }

  # 5) selected(var,'opt') -> grepl("(^|\\s)opt(\\s|$)", var)
  m <- stringr::str_match(x, 'grepl\\("\\(\\^\\|\\\\s\\)([^"]+)\\(\\\\s\\|\\$\\)",\\s*([A-Za-z0-9_]+)\\)')
  if (!any(is.na(m))) {
    opt <- m[2]
    return(paste0("Debe estar marcada la opción «", opt, "» en «", var_lab, "»."))
  }

  # 6) count-selected(var) [op] K
  m <- stringr::str_match(
    x,
    'ifelse\\(is\\.na\\((\\w+)\\)\\s*\\|\\s*trimws\\(\\1\\)\\s*==\\s*""\\,\\s*0\\,\\s*1\\s*\\+\\s*stringr::str_count\\(\\1\\,\\s*"\\\\\\\\s\\+"\\)\\)\\s*(<=|<|>=|>)\\s*([0-9]+)'
  )
  if (!any(is.na(m))) {
    op <- m[3]; k <- m[4]
    texto <- switch(op,
                    "<=" = paste0("como máximo ", k, " selección(es)"),
                    "<"  = paste0("menos de ", k, " selección(es)"),
                    ">=" = paste0("al menos ", k, " selección(es)"),
                    ">"  = paste0("más de ", k, " selección(es)")
    )
    return(paste0("En «", var_lab, "» se permite ", texto, "."))
  }

  # 7) stringr::str_detect(var, '...') — intenta detectar patrón de año
  m <- stringr::str_match(x, "stringr::str_detect\\((\\w+)\\s*,\\s*'([^']+)'\\)")
  if (!any(is.na(m))) {
    pat <- m[3]
    rng <- .range_from_year_regex(pat)
    if (!is.null(rng)) {
      return(paste0("El año en «", var_lab, "» debe estar entre ", rng[1], " y ", rng[2], "."))
    }
    return(paste0("El valor de «", var_lab, "» debe cumplir el formato definido."))
  }

  # 8) Fallback genérico
  paste0("El valor registrado en «", var_lab, "» respeta la regla del formulario.")
}

# =============================================================================
# Builders G-aware
# =============================================================================

nombre_regla_simple  <- function(var) paste0("cruce_",  var, "_x")
nombre_regla_cruce   <- function(a,b) paste0("cruce_",  a, "x", b)
nombre_regla_cruce2  <- function(a,b) paste0("cruce2_", a, "x", b)
nombre_regla_cruce_g <- function(a,b) paste0("cruce_",  a, "x", b, "_G")
nombre_regla_cruce2_g<- function(a,b) paste0("cruce2_", a, "x", b, "_G")
nombre_regla_noG     <- function(a,b) paste0("cruce3_", a, "x", b, "_noG")

# ---- REQUERIDAS (control) SIN DUPLICAR LO QUE YA ES "relevant" --------
build_required_g <- function(survey, section_map, label_col, gmap){
  dat <- survey %>%
    dplyr::mutate(type_base = tolower(trimws(sub("\\s.*$", "", .data$type)))) %>%
    dplyr::filter(
      .data$required,
      .data$type_base %in% c("select_one","select_multiple","integer","decimal","text","date","datetime","string"),
      .data$type_base != "note"
    )

  if (!nrow(dat)) return(tibble::tibble())

  purrr::pmap_dfr(dat, function(...){
    row  <- list(...)
    var  <- row$name
    lab  <- etq_pregunta(survey, label_col, var)
    pref <- prefijo_para(row$group_name, section_map)

    # relevant de la PREGUNTA (si existe)
    rel_parsed <- tryCatch(
      relevant_a_r_y_humano(as_chr1(row$relevant), survey, NULL, label_col),
      error = function(e) list(expr_r = "", human = "", vars = character(0))
    )
    rel_expr <- as_chr1(rel_parsed$expr_r)
    rel_h    <- as_chr1(rel_parsed$human)

    # 🔴 Anti-duplicados:
    # Si la pregunta YA tiene relevant, NO generamos regla "control" aquí
    # (ese caso lo cubre build_relevant_g).
    if (nz1(rel_expr)) return(tibble::tibble())

    # Info del G del grupo (si existe)
    ginfo <- gmap[gmap$group_name == row$group_name, , drop = FALSE]
    G  <- as_chr1(if (nrow(ginfo)) ginfo$G_expr[[1]]   else "")
    Gh <- as_chr1(if (nrow(ginfo)) ginfo$G_humano[[1]] else "")
    Gv <- if (nrow(ginfo)) (ginfo$G_vars[[1]] %||% character(0)) else character(0)

    # Con G del grupo: usar primer driver como Variable 2 (si existe)
    if (nz1(G)) {
      var2     <- if (length(Gv)) Gv[1] else "condicion"
      var2_lab <- if (length(Gv)) etq_pregunta(survey, label_col, var2) else Gh

      nom1 <- nombre_regla_cruce_g(var, var2)
      nom2 <- nombre_regla_cruce2_g(var, var2)

      obj1 <- paste0("En caso la persona cumpla: (", Gh, "), entonces **DEBE** responder «", lab, "».")
      pr1  <- paste0(nom1, " <- ( !(", G, ") | ", expr_no_vacio(var), " )")

      obj2 <- paste0("Si NO se cumple la condición: (", Gh, "), entonces **NO DEBE** responder «", lab, "».")
      pr2  <- paste0(nom2, " <- ( (", G, ") | (is.na(", var, ") | trimws(", var, ") == \"\") )")

      return(tibble::tibble(
        ID = NA_character_,
        `Tipo de observación`   = "Preguntas de control",
        Objetivo                = c(obj1, obj2),
        `Variable 1`            = c(var, var),
        `Variable 1 - Etiqueta` = c(lab, lab),
        `Variable 2`            = c(var2, var2),
        `Variable 2 - Etiqueta` = c(var2_lab, var2_lab),
        `Variable 3`            = c(NA_character_, NA_character_),
        `Variable 3 - Etiqueta` = c(NA_character_, NA_character_),
        `Nombre de regla`       = c(nom1, nom2),
        `Procesamiento`         = c(pr1, pr2),
        .pref = pref,
        .gord = section_map$.gord[match(row$group_name, section_map$group_name)],
        .qord = row$.qord
      ))
    }

    # Sin relevant ni G → control simple
    nombre <- nombre_regla_simple(var)
    tibble::tibble(
      ID = NA_character_,
      `Tipo de observación`   = "Preguntas de control",
      Objetivo                = paste0("**DEBE** responder «", lab, "»."),
      `Variable 1`            = var,
      `Variable 1 - Etiqueta` = lab,
      `Variable 2`            = NA_character_,
      `Variable 2 - Etiqueta` = NA_character_,
      `Variable 3`            = NA_character_,
      `Variable 3 - Etiqueta` = NA_character_,
      `Nombre de regla`       = nombre,
      `Procesamiento`         = paste0(nombre, " <- ", expr_no_vacio(var)),
      .pref = pref,
      .gord = section_map$.gord[match(row$group_name, section_map$group_name)],
      .qord = row$.qord
    )
  })
}

# ---- OTHER (A×B; con G: espejo + ¬G) ----------------------------------------

PATRON_OTHER_LABEL <- "(?i)(^|\\s)(other|otro|otra)(\\s|$)"

detectar_other_links <- function(survey, label_col) {
  survey %>%
    mutate(type_base = tolower(trimws(sub("\\s.*$", "", .data$type)))) %>%
    filter(type_base %in% c("text","string"), !is.na(.data$relevant), nzchar(.data$relevant)) %>%
    mutate(parent = str_match(.data$relevant, "selected\\(\\$\\{([A-Za-z0-9_]+)\\},\\s*'Other'\\)")[,2]) %>%
    filter(!is.na(parent), nzchar(parent)) %>%
    transmute(
      parent_var   = parent,
      other_var    = name,
      group_name   = .data$group_name,
      .qord        = .data$.qord
    ) %>%
    distinct()
}

build_other_g <- function(survey, section_map, label_col, gmap){
  links <- detectar_other_links(survey, label_col)
  if (!nrow(links)) return(tibble())

  pmap_dfr(links, function(parent_var, other_var, group_name, .qord){
    pref  <- prefijo_para(group_name, section_map)
    labP  <- etq_pregunta(survey, label_col, parent_var)
    labO  <- etq_pregunta(survey, label_col, other_var)
    ginfo <- gmap[gmap$group_name == group_name, , drop = FALSE]

    # "Otros" (texto y R)
    cond_h_opt <- paste0("marcó «Otros» en «", labP, "»")
    cond_r_opt <- paste0('grepl("(^|\\s)other(\\s|$)", ', parent_var, ', perl = TRUE)')

    if (nrow(ginfo) && .nz(ginfo$G_expr[[1]])) {
      G   <- as.character(ginfo$G_expr[[1]])
      G_h <- as.character(ginfo$G_humano[[1]])

      gat_var  <- { vv <- ginfo$G_vars[[1]] %||% character(0); if (length(vv)) vv[1] else NA_character_ }
      gat_lab  <- if (.nz(gat_var)) etq_pregunta(survey, label_col, gat_var) else G_h

      nom1 <- nombre_regla_cruce_g(parent_var, other_var)
      nom2 <- nombre_regla_cruce2_g(parent_var, other_var)
      nom3 <- nombre_regla_noG(parent_var, other_var)

      # Condición conjunta SOLO con piezas no vacías
      cond_h_full <- .join_and(.as_paren(G_h), .as_paren(cond_h_opt))
      cond_r_full <- .join_and(.as_paren(G),   .as_paren(cond_r_opt))

      # DEBE: si G & (marcó Otros) => texto no vacío
      pr1 <- paste0(
        nom1, " <- ( !(", cond_r_full, ") | ",
        "(!is.na(", other_var, ") & trimws(", other_var, ") != \"\") )"
      )
      # NO DEBE: si G & !(marcó Otros) => debe estar vacío
      pr2 <- paste0(
        nom2, " <- ( !(", .join_and(.as_paren(G), .as_paren(paste0("!(", cond_r_opt, ")"))), ") | ",
        "(is.na(", other_var, ") | trimws(", other_var, ") == \"\") )"
      )
      # Fuera de G: ambos vacíos
      pr3 <- paste0(
        nom3, " <- ( (", G, ") | ",
        "(is.na(", parent_var, ") | trimws(", parent_var, ") == \"\") & ",
        "(is.na(", other_var,  ") | trimws(", other_var,  ") == \"\") )"
      )

      obj1 <- paste0("En caso la persona cumpla: ", cond_h_full,
                     ", entonces **DEBE** responder «", labO, "».")
      obj2 <- paste0("En caso la persona cumpla: (", G_h, "), si **NO** ", cond_h_opt,
                     ", entonces **NO DEBE** responder «", labO, "».")
      obj3 <- paste0("Si **NO** se cumple: (", G_h, "), entonces **NO DEBE** responder «",
                     labP, "» ni «", labO, "».")

      return(tibble(
        ID = NA_character_,
        `Tipo de observación` = "Saltos de preguntas",
        Objetivo = c(obj1, obj2, obj3),
        `Variable 1` = c(parent_var, parent_var, parent_var),
        `Variable 1 - Etiqueta` = c(labP, labP, labP),
        `Variable 2` = c(other_var, other_var, other_var),
        `Variable 2 - Etiqueta` = c(labO, labO, labO),
        `Variable 3` = c(gat_var, gat_var, gat_var),
        `Variable 3 - Etiqueta` = c(gat_lab, gat_lab, gat_lab),
        `Nombre de regla` = c(nom1, nom2, nom3),
        `Procesamiento` = c(pr1, pr2, pr3),
        .pref = pref,
        .gord = section_map$.gord[match(group_name, section_map$group_name)],
        .qord = .qord
      ))
    }

    # Sin G (no hay NA que pegar)
    nom1 <- nombre_regla_cruce(parent_var, other_var)
    nom2 <- nombre_regla_cruce2(parent_var, other_var)

    pr1  <- paste0(
      nom1, " <- !(",
      cond_r_opt, " & ",
      "(is.na(", other_var, ") | trimws(", other_var, ") == \"\")",
      ")"
    )
    pr2  <- paste0(
      nom2, " <- !(",
      "(!", cond_r_opt, ") & ",
      "(!is.na(", other_var, ") & trimws(", other_var, ") != \"\")",
      ")"
    )

    obj1 <- paste0("Si ", cond_h_opt, ", entonces **DEBE** responder «", labO, "».")
    obj2 <- paste0("Si **NO** ", cond_h_opt, ", entonces **NO DEBE** responder «", labO, "».")
    tibble(
      ID = NA_character_,
      `Tipo de observación` = "2. Saltos de preguntas",
      Objetivo = c(obj1, obj2),
      `Variable 1` = c(parent_var, parent_var),
      `Variable 1 - Etiqueta` = c(labP, labP),
      `Variable 2` = c(other_var, other_var),
      `Variable 2 - Etiqueta` = c(labO, labO),
      `Variable 3` = c(NA_character_, NA_character_),
      `Variable 3 - Etiqueta` = c(NA_character_, NA_character_),
      `Nombre de regla` = c(nom1, nom2),
      `Procesamiento` = c(pr1, pr2),
      .pref = pref,
      .gord = section_map$.gord[match(group_name, section_map$group_name)],
      .qord = .qord
    )
  })
}

# ---- RELEVANT (saltos a una pregunta) ---------------------------------------

# helper chiquito por si no lo tienes
as_chr1 <- function(x) { x <- suppressWarnings(as.character(x)); if (!length(x)) "" else x[1] }

pick_g_driver <- function(vars){
  # elige la 1ra var “driver” de G si existe
  if (is.null(vars) || !length(vars)) return(NA_character_)
  vars[1]
}

# ---- RELEVANT (saltos a una pregunta) ---------------------------------------

# --- Helpers para armar condiciones sin "NA" ni huecos -----------------
.nz <- function(x) {
  is.character(x) && length(x)==1 && !is.na(x) && nzchar(trimws(x))
}
.as_paren <- function(x) {
  if (.nz(x)) paste0("(", trimws(x), ")") else ""
}
.join_and <- function(...) {
  parts <- Filter(.nz, c(...))
  if (!length(parts)) "" else paste(parts, collapse = " & ")
}

#------------------------

# ---- RELEVANT (saltos a una pregunta, G-aware, sin duplicar Var2/Var3) ----
build_relevant_g <- function(survey, section_map, label_col, choices, gmap){
  # helpers locales mínimos (idénticos a los tuyos)
  as_chr1 <- function(x){ x <- suppressWarnings(as.character(x)); if (!length(x)) "" else x[1] }
  nz1     <- function(x) is.character(x) && length(x)==1 && !is.na(x) && nzchar(trimws(x))
  `%||%`  <- function(a,b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a

  # prioriza drivers de G distintos de "Consent"
  prioritize_g <- function(vs){
    if (is.null(vs) || !length(vs)) return(character(0))
    v <- as.character(vs)
    c(v[!tolower(v) %in% "consent"], v[tolower(v) %in% "consent"])
  }

  dat <- survey %>%
    dplyr::mutate(type_base = tolower(trimws(sub("\\s.*$", "", .data$type)))) %>%
    dplyr::filter(type_base != "note", !is.na(.data$name))
  if (!nrow(dat)) return(tibble::tibble())

  purrr::pmap_dfr(dat, function(...){
    row <- list(...)
    var     <- row$name
    var_lab <- etq_pregunta(survey, label_col, var)
    pref    <- prefijo_para(row$group_name, section_map)

    # --- G del grupo desde gmap ---
    ginfo   <- gmap[gmap$group_name == row$group_name, , drop = FALSE]
    G_r     <- as_chr1(if (nrow(ginfo)) ginfo$G_expr[[1]]   else "")
    G_h     <- as_chr1(if (nrow(ginfo)) ginfo$G_humano[[1]] else "")
    G_vars  <- if (nrow(ginfo)) (ginfo$G_vars[[1]] %||% character(0)) else character(0)

    # --- relevant del ítem ---
    rel_parsed <- tryCatch(
      relevant_a_r_y_humano(as_chr1(row$relevant %||% ""), survey, choices, label_col),
      error = function(e) list(expr_r = "", human = "", vars = character(0))
    )
    rel_r    <- as_chr1(rel_parsed$expr_r %||% "")
    rel_h    <- as_chr1(rel_parsed$human  %||% "")
    rel_vars <- rel_parsed$vars %||% character(0)

    # --- condición efectiva (R) ---
    cond_r <- if (nz1(G_r) && nz1(rel_r)) paste0("(", G_r, ") & (", rel_r, ")")
    else if (nz1(G_r))          G_r
    else if (nz1(rel_r))        rel_r
    else                        ""
    if (!nz1(cond_r)) return(tibble::tibble())

    # ===== Variables 2 y 3 (drivers) =====
    # Var2: primero del relevant del ítem; si no hay, primero de G
    if (length(rel_vars) >= 1) {
      var2 <- as_chr1(rel_vars[1])
    } else if (length(G_vars) >= 1) {
      var2 <- as_chr1(G_vars[1])
    } else {
      var2 <- NA_character_
    }
    var2_lab <- if (!is.na(var2)) etq_pregunta(survey, label_col, var2)
    else if (nz1(G_h) || nz1(rel_h)) if (nz1(G_h)) G_h else rel_h
    else NA_character_

    # Var3: tomar de G uno distinto de Var2, priorizando NO 'Consent'
    g_pool   <- prioritize_g(unique(as.character(G_vars)))
    g_pool   <- g_pool[ g_pool != var2 ]  # quitar duplicado
    var3     <- if (length(g_pool)) as_chr1(g_pool[1]) else NA_character_
    var3_lab <- if (!is.na(var3)) etq_pregunta(survey, label_col, var3) else NA_character_

    # DEDUPE final por si acaso
    if (!is.na(var2) && !is.na(var3) && identical(as.character(var2), as.character(var3))) {
      var3 <- NA_character_
      var3_lab <- NA_character_
    }

    # --- nombres de reglas ---
    nom1 <- nombre_regla_cruce_g(var, var2)
    nom2 <- nombre_regla_cruce2_g(var, var2)

    # --- procesamiento ---
    pr1  <- paste0(nom1, " <- ( !(", cond_r, ") | ", expr_no_vacio(var), " )")
    pr2  <- paste0(nom2, " <- ( (", cond_r, ") | (is.na(", var, ") | trimws(", var, ") == \"\") )")

    # --- objetivo humano ---
    cond_h <- if (nz1(G_h) && nz1(rel_h)) paste0("(", G_h, ") y (", rel_h, ")")
    else if (nz1(G_h))          paste0("(", G_h, ")")
    else                        paste0("(", rel_h, ")")

    obj1 <- paste0("En caso la persona cumpla: ", cond_h, ", entonces **DEBE** responder «", var_lab, "».")
    obj2 <- paste0("Si NO se cumple la condición: ", cond_h, ", entonces **NO DEBE** responder «", var_lab, "».")

    tibble::tibble(
      ID                         = NA_character_,
      `Tipo de observación`      = "Saltos de preguntas",
      Objetivo                   = c(obj1, obj2),
      `Variable 1`               = c(var, var),
      `Variable 1 - Etiqueta`    = c(var_lab, var_lab),
      `Variable 2`               = c(var2, var2),
      `Variable 2 - Etiqueta`    = c(var2_lab, var2_lab),
      `Variable 3`               = c(var3, var3),
      `Variable 3 - Etiqueta`    = c(var3_lab, var3_lab),
      `Nombre de regla`          = c(nom1, nom2),
      `Procesamiento`            = c(pr1, pr2),
      .pref                      = pref,
      .gord                      = section_map$.gord[match(row$group_name, section_map$group_name)],
      .qord                      = row$.qord
    )
  })
}


# ---- CONSISTENCIA (constraint) — G-aware, variable gatilladora real, ODK→R con "." ----

# Reemplaza el "." por var1 en los casos típicos de ODK
._swap_dot_with_var1 <- function(x, var1) {
  if (is.null(var1) || !nzchar(var1)) return(x)

  # 1) Funciones con punto como primer argumento
  x <- gsub("(selected\\s*\\(\\s*)\\.\\s*,",           paste0("\\1", var1, ","), x, perl = TRUE)
  x <- gsub("(regex\\s*\\(\\s*)\\.\\s*,",              paste0("\\1", var1, ","), x, perl = TRUE)
  x <- gsub("(count\\s*-?\\s*selected\\s*\\(\\s*)\\.\\s*\\)", paste0("\\1", var1, ")"), x, perl = TRUE)

  # 2) Comparaciones con operador + .
  #    op .  -> op var1
  x <- gsub("(==|!=|<=|>=|<|>|%in%)\\s*\\.", paste0("\\1 ", var1), x, perl = TRUE)
  #    . op  -> var1 op
  x <- gsub("\\.\\s*(==|!=|<=|>=|<|>|%in%)", paste0(var1, " \\1"), x, perl = TRUE)

  # 3) Punto “suelto” al inicio o tras un no-alfa-num subiendo var1
  #    (ej. ". != '0'" o " (.) " como token)
  x <- gsub("(^|[^A-Za-z0-9_])\\.(?=\\b)", paste0("\\1", var1), x, perl = TRUE)

  x
}



build_constraint_g <- function(survey, section_map, label_col, gmap){
  dat <- survey %>%
    dplyr::mutate(type_base = tolower(trimws(sub("\\s.*$", "", .data$type)))) %>%
    dplyr::filter(!is.na(.data$name),
                  !is.na(.data$constraint) & nzchar(.data$constraint),
                  type_base %in% c("select_one","select_multiple","integer","decimal","text","date","datetime","string"),
                  type_base != "note")

  if (!nrow(dat)) return(tibble::tibble())

  purrr::pmap_dfr(dat, function(...){
    row  <- list(...)
    var  <- row$name
    lab  <- etq_pregunta(survey, label_col, var)
    pref <- prefijo_para(row$group_name, section_map)

    # relevant de la pregunta (por si hay condición a nivel de ítem)
    rel_parsed <- tryCatch(
      relevant_a_r_y_humano(as_chr1(row$relevant), survey, NULL, label_col),
      error = function(e) list(expr_r = "", human = "", vars = character(0))
    )
    rel_expr <- as_chr1(rel_parsed$expr_r)
    rel_h    <- as_chr1(rel_parsed$human)
    rel_vars <- rel_parsed$vars %||% character(0)

    # G del grupo
    ginfo  <- gmap[gmap$group_name == row$group_name, , drop = FALSE]
    G      <- as_chr1(if (nrow(ginfo)) ginfo$G_expr[[1]]   else "")
    Gh     <- as_chr1(if (nrow(ginfo)) ginfo$G_humano[[1]] else "")
    G_vars <- if (nrow(ginfo)) (ginfo$G_vars[[1]] %||% character(0)) else character(0)

    # Condición efectiva
    cond_expr <- if (nz1(G) && nz1(rel_expr)) paste0("(", G, ") & (", rel_expr, ")")
    else if (nz1(G))              G
    else if (nz1(rel_expr))       rel_expr
    else                          ""

    # ODK constraint → R, pasando var1 para cubrir selected(., ...) / count-selected(.)
    constraint_raw <- as_chr1(row$constraint)
    rhs_r <- reescribir_odk_fix(constraint_raw, var1 = var)
    if (!nz1(rhs_r)) return(tibble::tibble())

    # Variable gatilladora (Variable 2)
    drivers <- unique(c(G_vars, rel_vars))
    if (length(drivers) >= 1) {
      var2     <- drivers[1]
      var2_lab <- etq_pregunta(survey, label_col, var2)
    } else {
      var2     <- NA_character_
      var2_lab <- NA_character_
    }

    # Nombre y objetivo
    nom   <- nombre_regla_consistencia_g(var, ifelse(is.na(var2), "NA", var2))
    cond_h <- if (nz1(Gh) && nz1(rel_h)) paste0("(", Gh, ") & (", rel_h, ")")
    else if (nz1(Gh))          paste0("(", Gh, ")")
    else if (nz1(rel_h))       paste0("(", rel_h, ")")
    else                       ""

    # Tip human-readable si fuera cardinalidad
    tip_card <- ""
    m_card <- regexec("count\\s*-\\s*selected\\s*\\(\\s*(?:\\.|[A-Za-z0-9_]+)\\s*\\)\\s*([<>]=?|==)\\s*(\\d+)",
                      constraint_raw, perl = TRUE)
    mr <- regmatches(constraint_raw, m_card)[[1]]
    if (length(mr) == 3) {
      op <- mr[2]; k <- mr[3]
      tip_card <- switch(op,
                         "<=" = paste0(" (máximo ", k, " selecciones)"),
                         "==" = paste0(" (exactamente ", k, " selección", if (k != "1") "es" else "", ")"),
                         ">=" = paste0(" (al menos ", k, " selecciones)"),
                         ">"  = paste0(" (más de ", k, " selecciones invalida)"),
                         "<"  = paste0(" (menos de ", k, " selecciones)"),
                         ""
      )
    }

    objetivo <- if (nz1(cond_h)) {
      paste0("En caso la persona cumpla: ", cond_h, ", «", lab, "» respeta la regla del formulario", tip_card, ".")
    } else {
      paste0("«", lab, "» respeta la regla del formulario", tip_card, ".")
    }

    # Procesamiento: !(COND) | (RHS)
    proc <- if (nz1(cond_expr)) {
      paste0(nom, " <- ( !(", cond_expr, ") | (", rhs_r, ") )")
    } else {
      paste0(nom, " <- (", rhs_r, ")")
    }

    tibble::tibble(
      ID = NA_character_,
      `Tipo de observación`      = "Consistencia",
      Objetivo                   = objetivo,
      `Variable 1`               = var,
      `Variable 1 - Etiqueta`    = lab,
      `Variable 2`               = var2,
      `Variable 2 - Etiqueta`    = var2_lab,
      `Variable 3`               = NA_character_,
      `Variable 3 - Etiqueta`    = NA_character_,
      `Nombre de regla`          = nom,
      `Procesamiento`            = proc,
      .pref = pref,
      .gord = section_map$.gord[match(row$group_name, section_map$group_name)],
      .qord = row$.qord
    )
  })
}

# ---- CALCULATE — G-aware, variable gatilladora real --------------------------

# Traduce expresiones de calculation ODK → R
calculate_a_r <- function(txt) {
  if (is.null(txt) || !nzchar(trimws(txt))) return(NA_character_)
  x <- as.character(txt)

  # ${var} → as.numeric(var) (por defecto tratamos como numérico)
  x <- gsub("\\$\\{([A-Za-z0-9_]+)\\}", "as.numeric(\\1)", x, perl = TRUE)

  # Operadores lógicos de ODK
  x <- gsub("(?i)\\band\\b", "&", x, perl = TRUE)
  x <- gsub("(?i)\\bor\\b",  "|", x, perl = TRUE)
  x <- gsub("(?i)\\bnot\\b", "!", x,  perl = TRUE)

  # if( … , … , … ) → ifelse( … , … , … )
  x <- gsub("(?i)\\bif\\s*\\(", "ifelse(", x, perl = TRUE)

  # Limpieza
  x <- gsub("\\s+", " ", trimws(x))
  x
}

# Limpia y “normaliza” el texto de Procesamiento
normalizar_proc <- function(x) {
  if (is.null(x)) return(x)
  x <- as.character(x)

  # Comillas tipográficas → comillas ASCII
  x <- gsub("\u201C|\u201D", "\"", x, perl = TRUE)
  x <- gsub("\u2018|\u2019", "'",  x, perl = TRUE)

  # "=" suelto → "==" (sin tocar <=, >=, !=)
  x <- gsub("(?<!<|>|!|<-|=)=(?!=)", "==", x, perl = TRUE)
  # Si alguien dejó "===" o más, redúcelo a "=="
  x <- gsub("={3,}", "==", x, perl = TRUE)

  # Espacios raros (no-break) → espacio normal
  x <- gsub("[\u00A0\u2007\u202F]", " ", x, perl = TRUE)

  # Balanceo de paréntesis por si quedaron desparejos
  n_open  <- stringr::str_count(x, "\\(")
  n_close <- stringr::str_count(x, "\\)")
  need <- n_open > n_close
  x[need] <- paste0(x[need], strrep(")", n_open[need] - n_close[need]))

  x
}


build_calculate_g <- function(survey, section_map, label_col, gmap){
  dat <- survey %>%
    dplyr::mutate(type_base = tolower(trimws(sub("\\s.*$", "", .data$type)))) %>%
    dplyr::filter(!is.na(.data$name),
                  type_base == "calculate",
                  !is.na(.data$calculation) & nzchar(.data$calculation))

  if (!nrow(dat)) return(tibble::tibble())

  purrr::pmap_dfr(dat, function(...){
    row   <- list(...)
    var   <- row$name
    lab   <- etq_pregunta(survey, label_col, var)
    pref  <- prefijo_para(row$group_name, section_map)
    txt   <- row$calculation %||% ""

    # relevant de la pregunta (rara vez en calculate, pero lo soportamos)
    rel_parsed <- tryCatch(
      relevant_a_r_y_humano(as_chr1(row$relevant), survey, NULL, label_col),
      error = function(e) list(expr_r = "", human = "", vars = character(0))
    )
    rel_expr <- as_chr1(rel_parsed$expr_r)
    rel_h    <- as_chr1(rel_parsed$human)
    rel_vars <- rel_parsed$vars %||% character(0)

    # G del grupo
    ginfo  <- gmap[gmap$group_name == row$group_name, , drop = FALSE]
    G      <- as_chr1(if (nrow(ginfo)) ginfo$G_expr[[1]]   else "")
    Gh     <- as_chr1(if (nrow(ginfo)) ginfo$G_humano[[1]] else "")
    G_vars <- if (nrow(ginfo)) (ginfo$G_vars[[1]] %||% character(0)) else character(0)

    # Condición efectiva
    cond_expr <- if (nz1(G) && nz1(rel_expr)) paste0("(", G, ") & (", rel_expr, ")")
    else if (nz1(G))              G
    else if (nz1(rel_expr))       rel_expr
    else                          ""
    cond_h <- if (nz1(Gh) && nz1(rel_h)) paste0("(", Gh, ") & (", rel_h, ")")
    else if (nz1(Gh))          paste0("(", Gh, ")")
    else if (nz1(rel_h))       paste0("(", rel_h, ")")
    else                       ""

    # Variable 2 (gatilladora)
    drivers <- unique(c(G_vars, rel_vars))
    if (length(drivers) >= 1) {
      var2     <- drivers[1]
      var2_lab <- etq_pregunta(survey, label_col, var2)
    } else {
      var2     <- NA_character_
      var2_lab <- NA_character_
    }

    # Detección de pulldata: la regla es "no debe estar vacío" bajo COND
    tiene_pulldata <- grepl("\\bpulldata\\s*\\(", txt, ignore.case = TRUE)

    if (tiene_pulldata) {
      nom      <- nombre_regla_consistencia_g(var, ifelse(is.na(var2), "NA", var2))
      objetivo <- if (nz1(cond_h)) {
        paste0("En caso la persona cumpla: ", cond_h, ", la variable calculada «", lab, "» no debe estar vacía.")
      } else {
        paste0("La variable calculada «", lab, "» no debe estar vacía.")
      }
      proc <- if (nz1(cond_expr)) {
        paste0(nom, " <- ( !(", cond_expr, ") | ", expr_no_vacio(var), " )")
      } else {
        paste0(nom, " <- ", expr_no_vacio(var))
      }

      return(tibble::tibble(
        ID = NA_character_,
        `Tipo de observación`   = "Consistencia",
        Objetivo                = objetivo,
        `Variable 1`            = var,
        `Variable 1 - Etiqueta` = lab,
        `Variable 2`            = var2,
        `Variable 2 - Etiqueta` = var2_lab,
        `Variable 3`            = NA_character_,
        `Variable 3 - Etiqueta` = NA_character_,
        `Nombre de regla`       = nom,
        `Procesamiento`         = proc,
        .pref = pref,
        .gord = section_map$.gord[match(row$group_name, section_map$group_name)],
        .qord = row$.qord
      ))
    }

    # Caso general: comparamos con el cálculo (asumiendo numérico por defecto)
    rhs_r <- calculate_a_r(txt)  # tu helper que convierte a expresión R
    if (!nz1(rhs_r)) return(tibble::tibble())

    nom      <- nombre_regla_consistencia_g(var, ifelse(is.na(var2), "NA", var2))
    objetivo <- if (nz1(cond_h)) {
      paste0("En caso la persona cumpla: ", cond_h, ", «", lab, "» coincide con el cálculo.")
    } else {
      paste0("El valor de «", lab, "» coincide con el cálculo definido.")
    }
    comp <- paste0("(suppressWarnings(as.numeric(", var, ")) == (", rhs_r, "))")
    proc <- if (nz1(cond_expr)) {
      paste0(nom, " <- ( !(", cond_expr, ") | ", comp, " )")
    } else {
      paste0(nom, " <- ", comp)
    }

    tibble::tibble(
      ID = NA_character_,
      `Tipo de observación`   = "Valores atípicos",
      Objetivo                = objetivo,
      `Variable 1`            = var,
      `Variable 1 - Etiqueta` = lab,
      `Variable 2`            = var2,
      `Variable 2 - Etiqueta` = var2_lab,
      `Variable 3`            = NA_character_,
      `Variable 3 - Etiqueta` = NA_character_,
      `Nombre de regla`       = nom,
      `Procesamiento`         = proc,
      .pref = pref,
      .gord = section_map$.gord[match(row$group_name, section_map$group_name)],
      .qord = row$.qord
    )
  })
}

# ---- FECHA CAMPO -------------------------------------------------------------

build_fecha_campo <- function(survey, section_map, label_col,
                              fecha_var = "mand_Date",
                              fecha_campo = Sys.Date()) {
  if (!fecha_var %in% survey$name) return(tibble())
  row <- survey[survey$name == fecha_var, , drop = FALSE]
  lab <- etq_pregunta(survey, label_col, fecha_var)
  pref <- prefijo_para(row$group_name, section_map)
  limite <- as.character(as.Date(fecha_campo))
  nombre <- nombre_regla_simple(fecha_var)
  objetivo <- paste0("La fecha «", lab, "» no debe ser posterior a ", limite, ".")
  proc <- paste0(
    nombre, " <- ( as.Date(suppressWarnings(as.character(", fecha_var, "))) <= as.Date(\"", limite, "\") )"
  )
  tibble(
    ID = NA_character_,
    `Tipo de observación` = "7. Consistencia",
    Objetivo = objetivo,
    `Variable 1` = fecha_var,
    `Variable 1 - Etiqueta` = lab,
    `Variable 2` = NA_character_,
    `Variable 2 - Etiqueta` = NA_character_,
    `Variable 3` = NA_character_,
    `Variable 3 - Etiqueta` = NA_character_,
    `Nombre de regla` = nombre,
    `Procesamiento` = proc,
    .pref = pref,
    .gord = section_map$.gord[match(as_chr1(row$group_name), section_map$group_name)],
    .qord = as.numeric(row$.qord %||% NA_real_)
  )
}


# ---- CHOICE FILTER (select_one) — G-aware, una fila por driver con redacción clara ----
build_choice_filter_g <- function(x){
  stopifnot(is.list(x), all(c("survey","choices","meta") %in% names(x)))

  survey      <- x$survey
  choices     <- x$choices %||% tibble::tibble()
  section_map <- x$meta$section_map %||% tibble::tibble()
  ccols_map   <- x$meta$choice_cols_by_list %||% tibble::tibble(list_name=character(), extra_cols=list())
  label_col   <- x$meta$label_col_survey %||% "label"
  if (!nrow(survey)) return(tibble::tibble())

  gmap <- tryCatch(.make_gmap(x), error = function(e) tibble::tibble())

  dat <- survey %>%
    dplyr::mutate(
      type_base     = tolower(trimws(sub("\\s.*$", "", .data$type))),
      choice_filter = ifelse(is.na(.data$choice_filter), "", .data$choice_filter)
    ) %>%
    dplyr::filter(type_base == "select_one", nzchar(choice_filter), nzchar(list_name))

  if (!nrow(dat)) return(tibble::tibble())

  purrr::pmap_dfr(dat, function(...){
    row <- list(...)
    var      <- row$name
    lab      <- etq_pregunta(survey, label_col, var)
    pref     <- prefijo_para(row$group_name, section_map)
    ln       <- row$list_name

    # Condición de aplicación (G del grupo + relevant del ítem)
    ginfo   <- gmap[gmap$group_name == row$group_name, , drop = FALSE]
    G_r     <- as_chr1(if (nrow(ginfo)) ginfo$G_expr[[1]]   else "")
    rel_r   <- as_chr1(tryCatch(relevant_a_r_y_humano(as_chr1(row$relevant %||% ""), survey, choices, label_col)$expr_r,
                                error = function(e) ""))
    cond_r  <- if (nz1(G_r) && nz1(rel_r)) paste0("(", G_r, ") & (", rel_r, ")")
    else if (nz1(G_r))          G_r
    else if (nz1(rel_r))        rel_r
    else                        ""

    # Drivers disponibles (columnas filter_* en choices)
    extras <- ccols_map$extra_cols[[ match(ln, ccols_map$list_name) ]] %||% character(0)
    cols_driver <- extras[grepl("^filter_", extras)]
    if (!length(cols_driver)) return(tibble::tibble())

    # Choices de la lista
    ch_ln <- choices[choices$list_name == ln, , drop = FALSE]
    if (!nrow(ch_ln)) return(tibble::tibble())

    purrr::map_dfr(cols_driver, function(colf){
      drv <- sub("^filter_", "", colf)
      if (!drv %in% survey$name) return(tibble::tibble())
      drv_lab <- etq_pregunta(survey, label_col, drv)

      # Condiciones por opción según ese driver
      terms_drv <- ch_ln %>%
        dplyr::group_by(name) %>%
        dplyr::summarise(
          cond = {
            vals <- unique(na.omit(as.character(cur_data_all()[[colf]])))
            vals <- vals[nzchar(vals)]
            if (!length(vals)) "" else {
              vals_q <- paste0("'", gsub("'", "\\\\'", vals), "'")
              paste0("as.character(", drv, ") %in% c(", paste(vals_q, collapse=","), ")")
            }
          },
          .groups = "drop"
        ) %>%
        dplyr::filter(nzchar(cond)) %>%
        dplyr::mutate(
          expr = paste0("((", var, " == '", gsub("'", "\\\\'", name), "') & (", cond, "))")
        ) %>%
        dplyr::pull(expr)

      if (!length(terms_drv)) return(tibble::tibble())

      allowed_expr_drv <- .or_join(terms_drv)

      drv_presente <- paste0("(!is.na(", drv, ") & trimws(", drv, ") != \"\")")
      base_ok_drv  <- paste0("(is.na(", var, ") | trimws(", var, ") == \"\") | (", allowed_expr_drv, ")")

      nombre_drv <- paste0("cruce_", var, "_cf_", drv)
      proc_drv <- if (nz1(cond_r)) {
        paste0(nombre_drv, " <- ( !(", cond_r, ") | !(", drv_presente, ") | (", base_ok_drv, ") )")
      } else {
        paste0(nombre_drv, " <- ( !(", drv_presente, ") | (", base_ok_drv, ") )")
      }

      # Valores posibles para descripción
      vals_presentes <- ch_ln[[colf]] |> as.character() |> unique() |> na.omit()
      vals_presentes <- vals_presentes[nzchar(vals_presentes)]
      vals_human <- if (length(vals_presentes)) paste0("{", paste(vals_presentes, collapse = ", "), "}") else "—"

      # ---------------------- OBJETIVO ----------------------
      objetivo_drv <- paste0(
        "Valida la coherencia entre «", lab, "» y «", drv_lab, "». ",
        "Si la pregunta aplica y «", drv_lab, "» tiene una respuesta válida, ",
        "la opción marcada en «", lab, "» debe encontrarse entre las opciones que el formulario muestra ",
        "cuando «", drv_lab, "» toma los valores ", vals_human, ". ",
        "En caso contrario, la respuesta en «", lab, "» se considera inconsistente."
      )

      tibble::tibble(
        ID = NA_character_,
        `Tipo de observación`      = "Filtro de opciones",
        Objetivo                   = objetivo_drv,
        `Variable 1`               = var,
        `Variable 1 - Etiqueta`    = lab,
        `Variable 2`               = drv,
        `Variable 2 - Etiqueta`    = drv_lab,
        `Variable 3`               = NA_character_,
        `Variable 3 - Etiqueta`    = NA_character_,
        `Nombre de regla`          = nombre_drv,
        `Procesamiento`            = proc_drv,
        .pref = pref,
        .gord = section_map$.gord[match(row$group_name, section_map$group_name)],
        .qord = row$.qord
      )
    })
  })
}



# =============================================================================
# Roxygen (export principal)
# =============================================================================

#' Generar plan de limpieza (G-aware) desde un XLSForm ya leído
#'
#' @param x Lista devuelta por `leer_xlsform_limpieza()` con `$survey`, `$choices`, `$meta`.
#' @param incluir Lista de banderas lógicas para incluir bloques.
#' @param fecha_var Nombre de la variable de fecha a verificar (default `"mand_Date"`).
#' @param fecha_campo Fecha límite para `fecha_var` (Date o `"YYYY-MM-DD"`).
#' @return tibble con plan.
#' @export
generar_plan_limpieza <- function(
    x,
    incluir = list(
      required      = TRUE,
      other         = TRUE,
      relevant      = TRUE,
      constraint    = TRUE,
      calculate     = TRUE,
      choice_filter = TRUE,
      fecha_campo   = TRUE
    ),
    fecha_var   = "mand_Date",
    fecha_campo = Sys.Date()
){
  stopifnot(all(c("survey","meta") %in% names(x)))
  x <- coalesce_choice_filter(x)
  safe <- .columna_label_segura(x)
  labelcol <- safe$labelcol
  survey   <- safe$survey

  # Coerciones suaves / auxiliares
  survey <- survey %>%
    mutate(
      across(any_of(c("type","name","relevant","constraint","calculation","group_name","required")),
             ~{ if (is.function(.x)) NA_character_ else suppressWarnings(as.character(.x)) }),
      .qord     = row_number(),
      type_base = tolower(trimws(sub("\\s.*$", "", .data$type))),
      required  = {
        s <- tolower(trimws(as.character(.data$required)))
        s <- ifelse(is.na(s), "", s)
        s <- iconv(s, from = "", to = "ASCII//TRANSLIT")
        s %in% c("true()", "true", "1", "yes", "si", "y", "s")
      },
      relevant = gsub("\\s+", " ", trimws(coalesce(.data$relevant, "")))
    )
  survey <- .ensure_group_names(survey)
  section_map <- x$meta$section_map %>% mutate(.gord = row_number())
  choices     <- x$choices %||% tibble()

  # Mapa de G (relevant del begin_group) — SOLO dentro de la función
  gmap <- .make_gmap(x)

  # ----- construir bloques -----
  bloques <- list()
  if (isTRUE(incluir$required))      bloques$required      <- build_required_g(survey, section_map, labelcol, gmap)
  if (isTRUE(incluir$other))         bloques$other         <- build_other_g(survey, section_map, labelcol, gmap)
  if (isTRUE(incluir$relevant))      bloques$relevant      <- build_relevant_g(survey, section_map, labelcol, choices, gmap)
  if (isTRUE(incluir$constraint))    bloques$constraint    <- build_constraint_g(survey, section_map, labelcol, gmap)
  if (isTRUE(incluir$calculate))     bloques$calculate     <- build_calculate_g(survey, section_map, labelcol, gmap)
  if (isTRUE(incluir$choice_filter)) bloques$choicefilter  <- build_choice_filter_g(x)
  if (isTRUE(incluir$fecha_campo))   bloques$fecha         <- build_fecha_campo(survey, section_map, labelcol, fecha_var, fecha_campo)
  plan <- bind_rows(bloques)

  if (nrow(plan) == 0) {
    return(tibble(
      ID = character(),
      `Tipo de observación` = character(),
      Objetivo = character(),
      `Variable 1` = character(),
      `Variable 1 - Etiqueta` = character(),
      `Variable 2` = character(),
      `Variable 2 - Etiqueta` = character(),
      `Variable 3` = character(),
      `Variable 3 - Etiqueta` = character(),
      `Nombre de regla` = character(),
      `Procesamiento` = character()
    ))
  }

  # Orden/IDs
  plan <- plan %>%
    arrange(.gord, .qord) %>%
    group_by(.pref) %>%
    mutate(.idx = row_number()) %>%
    ungroup() %>%
    mutate(
      ID            = paste0(.pref, sprintf("%03d", .idx)),
      Procesamiento = normalizar_proc(Procesamiento)
    ) %>%
    select(-.idx, - .pref, - .gord, - .qord)

  plan
}
