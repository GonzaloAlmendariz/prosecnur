#' PPRA – Flujo híbrido con Excel de especificación de familias
#'
#' Helpers y flujo para construir la plantilla de codificación PPRA
#' desde un XLSForm (survey/choices), datos crudos y una especificación
#' de familias en Excel.
#'
#' @keywords internal
#' @name ppra_flujo_hibrido
NULL
suppressPackageStartupMessages({library(tidyverse); library(openxlsx); library(janitor)})

`%||%` <- function(x,y) if (is.null(x) || (length(x)==1 && is.na(x))) y else x
nzchr   <- function(x) is.character(x) && length(x)==1 && !is.na(x) && nzchar(x)

# -- utilidades ---------------------------------------------------------------
norm_list_name <- function(x){
  x <- tolower(trimws(as.character(x)))
  x <- gsub("\\s+", "_", x); gsub("[^a-z0-9_]", "_", x)
}
detect_label_col_robust <- function(nms){
  nms_l <- tolower(nms)
  exact <- which(nms_l %in% c("label::spanish (es)","label::spanish(es)","label::spanish","label::es","label::spanish_es"))
  if (length(exact)) return(nms[exact[1]])
  cand  <- which(grepl("^label($|[:_])", nms_l) | grepl("^label", nms_l))
  if (!length(cand)) return(NA_character_)
  nm_sub <- nms_l[cand]
  score <- 6*grepl("label::span|label::es", nm_sub) +
    4*grepl("span|españ|\\bes\\b|_es$|esp|cast", nm_sub) +
    2*grepl("^label($|[:_])", nm_sub) + 1*grepl("^label", nm_sub)
  cand[order(-score)][1] |> (\(i) nms[i])()
}
get_q_label_strict <- function(var, survey_clean, survey_raw){
  if (!nzchr(var)) return(NA_character_)
  var_clean <- janitor::make_clean_names(var)
  lc <- detect_label_col_robust(names(survey_clean)) %||% "label"
  if (!is.na(lc) && lc %in% names(survey_clean) && var_clean %in% survey_clean$name){
    v <- survey_clean[[lc]][match(var_clean, survey_clean$name)]
    if (!is.na(v) && nzchar(trimws(v))) return(v)
  }
  if (!is.null(survey_raw) && "name" %in% names(survey_raw)){
    idx <- which(janitor::make_clean_names(survey_raw$name) == var_clean)
    if (length(idx)){
      lc2 <- detect_label_col_robust(names(survey_raw))
      if (!is.na(lc2) && lc2 %in% names(survey_raw)){
        v <- survey_raw[[lc2]][idx]
        if (!is.na(v) && nzchar(trimws(v))) return(v)
      }
    }
  }
  NA_character_
}
# texto real (no binario)
is_textlike <- function(v){
  if (is.null(v)) return(FALSE)
  x <- as.character(v); x <- x[!is.na(x)]
  if (!length(x)) return(FALSE)
  if (all(x %in% c("0","1","TRUE","FALSE","T","F"))) return(FALSE)
  max(nchar(x), na.rm = TRUE) > 1
}


# Busca la primera columna de "label en español" (case-insensitive) en un data.frame
.detect_label_es_col <- function(df) {
  if (is.null(df) || !ncol(df)) return(NA_character_)
  nms <- names(df); nmsl <- tolower(nms)
  # orden de preferencia
  pats <- c(
    "^label::spanish \\(es\\)$",
    "^label::spanish\\(es\\)$",
    "^label::spanish$",
    "^label::es$",
    "^label_es$",
    "^label_spanish_es$",
    "^label\\b"         # último fallback (si solo hay 'label' y sí está en español)
  )
  for (p in pats) {
    hit <- which(grepl(p, nmsl, perl = TRUE))
    if (length(hit)) return(nms[hit[1]])
  }
  NA_character_
}

# Etiqueta en ES por nombre de variable (survey)
s_lab_from_original <- function(orig, inst){
  if (is.null(orig)) return(NA_character_)
  sr <- inst$survey_raw
  s  <- inst$survey
  col_sr <- .detect_label_es_col(sr)
  cand_s <- intersect(c("label_spanish_es","label_es","label"), names(s))

  vapply(as.list(orig), function(o){
    o <- as.character(o)[1]; if (is.na(o) || !nzchar(o)) return(NA_character_)
    nm <- janitor::make_clean_names(o)

    # 1) survey_raw (preferente)
    if (!is.null(sr) && !is.na(col_sr)) {
      i <- match(nm, janitor::make_clean_names(sr$name))
      if (!is.na(i)) {
        val <- as.character(sr[[col_sr]][i])
        if (length(val)==1 && !is.na(val) && nzchar(trimws(val))) return(val)
      }
    }
    # 2) survey (labels)
    if (length(cand_s)) {
      i2 <- match(nm, janitor::make_clean_names(s$name))
      if (!is.na(i2)) {
        for (cname in cand_s) {
          if (cname %in% names(s)) {
            val <- as.character(s[[cname]][i2])
            if (length(val)==1 && !is.na(val) && nzchar(trimws(val))) return(val)
          }
        }
      }
    }
    # 3) fallback: nombre limpio
    nm
  }, FUN.VALUE = character(1))
}



#' Resolver nombre limpio de parent (desde split)
#'
#' Toma `parent` si existe; si no, `parent_col`. Devuelve el equivalente
#' limpio (coincide con `survey$name`).
#'
#' @param df data.frame con `parent` y/o `parent_col`.
#' @return character con nombres limpios.
#' @export
resolver_parent_limpio <- function(df){
  pc <- if ("parent" %in% names(df) && any(!is.na(df$parent))) df$parent else df$parent_col
  janitor::make_clean_names(pc)
}

# Alias interno para compatibilidad con código previo
#' @keywords internal
#' @noRd
resolve_parent_clean <- function(df) resolver_parent_limpio(df)



#' Detectar la mejor columna de label en español (con fallback a "label")
#'
#' Busca primero variantes de español (label::Spanish (ES), etc.). Si no hay,
#' devuelve "label" si existe; si tampoco, NA.
#' @param nms character vector con nombres de columnas
#' @return nombre de columna o NA_character_
#' @export
detect_spanish_col <- function(nms){
  if (is.null(nms) || !length(nms)) return(NA_character_)
  nmsl <- tolower(nms)

  # preferencia ES
  exact <- c(
    "label::spanish (es)", "label::spanish(es)", "label::spanish_es",
    "label_spanish_es", "label::spanish", "label::es"
  )
  hit <- which(nmsl %in% exact)
  if (length(hit)) return(nms[hit[1]])

  # patrón laxo por si hay variantes raras
  hit2 <- which(grepl("^label(::)?span|label[_:]spanish|label[_:]es$", nmsl, perl = TRUE))
  if (length(hit2)) return(nms[hit2[1]])

  # fallback: usar "label" si existe
  if ("label" %in% nmsl) return(nms[which(nmsl == "label")[1]])

  NA_character_
}

#' Label ES desde el instrumento (vectorizado; prioriza survey_raw)
#' @export
s_lab_from_original <- function(orig, inst){
  if (is.null(orig)) return(NA_character_)
  v <- as.character(orig)
  nm_clean <- janitor::make_clean_names(v)

  # 1) survey_raw
  sr <- inst$survey_raw
  if (!is.null(sr) && "name" %in% names(sr)) {
    i <- match(nm_clean, janitor::make_clean_names(sr$name))
    col_es <- detect_spanish_col(names(sr))
    if (!is.na(col_es)) {
      out <- as.character(sr[[col_es]][i])
    } else {
      out <- rep(NA_character_, length(nm_clean))
    }
  } else {
    out <- rep(NA_character_, length(nm_clean))
  }

  # 2) fallback a survey (label_spanish_es / label_es / label)
  s <- inst$survey %||% inst$survey_raw
  if (!is.null(s) && "name" %in% names(s)) {
    i2 <- match(nm_clean, janitor::make_clean_names(s$name))
    col2 <- detect_spanish_col(names(s))
    add  <- if (!is.na(col2)) as.character(s[[col2]][i2]) else NA_character_
    out  <- ifelse(!nzchar(out) | is.na(out), add, out)
  }

  # 3) último recurso: nombre limpio
  out[!nzchar(out) | is.na(out)] <- nm_clean[!nzchar(out) | is.na(out)]
  out
}

#' Tabla de choices con label ES robusto (independiente del formato)
#' @export
choices_es_tbl <- function(inst){
  ch <- inst$choices_raw %||% inst$choices
  if (is.null(ch)) {
    return(tibble::tibble(list_name=character(), list_norm=character(),
                          name=character(), label_es=character()))
  }

  # columna de label (ES si hay; si no, "label"; si no, name)
  col_es <- detect_spanish_col(names(ch))
  lbl    <- if (!is.na(col_es)) ch[[col_es]] else ch$name

  # asegurar list_norm
  if (!"list_norm" %in% names(ch)) {
    ln <- if ("list_name" %in% names(ch)) ch$list_name else NA_character_
    list_norm <- tolower(gsub("[^a-z0-9_]", "_", gsub("\\s+","_", as.character(ln))))
    ch$list_norm <- list_norm
  }

  tibble::tibble(
    list_name = ch$list_name %||% NA_character_,
    list_norm = ch$list_norm %||% NA_character_,
    code      = as.character(ch$name),
    label_es  = as.character(lbl)
  )
}


#' Normalizar etiquetas y listas en el instrumento (survey/choices)
#'
#' Asegura columnas estandarizadas para trabajar a gusto:
#' \itemize{
#'   \item En \code{survey}: crea/rellena \code{label_es} a partir de las columnas
#'         de label disponibles (preferencia ES > EN > label genérica > name).
#'         También garantiza \code{list_norm} (normalización de \code{list_name}).
#'   \item En \code{choices}: crea/rellena \code{choice_label} (preferencia ES > EN > label > name)
#'         y garantiza \code{list_norm}.
#' }
#'
#' La función tolera entradas con \code{survey}/\code{choices} limpios o
#' \code{survey_raw}/\code{choices_raw} (en cuyo caso limpia nombres primero).
#'
#' @param inst Lista con, idealmente, \code{$survey} y \code{$choices}. Si no existen,
#'   intenta usar \code{$survey_raw} y \code{$choices_raw}.
#'
#' @return La misma lista \code{inst}, con \code{$survey} y \code{$choices}
#'   normalizados (columnas \code{label_es}, \code{choice_label}, \code{list_norm} presentes).
#' @export
#'
#' @examples
#' \dontrun{
#' inst <- leer_instrumento_xlsform("instrumento_pdm.xlsx")
#' inst <- normalizar_labels_inst(inst)
#' }
normalizar_labels_inst <- function(inst){
  stopifnot(is.list(inst))


  #' Resolver nombre limpio de parent desde split
  #'
  #' Función auxiliar para unificar la clave de las preguntas.
  #' Toma la columna `parent` si existe (y no está vacía); de lo contrario usa `parent_col`.
  #' Ambos se normalizan con \code{janitor::make_clean_names()} para que coincidan con `survey$name`.
  #'
  #' @param df Data frame que proviene de \code{split$select_one}, \code{split$select_multiple} o \code{split$text}.
  #' @return Un vector de nombres "clean" (caracter).
  #' @keywords internal
  #' @export
  resolver_parent_limpio <- function(df) {
    pc <- if ("parent" %in% names(df) && any(!is.na(df$parent))) {
      df$parent
    } else {
      df$parent_col
    }
    janitor::make_clean_names(pc)
  }

  # -- helpers locales --------------------------------------------------------
  `%||%` <- function(x, y) if (is.null(x) || (length(x) == 1 && is.na(x))) y else x

  norm_list_name <- function(x){
    x <- tolower(trimws(as.character(x)))
    x <- gsub("\\s+", "_", x)
    gsub("[^a-z0-9_]", "_", x)
  }

  detect_label_col_robust <- function(nms){
    nms_l <- tolower(nms)
    exact <- which(nms_l %in% c("label::spanish (es)","label::spanish(es)","label::spanish",
                                "label::es","label_spanish_es","label::spanish_es"))
    if (length(exact)) return(nms[exact[1]])
    cand  <- which(grepl("^label($|[:_])", nms_l) | grepl("^label", nms_l))
    if (!length(cand)) return(NA_character_)
    nms[cand[1]]
  }

  coalesce_first <- function(df, out_col, candidates, fallback){
    if (!out_col %in% names(df)) df[[out_col]] <- NA_character_
    for (c in candidates) if (c %in% names(df)) {
      df[[out_col]] <- dplyr::coalesce(df[[out_col]], as.character(df[[c]]))
    }
    df[[out_col]] <- dplyr::coalesce(df[[out_col]], as.character(df[[fallback]]))
    df
  }

  # -- resolver survey/choices base ------------------------------------------
  survey  <- inst$survey  %||% (inst$survey_raw  %||% NULL)
  choices <- inst$choices %||% (inst$choices_raw %||% NULL)

  if (is.null(survey) || is.null(choices)) {
    abort("`inst` debe incluir `survey` y `choices` (o `survey_raw`/`choices_raw`).")
  }

  # limpiar nombres si vienen crudos
  if (!"name" %in% names(survey))  survey  <- janitor::clean_names(survey)
  if (!"name" %in% names(choices)) choices <- janitor::clean_names(choices)

  # -- SURVEY -----------------------------------------------------------------
  # detectar label en survey (ES si existe, si no el mejor disponible)
  lab_s <- detect_label_col_robust(names(survey)) %||% (if ("label" %in% names(survey)) "label" else NA_character_)
  if (!"label_es" %in% names(survey)) survey$label_es <- NA_character_
  if (!is.na(lab_s) && lab_s %in% names(survey)) {
    survey$label_es <- dplyr::coalesce(survey$label_es, as.character(survey[[lab_s]]))
  }
  # fallback: inglés -> label genérica -> name
  for (alt in c("label_english_en","label")) {
    if (alt %in% names(survey)) {
      survey$label_es <- dplyr::coalesce(survey$label_es, as.character(survey[[alt]]))
    }
  }
  survey$label_es <- dplyr::coalesce(survey$label_es, as.character(survey$name))

  # list_name y list_norm desde type (si faltan)
  if (!"list_name" %in% names(survey)) {
    if (!"type" %in% names(survey)) survey$type <- NA_character_
    survey$list_name <- trimws(sub("^\\S+\\s+","", survey$type))
  }
  if (!"list_norm" %in% names(survey)) {
    survey$list_norm <- norm_list_name(survey$list_name)
  } else {
    survey$list_norm <- dplyr::coalesce(survey$list_norm, norm_list_name(survey$list_name))
  }

  # -- CHOICES ----------------------------------------------------------------
  # asegurar choice_label con preferencia ES -> EN -> label -> name
  # y completar list_norm
  if (!"list_name" %in% names(choices)) {
    abort("La hoja `choices` no tiene `list_name`. Renómbrala o provee un mapeo previo.")
  }

  # detectar label en choices
  lab_c <- detect_label_col_robust(names(choices)) %||% (if ("label" %in% names(choices)) "label" else NA_character_)
  if (!"choice_label" %in% names(choices)) choices$choice_label <- NA_character_
  if (!is.na(lab_c) && lab_c %in% names(choices)) {
    choices$choice_label <- dplyr::coalesce(choices$choice_label, as.character(choices[[lab_c]]))
  }
  for (alt in c("label_english_en","label")) {
    if (alt %in% names(choices)) {
      choices$choice_label <- dplyr::coalesce(choices$choice_label, as.character(choices[[alt]]))
    }
  }
  choices$choice_label <- dplyr::coalesce(choices$choice_label, as.character(choices$name))

  # list_norm
  if (!"list_norm" %in% names(choices)) {
    choices$list_norm <- norm_list_name(choices$list_name)
  } else {
    choices$list_norm <- dplyr::coalesce(choices$list_norm, norm_list_name(choices$list_name))
  }

  # homogenizar por si acaso
  choices$list_norm <- norm_list_name(choices$list_norm)
  survey$list_norm  <- norm_list_name(survey$list_norm)

  # -- devolver en inst -------------------------------------------------------
  inst$survey  <- survey
  inst$choices <- choices
  inst
}


# --- AUDITORÍA Y ENRIQUECIMIENTO -------------------------------------------

auditar_split <- function(split){
  req_common <- c("parent_col")
  req_one    <- c(req_common, "text_col")
  req_mult   <- c(req_common, "text_col", "other_dummy_col")
  req_text   <- c("parent_col")

  faltan <- function(df, req){
    setdiff(req, names(df))
  }

  cat("\n[AUDITORÍA split]\n")
  if (!is.null(split$select_one)) {
    f <- faltan(split$select_one, req_one)
    cat("select_one   -> faltan:", if (length(f)) paste(f, collapse=", ") else "ok", "\n")
  } else cat("select_one   -> objeto ausente\n")

  if (!is.null(split$select_multiple)) {
    f <- faltan(split$select_multiple, req_mult)
    cat("select_multiple -> faltan:", if (length(f)) paste(f, collapse=", ") else "ok", "\n")
  } else cat("select_multiple -> objeto ausente\n")

  if (!is.null(split$text)) {
    f <- faltan(split$text, req_text)
    cat("text         -> faltan:", if (length(f)) paste(f, collapse=", ") else "ok", "\n")
  } else cat("text         -> objeto ausente\n")
  invisible(split)
}

enriquecer_split_con_survey <- function(split, inst){
  survey <- inst$survey
  stopifnot(!is.null(survey), "name" %in% names(survey))

  # helper: a partir de parent_col (nombre original),
  # tratar de encontrar la fila del survey (por name ya "clean"):
  find_s_row <- function(parent_col){
    # parent_col viene en "original". Lo limpiamos para buscar en survey$name:
    nm_clean <- janitor::make_clean_names(parent_col)
    which(survey$name == nm_clean)[1] %||% NA_integer_
  }

  enrich_df <- function(df){
    if (is.null(df) || !nrow(df)) return(df)

    # columnas destino para no fallar
    for (k in c("parent","parent_label","q_order","list_name","list_norm","parent_key")) {
      if (!k %in% names(df)) df[[k]] <- NA
    }

    # 1) clave limpia para cruzar (NO tocar `parent` original del Excel)
    #    usamos parent si viene; si no, parent_col; y lo limpiamos -> parent_clean
    src_parent <- if ("parent" %in% names(df) && any(!is.na(df$parent))) df$parent else df$parent_col
    parent_clean <- janitor::make_clean_names(src_parent)

    # 2) match contra survey$name
    survey <- inst$survey
    idx <- match(parent_clean, survey$name)
    ok  <- !is.na(idx)

    # 3) parent_key: guardar SIEMPRE la clave limpia que matchea survey
    df$parent_key[ok] <- survey$name[idx[ok]]

    # 4) NO SOBREESCRIBIR `parent` si ya venía; si está vacío,
    #    preferimos mostrar el original de datos (parent_col) y, si no, la clave limpia
    df$parent <- dplyr::coalesce(df$parent, df$parent_col, df$parent_key)

    # 5) label en español usando survey_raw/survey (helper robusto)
    df$parent_label[ok] <- label_es_from_inst(df$parent[ok], inst)

    # 6) metadatos si existen en survey
    if ("q_order"   %in% names(survey))   df$q_order[ok]   <- survey$q_order[idx[ok]]
    if ("list_name" %in% names(survey))   df$list_name[ok] <- survey$list_name[idx[ok]]
    if ("list_norm" %in% names(survey))   df$list_norm[ok] <- survey$list_norm[idx[ok]]

    df
  }

  split$select_one      <- enrich_df(split$select_one)
  split$select_multiple <- enrich_df(split$select_multiple)
  split$text            <- enrich_df(split$text)
  split
}



# -- 1) Instrumento -----------------------------------------------------------
#' Leer XLSForm sin normalización agresiva (solo auxiliares mínimas)
#'
#' Lee `survey` y `choices` y NO mezcla idiomas ni reescribe labels.
#' Solo añade:
#' - `q_order` (número de fila),
#' - `type_base` (primer token de `type`),
#' - `list_name` (si falta, se extrae de `type`),
#' - `list_norm` (normalización de `list_name` para cruzar con choices),
#' - `label_spanish_es` en ambas hojas si existe alguna variante de columna ES.
#'
#' Nada más se modifica. Las columnas originales quedan intactas.
#'
#' @param path Ruta al XLSForm (.xlsx)
#' @return lista con `survey_raw`, `choices_raw`, `survey`, `choices`
#' @export
leer_instrumento_xlsform <- function(path){
  # --- helpers locales, sin tocar labels originales ---
  .norm_list_name <- function(x){
    x <- tolower(trimws(as.character(x)))
    x <- gsub("\\s+", "_", x)
    gsub("[^a-z0-9_]", "_", x)
  }
  .find_spanish_col <- function(nms){
    nms_l <- tolower(nms)
    # preferimos coincidencias exactas comunes
    exact <- c(
      "label::spanish (es)","label::spanish(es)","label::spanish_es",
      "label_spanish_es","label::spanish","label::es"
    )
    hit <- which(nms_l %in% exact)
    if (length(hit)) return(nms[hit[1]])
    # fallback muy suave (si realmente hiciera falta)
    hit2 <- which(grepl("^label(::)?span|label[_:]es$", nms_l, perl = TRUE))
    if (length(hit2)) return(nms[hit2[1]])
    NA_character_
  }

  # --- leer crudo tal cual ---
  survey_raw  <- suppressWarnings(readxl::read_excel(path, sheet = "survey"))
  choices_raw <- suppressWarnings(readxl::read_excel(path, sheet = "choices"))

  # trabajamos sobre copias (no clean_names para no romper nada)
  survey  <- survey_raw
  choices <- choices_raw

  # columnas base en survey
  if (!"name" %in% names(survey)) stop("La hoja 'survey' debe tener columna 'name'.")
  if (!"type" %in% names(survey)) survey$type <- NA_character_

  # auxiliares mínimas
  survey$q_order   <- seq_len(nrow(survey))
  survey$type_base <- sub("\\s.*$", "", as.character(survey$type %||% ""))
  if (!"list_name" %in% names(survey) || all(is.na(survey$list_name))) {
    survey$list_name <- trimws(sub("^\\S+\\s+","", as.character(survey$type %||% "")))
  }
  survey$list_norm <- .norm_list_name(survey$list_name)

  # alias del label ES si existe alguna variante
  s_es_col <- .find_spanish_col(names(survey))
  survey$label_spanish_es <- if (!is.na(s_es_col)) as.character(survey[[s_es_col]]) else NA_character_

  # columnas base en choices
  if (!"list_name" %in% names(choices)) stop("La hoja 'choices' debe tener 'list_name'.")
  if (!"name" %in% names(choices))     stop("La hoja 'choices' debe tener 'name' (código).")

  choices$list_norm <- .norm_list_name(choices$list_name)
  choices$choice_code <- as.character(choices$name)

  c_es_col <- .find_spanish_col(names(choices))
  choices$label_spanish_es <- if (!is.na(c_es_col)) as.character(choices[[c_es_col]]) else NA_character_

  list(
    survey_raw  = survey_raw,
    choices_raw = choices_raw,
    survey      = survey,
    choices     = choices
  )
}

#' Auditoría mínima (sin normalizar)
#' @param inst lista devuelta por `leer_instrumento_xlsform_min()`
#' @export
auditar_inst_min <- function(inst){
  cat("[inst] elementos:", paste(names(inst), collapse=", "), "\n")
  s <- inst$survey; c <- inst$choices
  cat("\n[survey] filas/cols:", nrow(s), "/", ncol(s), "\n")
  cat("  columnas claves presentes:",
      paste(intersect(c("name","type","q_order","type_base","list_name","list_norm","label_spanish_es"), names(s)), collapse=", "),
      "\n")
  cat("  NAs en label_spanish_es:", sum(is.na(s$label_spanish_es)), "de", nrow(s), "\n")
  cat("\n[choices] filas/cols:", nrow(c), "/", ncol(c), "\n")
  cat("  columnas claves presentes:",
      paste(intersect(c("list_name","list_norm","name","choice_code","label_spanish_es"), names(c)), collapse=", "),
      "\n")
  cat("  NAs en label_spanish_es (choices):", sum(is.na(c$label_spanish_es)), "de", nrow(c), "\n")
  invisible(inst)
}


# -- 2) Datos -----------------------------------------------------------------

#' Leer datos manteniendo nombres originales + mapa clean↔original
#'
#' Lee un archivo de datos (.xlsx/.csv) **sin alterar tipos ni valores** y
#' preservando los nombres de columnas tal como vienen. Además construye:
#' \itemize{
#'   \item \code{$clean}: una copia con nombres normalizados vía
#'         \code{janitor::make_clean_names()} (para cruzar con \code{survey$name}).
#'   \item \code{$name_map}: tibble con el mapeo \code{clean ↔ original}.
#' }
#'
#' @param path Ruta al archivo (xlsx/csv).
#' @param sheet Hoja (si aplica; usada solo si el archivo es Excel).
#'
#' @return Lista con \code{raw}, \code{clean}, \code{name_map}.
#' @export
leer_datos <- function(path, sheet = NULL){
  ext <- tolower(tools::file_ext(path))

  raw <- if (ext %in% c("csv","txt")){
    # No forzar clases ni mostrar el banner de tipos
    suppressWarnings(readr::read_csv(path, show_col_types = FALSE))
  } else {
    # Excel: no tocar tipos; hoja opcional
    suppressWarnings(readxl::read_excel(path, sheet = sheet))
  }

  orig  <- names(raw)
  clean <- janitor::make_clean_names(orig)
  dfc   <- raw; names(dfc) <- clean

  name_map <- tibble::tibble(clean = clean, original = orig)

  list(raw = raw, clean = dfc, name_map = name_map)
}

# -- detectores con nombres ORIGINALES ----------------------------------------

#' Resolver la columna ORIGINAL en datos para un parent dado
#'
#' Prefiere coincidencia EXACTA con el nombre original; si no, usa el
#' equivalente \code{clean} para hallar el original en \code{name_map}.
#'
#' @param parent_name Nombre de variable (puede venir de \code{survey$name}).
#' @param name_map Tibble con columnas \code{clean, original} (de \code{leer_datos}).
#' @return Cadena con el nombre ORIGINAL en la base de datos, o NA si no se encuentra.
#' @export
resolve_parent_col_original <- function(parent_name, name_map){
  if (is.null(parent_name) || is.na(parent_name)) return(NA_character_)
  # 1) exacto contra original
  if (parent_name %in% name_map$original) return(parent_name)
  # 2) por clean
  cl  <- janitor::make_clean_names(parent_name)
  hit <- name_map$original[name_map$clean == cl]
  hit[1] %||% NA_character_
}

#' Encontrar columna de TEXTO tipo *_other/_specify/_text para un parent ORIGINAL
#'
#' Busca una columna que siga la convención "parent + sufijo de texto".
#'
#' @param parent_original Nombre ORIGINAL del parent en los datos.
#' @param name_map Tibble \code{clean, original}.
#' @return Nombre ORIGINAL encontrado o NA.
#' @export
find_text_other_for_parent <- function(parent_original, name_map){
  if (!nzchr(parent_original)) return(NA_character_)
  suf <- c("_other","_otra","_otro","_specify","_text","_elsewhere","_elswhere")
  rx  <- paste0("^", gsub("([\\W])","\\\\\\1", parent_original),
                "(", paste0(suf, collapse="|"), ")$")
  hit <- name_map$original[grepl(rx, name_map$original, ignore.case = TRUE, perl = TRUE)]
  hit[1] %||% NA_character_
}

#' Encontrar columnas dummy "Parent/Code" (select_multiple) para un parent ORIGINAL
#'
#' @param parent_original Nombre ORIGINAL del parent en los datos.
#' @param name_map Tibble \code{clean, original}.
#' @return Vector de nombres ORIGINALES que cumplen el patrón \code{"Parent/Algo"}.
#' @export
find_dummy_cols_for_parent <- function(parent_original, name_map){
  if (!nzchr(parent_original)) return(character(0))
  rx  <- paste0("^", gsub("([\\W])","\\\\\\1", parent_original), "\\s*/\\s*[^/]+$")
  name_map$original[grepl(rx, name_map$original, perl = TRUE)]
}

#' Encontrar dummy específica "/Other" (u "Otro/Otra") para un parent ORIGINAL
#'
#' @param parent_original Nombre ORIGINAL del parent en los datos.
#' @param name_map Tibble \code{clean, original}.
#' @return Nombre ORIGINAL de la dummy "/Other" si existe; NA en caso contrario.
#' @export
find_other_dummy_for_parent <- function(parent_original, name_map){
  dums <- find_dummy_cols_for_parent(parent_original, name_map)
  out  <- dums[grepl("/\\s*(other|otro|otra)\\s*$", dums, ignore.case = TRUE)]
  out[1] %||% NA_character_
}

#' Obtener catálogo de opciones (choices) para un parent dado
#'
#' Cruza \code{survey$name} → \code{list_norm} → \code{choices} y devuelve
#' \code{code/label} para esa lista si aplica.
#'
#' @param parent_name Nombre de pregunta (p. ej., \code{survey$name} o \code{parent}).
#' @param inst Objeto instrumento (lista) con \code{$survey} y \code{$choices} cargados por el lector minimalista.
#'
#' @return Tibble con columnas \code{code} y \code{label} (o \code{NULL} si no aplica).
#' @export
choices_for_parent <- function(parent_name, inst){
  if (!nzchr(parent_name)) return(NULL)
  pv_clean <- janitor::make_clean_names(parent_name)
  row <- inst$survey %>%
    dplyr::filter(.data$name == pv_clean) %>%
    dplyr::slice(1)
  if (!nrow(row)) return(NULL)

  ln <- row$list_norm %||% row$list_name %||% NA_character_
  if (is.na(ln) || !nzchar(ln)) return(NULL)

  opts <- inst$choices %>%
    dplyr::filter(.data$list_norm == ln) %>%
    dplyr::transmute(code = .data$name, label = .data$label_spanish_es %||% .data$label)
  if (!nrow(opts)) NULL else opts
}

#' Label en español desde `survey` (nombre ORIGINAL o limpio)
#'
#' Dado un nombre ORIGINAL de columna (o un \code{parent}), obtiene su label ES
#' desde \code{inst$survey} sin mezclar idiomas. Orden de preferencia:
#' \code{label_spanish_es} > \code{label_es} > \code{label} > nombre limpio.
#'
#' @param var_orig Character (escalar o vector) con nombres ORIGINALES/parent.
#' @param survey Data frame de la hoja `survey` del lector minimalista.
#'
#' @return Character vector con el label en español (o el nombre limpio si no hay label).
#' @export
label_es_from_survey <- function(var_orig, survey){
  if (is.null(var_orig)) return(NA_character_)
  labcol <- {
    nms <- tolower(names(survey))
    if ("label_spanish_es" %in% nms) "label_spanish_es"
    else if ("label_es" %in% nms) "label_es"
    else if ("label" %in% nms) "label"
    else NA_character_
  }
  vapply(as.list(var_orig), function(v){
    if (is.null(v)) return(NA_character_)
    v <- as.character(v)[1]
    if (is.na(v) || !nzchar(v)) return(NA_character_)
    nm <- janitor::make_clean_names(v)
    i  <- match(nm, survey$name)
    if (is.na(i) || is.na(labcol)) return(nm)
    val <- as.character(survey[[labcol]][i])
    if (length(val) == 1 && !is.na(val) && nzchar(trimws(val))) val else nm
  }, FUN.VALUE = character(1))
}

# ---- utilidades locales usadas arriba ---------------------------------------
`%||%` <- function(x, y) if (is.null(x) || (length(x) == 1 && is.na(x))) y else x
nzchr   <- function(x) is.character(x) && length(x) == 1 && !is.na(x) && nzchar(x)


# --- Utilidades para la contruccion de la plantilla-------------------------


# Label ES para survey, buscando primero en survey_raw (sin normalizar)
label_es_from_inst <- function(var, inst){
  if (is.null(var)) return(NA_character_)
  v <- as.character(var)
  nm_clean <- janitor::make_clean_names(v)

  # 1) survey_raw (mejor fuente)
  sr <- inst$survey_raw
  if (!is.null(sr) && "name" %in% names(sr)) {
    sr_clean_names <- janitor::make_clean_names(sr$name)
    i <- match(nm_clean, sr_clean_names)
    escol <- detect_spanish_col(names(sr))
    if (!is.na(escol)) {
      out <- sr[[escol]][i]
      out <- ifelse(is.na(out) | !nzchar(trimws(out)), nm_clean, out)
      return(out)
    }
  }

  # 2) fallback a survey (si trae alias ya)
  s <- inst$survey
  if (!is.null(s) && "name" %in% names(s)) {
    i <- match(nm_clean, s$name)
    escol <- detect_spanish_col(names(s))
    if (!is.na(escol)) {
      out <- s[[escol]][i]
    } else if ("label_es" %in% names(s)) {
      out <- s$label_es[i]
    } else if ("label" %in% names(s)) {
      out <- s$label[i]
    } else {
      out <- nm_clean
    }
    out <- ifelse(is.na(out) | !nzchar(trimws(out)), nm_clean, out)
    return(out)
  }

  nm_clean
}



# ---- 3. Plantilla editable de familias (revisada) --------------------------
#' Escribir plantilla de familias (sin normalizar el instrumento)
#'
#' Genera un Excel \code{familias.xlsx} con sugerencias por cada pregunta
#' \code{select_one}, \code{select_multiple} (y opcionalmente \code{text})
#' detectada en \code{inst$survey}.
#' - NO reescribe/normaliza el instrumento: solo lee \code{type}, \code{name}
#'   y, si existen, \code{q_order}, \code{list_name}, \code{list_norm}.
#' - \strong{parent_label} se obtiene en \emph{español} usando
#'   \code{label_es_from_inst()} que prioriza \code{inst$survey_raw}.
#' - La columna \code{list_norm} se llama exactamente igual que en \code{choices}.
#' - Resalta filas por \code{tipo} en tonos pastel y:
#'     * NO colorea \code{other_dummy_col} ni \code{text_col}, pero les pone
#'       \emph{borde grueso} y \emph{encabezado pastel} para destacarlas.
#'
#' @param inst Objeto devuelto por el lector de XLSForm (con \code{$survey} y \code{$survey_raw}).
#' @param dat  Objeto de \code{leer_datos()} con \code{$raw}, \code{$name_map}.
#' @param path Ruta de salida (.xlsx). Default \code{"familias.xlsx"}.
#' @param incluir_text_vars Incluir preguntas \code{text} como filas (default \code{TRUE}).
#'
#' @return Ruta absoluta al archivo escrito (invisible).
#' @export
escribir_plantilla_familias <- function(inst, dat, path = "familias.xlsx", incluir_text_vars = TRUE){
  stopifnot(is.list(inst), is.list(dat), "survey" %in% names(inst))
  survey <- inst$survey
  if (is.null(survey) || !("name" %in% names(survey))) {
    stop("inst$survey debe existir y tener columna 'name'.")
  }

  # --- auxiliares mínimos, SIN normalizar el instrumento ---
  `%||%` <- function(x, y) if (is.null(x) || (length(x) == 1 && is.na(x))) y else x
  .type_base <- function(x) sub("\\s.*$", "", as.character(x %||% ""))
  .norm_list_name <- function(x){
    x <- tolower(trimws(as.character(x))); x <- gsub("\\s+","_", x); gsub("[^a-z0-9_]", "_", x)
  }

  # Variante local: detectar *_other/_specify/_text/_why (+ español)
  .find_text_for_parent <- function(parent_original, name_map){
    if (is.null(parent_original) || is.na(parent_original) || !nzchar(parent_original)) return(NA_character_)
    suf <- c("_other","_otra","_otro","_specify","_text","_elsewhere","_elswhere","_why")
    rx  <- paste0("^", gsub("([\\W])","\\\\\\1", parent_original),
                  "(", paste0(suf, collapse="|"), ")$")
    hit <- name_map$original[grepl(rx, name_map$original, ignore.case = TRUE, perl = TRUE)]
    hit[1] %||% NA_character_
  }

  # columnas auxiliares locales (no modifican inst en disco)
  tb <- tibble::tibble(
    name       = as.character(survey$name),
    type       = as.character(survey$type %||% NA_character_),
    q_order    = if ("q_order" %in% names(survey)) as.integer(survey$q_order) else seq_len(nrow(survey)),
    type_base  = .type_base(survey$type),
    list_name  = if ("list_name" %in% names(survey)) survey$list_name else trimws(sub("^\\S+\\s+","", as.character(survey$type %||% ""))),
    list_norm  = if ("list_norm" %in% names(survey)) survey$list_norm else .norm_list_name(if ("list_name" %in% names(survey)) survey$list_name else trimws(sub("^\\S+\\s+","", as.character(survey$type %||% ""))))
  )

  # candidatos por tipo
  tipos_objetivo <- c("select_one","select_multiple", if (isTRUE(incluir_text_vars)) "text" else character(0))
  cand <- tb %>%
    dplyr::filter(.data$type_base %in% tipos_objetivo) %>%
    dplyr::transmute(
      parent       = .data$name,
      parent_label = label_es_from_inst(.data$name, inst),  # <- ESPAÑOL desde survey_raw
      tipo_sugerido= .data$type_base,
      q_order      = .data$q_order,
      list_norm    = .data$list_norm
    )

  # sugerencias en nombres ORIGINALES de datos
  parent_col_sug <- vapply(cand$parent, resolve_parent_col_original,
                           FUN.VALUE = character(1), name_map = dat$name_map)
  # usar la variante local (incluye _why). Si prefieres, sustituye por tu find_text_other_for_parent() ampliada.
  text_col_sug   <- vapply(parent_col_sug, .find_text_for_parent,
                           FUN.VALUE = character(1), name_map = dat$name_map)
  otherdum_sug   <- vapply(parent_col_sug, find_other_dummy_for_parent,
                           FUN.VALUE = character(1), name_map = dat$name_map)
  # lista de dummies candidatas "Parent/Code"
  dummy_list <- lapply(parent_col_sug, find_dummy_cols_for_parent, name_map = dat$name_map)
  dummy_cands <- vapply(dummy_list, function(v) paste(v, collapse="; "), FUN.VALUE = character(1))

  tpl <- cand %>%
    dplyr::mutate(
      use               = TRUE,
      tipo              = .data$tipo_sugerido,
      parent_col        = parent_col_sug,
      other_dummy_col   = otherdum_sug,
      text_col          = text_col_sug,
      parent_col_cands  = parent_col_sug,
      other_dummy_cands = otherdum_sug,
      text_col_cands    = text_col_sug,
      dummy_cands       = dummy_cands
    ) %>%
    # ← q_order inmediatamente después de use (como pediste)
    dplyr::select(use, q_order, tipo, parent, parent_label, list_norm,
                  parent_col, other_dummy_col, text_col,
                  parent_col_cands, other_dummy_cands, text_col_cands, dummy_cands)

  # --- Excel con formato -----------------------------------------------------
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "familias")

  # Estilos base
  style_hdr <- openxlsx::createStyle(
    textDecoration = "bold", halign = "center", valign = "center",
    border = "TopBottomLeftRight"
  )
  style_border_all <- openxlsx::createStyle(border = "TopBottomLeftRight", borderColour = "black")

  # Pastel para columnas especiales (sin negrita)
  style_special_hdr <- openxlsx::createStyle(
    fgFill = "#F9D5E5",  # pastel suave
    halign = "center", valign = "center",
    border = "TopBottomLeftRight", borderStyle = "thick"
  )
  style_special_body <- openxlsx::createStyle(
    fgFill = "#F9D5E5",
    border = "TopBottomLeftRight", borderStyle = "thick"
  )

  # Relleno por tipo (para el resto de columnas)
  tipo_fill <- function(t){
    hex <- switch(tolower(t),
                  "select_multiple" = "#E2F0D9",  # verde suave
                  "select_one"      = "#D9E1F2",  # azul suave
                  "text"            = "#FFF2CC",  # amarillo suave
                  "#EEEEEE")
    openxlsx::createStyle(fgFill = hex)
  }

  # Escribir datos y header general
  openxlsx::writeData(wb, "familias", tpl, startRow = 1, colNames = TRUE)
  openxlsx::addStyle(wb, "familias", style_hdr, rows = 1, cols = 1:ncol(tpl), gridExpand = TRUE)

  # Índices de columnas especiales (solo estas van con color distinto)
  idx_other <- which(colnames(tpl) == "other_dummy_col")
  idx_text  <- which(colnames(tpl) == "text_col")
  idx_special <- c(idx_other, idx_text)

  # Columnas que se pintan por fila (todas menos las especiales)
  idx_all <- seq_len(ncol(tpl))
  idx_body_fill <- setdiff(idx_all, idx_special)

  # 1) Relleno por tipo (por fila) solo en columnas NO especiales
  if (nrow(tpl) && length(idx_body_fill)){
    for (i in seq_len(nrow(tpl))){
      openxlsx::addStyle(
        wb, "familias", tipo_fill(tpl$tipo[i]),
        rows = i + 1, cols = idx_body_fill, gridExpand = TRUE, stack = TRUE
      )
    }
  }

  # 2) Aplicar estilo pastel a TODA la columna especial (encabezado + cuerpo)
  if (length(idx_special)){
    # Header (fila 1)
    openxlsx::addStyle(wb, "familias", style_special_hdr,
                       rows = 1, cols = idx_special, gridExpand = TRUE, stack = TRUE)
    # Cuerpo
    if (nrow(tpl)){
      openxlsx::addStyle(wb, "familias", style_special_body,
                         rows = 2:(nrow(tpl)+1), cols = idx_special, gridExpand = TRUE, stack = TRUE)
    }
  }

  # Nada especial para parent_label / list_norm: conservan el color por fila
  # (El resto de utilidades: filtros, freeze, anchos, bordes)
  openxlsx::addFilter(wb, "familias", rows = 1, cols = 1:ncol(tpl))
  openxlsx::freezePane(wb, "familias", firstActiveRow = 2, firstActiveCol = 2)
  openxlsx::setColWidths(wb, "familias", cols = 1:ncol(tpl), widths = "auto")
  openxlsx::addStyle(wb, "familias", style_border_all,
                     rows = 1:(nrow(tpl)+1), cols = 1:ncol(tpl), gridExpand = TRUE, stack = TRUE)


  # Hoja AYUDA mínima
  openxlsx::addWorksheet(wb, "ayuda")
  openxlsx::writeData(wb, "ayuda",
                      "Cómo usar 'familias':
- 'use' = TRUE/FALSE para incluir la fila.
- 'q_order' = número de la pregunta en el cuestionario (orden del survey).
- 'tipo' ∈ {select_one, select_multiple, text}.
- 'parent' = nombre de la pregunta (survey$name).
- 'parent_label' = etiqueta en español tomada del XLSForm (survey_raw).
- 'list_norm' = nombre normalizado de la lista de opciones (igual que en CHOICES).
- 'parent_col' = nombre EXACTO en tus datos (columna padre).
- 'other_dummy_col' = solo select_multiple: dummy 'Parent/Other'.
- 'text_col' = texto related (p. ej. _other/_why/_specify/_text).
Las columnas *_cands son sugerencias hechas a partir de los nombres originales del dataset.")
  openxlsx::setColWidths(wb, "ayuda", cols = 1, widths = 120)

  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
  invisible(normalizePath(path))
}


# ----- 4. Lector de la plantilla de Excel de familias ---------------------
#' Leer y **clasificar** el Excel de familias (con adopciones y huérfanas)
#'
#' Lee la hoja \code{sheet} de un Excel de familias ya editado por la persona usuaria
#' y devuelve subconjuntos listos para construir la plantilla de codificación,
#' además de avisos útiles sobre \emph{adopciones} (qué \code{text_col} fue
#' adoptada por qué SO/SM) y \emph{textos huérfanos} (text sin padre asignado).
#'
#' @details
#' La función:
#' \itemize{
#'   \item Normaliza levemente las columnas de la hoja \code{sheet} (\code{use}, \code{tipo}, etc.).
#'   \item Verifica si \code{parent_col}, \code{text_col} y \code{other_dummy_col} existen en \code{dat$raw}.
#'   \item Acepta filas \strong{select\_multiple} solo si existen \code{text_col} y \code{other_dummy_col}.
#'   \item Acepta filas \strong{select\_one} solo si existe \code{text_col}.
#'   \item Considera \strong{text} finales únicamente si NO aparecen como \code{text_col}
#'         de alguna SO/SM con \code{use = TRUE} (es decir, ya “adoptadas”).
#'   \item Construye un cuadro de \strong{adopciones} (quién adopta a quién) y otro de
#'         \strong{textos huérfanos} (text que no fueron adoptadas por ninguna SO/SM).
#'   \item Enriquece \code{parent\_label\_es} desde el XLSForm (\code{inst$survey\_raw} en español
#'         cuando es posible; si no, \code{inst$survey}).
#'   \item Devuelve un catálogo de \emph{choices} utilizadas (por \code{list\_norm}) para las familias aceptadas.
#' }
#'
#' @param path \code{character}. Ruta al Excel de familias (p. ej. \code{"familias.xlsx"}).
#' @param inst \code{list}. Instrumento XLSForm leído con tu “lector minimalista”; debe contener
#'   al menos \code{$survey}, \code{$choices} (idealmente también \code{$survey\_raw}, \code{$choices\_raw}).
#' @param dat \code{list}. Objeto devuelto por \code{leer_datos()} con \code{$raw} (y usualmente \code{$name\_map}).
#' @param sheet \code{character}. Nombre de la hoja a leer del Excel de familias. Default: \code{"familias"}.
#' @param verbose \code{logical}. Si \code{TRUE}, imprime un resumen y ejemplos de adopciones/huérfanas. Default: \code{TRUE}.
#'
#' @return \code{list} con los elementos:
#' \describe{
#'   \item{\code{familias\_filtradas}}{Tibble con las filas aceptadas (SO/SM/TEXT), tras aplicar reglas y existencia en datos.}
#'   \item{\code{select\_multiple}}{Subconjunto aceptado de \code{familias\_filtradas} para \code{tipo == "select_multiple"}.}
#'   \item{\code{select\_one}}{Subconjunto aceptado de \code{familias\_filtradas} para \code{tipo == "select_one"}.}
#'   \item{\code{text}}{Subconjunto aceptado de \code{familias\_filtradas} para \code{tipo == "text"}
#'         (solo \emph{huérfanas}, i.e., no usadas como \code{text_col} en SO/SM con \code{use=TRUE}).}
#'   \item{\code{familias\_enriquecidas}}{Tibble con \code{parent\_label\_es} y banderas \code{falta\_dummy\_sm}, \code{falta\_text}.}
#'   \item{\code{choices\_usadas}}{Tibble con \code{parent}, \code{parent\_col}, \code{list\_norm}, \code{code}, \code{label\_es}
#'         para las listas utilizadas por las familias aceptadas.}
#'   \item{\code{adopciones}}{Tibble de mapeo de \code{text_col} adoptadas: incluye \code{adoptada\_por\_parent},
#'         \code{adoptada\_por\_label}, \code{tipo\_padre} y si el \emph{padre} existe en los datos.}
#'   \item{\code{textos\_huerfanos}}{Tibble con \code{text_col} de \code{tipo == "text"} que quedaron sin adopción y motivo.}
#'   \item{\code{resumen}}{Tibble con contadores: totales aceptados por tipo, excluidas, \#adopciones y \#huérfanas.}
#' }
#'
#' @section Reglas de aceptación:
#' \itemize{
#'   \item \strong{select\_multiple}: \code{use == TRUE} \emph{y} existen \code{text_col} \emph{y} \code{other_dummy_col} en \code{dat$raw}.
#'   \item \strong{select\_one}: \code{use == TRUE} \emph{y} existe \code{text_col} en \code{dat$raw}.
#'   \item \strong{text}: \code{use == TRUE} \emph{y} su \code{text_col} \emph{no} aparece como hija de una SO/SM aceptada.
#' }
#'
#' @examples
#' \dontrun{
#' inst <- leer_instrumento_xlsform("instrumento.xlsx")
#' dat  <- leer_datos("datos.xlsx")
#' fams <- leer_familias_clasificar("familias.xlsx", inst = inst, dat = dat, sheet = "familias")
#'
#' # inspección rápida:
#' fams$resumen
#' head(fams$adopciones)
#' head(fams$textos_huerfanos)
#' }
#'
#' @seealso \code{\link{construir_plantilla_desde_familias}}, \code{\link{escribir_plantilla_familias}}
#' @export
leer_familias_clasificar <- function(path, inst, dat, sheet = "familias", verbose = TRUE){
  stopifnot(is.list(inst), is.list(dat), "survey" %in% names(inst), "choices" %in% names(inst))
  fam <- readxl::read_excel(path, sheet = sheet) %>%
    janitor::clean_names()

  # === SANITIZACIÓN CRÍTICA ==========================================
  chr_cols <- c("tipo","parent","parent_label","list_norm",
                "parent_col","other_dummy_col","text_col",
                "parent_col_cands","other_dummy_cands","text_col_cands","dummy_cands")
  for (cc in intersect(chr_cols, names(fam))) {
    # fuerza a character y reemplaza NA por cadena vacía
    fam[[cc]] <- as.character(fam[[cc]])
    fam[[cc]][is.na(fam[[cc]])] <- ""
    # limpia espacios invisibles
    fam[[cc]] <- trimws(fam[[cc]])
  }
  # use → lógico robusto
  if (!"use" %in% names(fam)) fam$use <- TRUE
  fam$use <- dplyr::case_when(
    is.logical(fam$use) ~ fam$use,
    tolower(as.character(fam$use)) %in% c("1","true","t","si","sí","yes","y") ~ TRUE,
    tolower(as.character(fam$use)) %in% c("0","false","f","no","n") ~ FALSE,
    TRUE ~ TRUE
  )
  # tipo en minúsculas
  if ("tipo" %in% names(fam)) fam$tipo <- tolower(fam$tipo)
  # ==========================================================================

  # columnas esperadas (tolerante)
  need <- c("use","tipo","parent","parent_label","q_order","list_norm",
            "parent_col","other_dummy_col","text_col")
  for (k in need) if (!k %in% names(fam)) fam[[k]] <- NA

  # normalizaciones suaves
  if (!"use" %in% names(fam)) fam$use <- TRUE
  fam$use  <- as.logical(fam$use)
  fam$tipo <- tolower(trimws(as.character(fam$tipo %||% "")))
  fam$parent_col <- trimws(as.character(fam$parent_col))
  fam$text_col   <- trimws(as.character(fam$text_col))

  # existencia en datos crudos
  cols <- names(dat$raw)
  fam <- fam %>%
    dplyr::mutate(
      exists_parent_col = !is.na(.data$parent_col)      & .data$parent_col      %in% cols,
      exists_text_col   = !is.na(.data$text_col)        & .data$text_col        %in% cols,
      exists_dummy_col  = !is.na(.data$other_dummy_col) & .data$other_dummy_col %in% cols
    )

  # reglas de aceptación por tipo (igual que antes)
  acc_sm <- fam$use & fam$tipo == "select_multiple" & fam$exists_text_col & fam$exists_dummy_col
  acc_so <- fam$use & fam$tipo == "select_one"      & fam$exists_text_col

  # --- NUEVO: adopciones y huérfanas ---------------------------------------
  # Columna de texto EFECTIVA: si tipo=="text" y text_col está vacío, usar parent_col
  fam$text_col_eff <- fam$text_col
  fam$text_col_eff[ fam$tipo == "text" & (is.na(fam$text_col_eff) | !nzchar(fam$text_col_eff)) ] <- fam$parent_col[ fam$tipo == "text" ]

  # set de textos asignados como hijas por alguna SO/SM con use=TRUE
  text_cols_asignadas <- unique(na.omit(fam$text_col[fam$use & fam$tipo %in% c("select_one","select_multiple")]))

  # TEXTO finales: solo las 'text' cuya EFECTIVA NO esté asignada como hija
  acc_tx <- fam$use & fam$tipo == "text" & !(fam$text_col_eff %in% text_cols_asignadas)

  # construir fam_ok con acc_sm|acc_so|acc_tx (sin cambios en SM/SO)
  fam_ok <- fam[acc_sm | acc_so | acc_tx, , drop = FALSE]

  # Para que aguas abajo “text” lleve el nombre correcto, clona text_col = text_col_eff
  fam_ok$text_col[ fam_ok$tipo == "text" ] <- fam_ok$text_col_eff[ fam_ok$tipo == "text" ]

  # --- adopciones pero ya no afecta el cálculo de 'text'
  adopt_rows <- fam[fam$use & fam$tipo %in% c("select_one","select_multiple"), , drop = FALSE]
  adopciones <- tibble::tibble(
    text_col              = text_cols_asignadas,
    adoptada_por_parent   = adopt_rows$parent_col[ match(text_cols_asignadas, adopt_rows$text_col) ] %||% NA_character_,
    adoptada_por_label    = adopt_rows$parent_label[ match(text_cols_asignadas, adopt_rows$text_col) ] %||% NA_character_,
    tipo_padre            = adopt_rows$tipo[ match(text_cols_asignadas, adopt_rows$text_col) ] %||% NA_character_,
    padre_existe_en_datos = adopt_rows$exists_parent_col[ match(text_cols_asignadas, adopt_rows$text_col) ] %||% NA
  )

  # --- huérfanas: usar la EFECTIVA y sacar las ya asignadas
  es_text <- fam$use & fam$tipo == "text"
  huerf <- fam[es_text & !(fam$text_col_eff %in% text_cols_asignadas), , drop = FALSE]
  textos_huerfanos <- tibble::tibble(
    text_col        = huerf$text_col_eff,
    parent_sugerido = huerf$parent_col,
    existe_en_datos = huerf$exists_text_col,   # (si quieres, reevalúa existencia usando text_col_eff)
    motivo          = "No asignada como text_col de ninguna SO/SM con use = TRUE"
  ) %>% dplyr::arrange(!existe_en_datos, text_col)

  # --- (resto) Enriquecer con label ES desde XLSForm  --
  if (exists("label_es_from_inst", mode = "function")) {
    fam_ok$parent_label_es <- label_es_from_inst(fam_ok$parent %||% fam_ok$parent_col, inst)
  } else {
    label_es_from_inst <- function(parent, inst){
      nm <- janitor::make_clean_names(parent)
      s <- inst$survey
      i <- match(nm, janitor::make_clean_names(s$name))
      lab <- rep(NA_character_, length(nm))
      if (!is.null(inst$survey_raw)) {
        sr <- inst$survey_raw
        col_es <- grep("^label(::)?spanish|label[_:]spanish|label[_:]es$", tolower(names(sr)), value = TRUE)[1]
        if (!is.na(col_es)) {
          j <- match(nm, janitor::make_clean_names(sr$name))
          lab_ok <- ifelse(!is.na(j), as.character(sr[[col_es]][j]), NA_character_)
          lab <- ifelse(!is.na(lab_ok) & nzchar(lab_ok), lab_ok, lab)
        }
      }
      if (any(is.na(lab))) {
        cand <- c("label_spanish_es","label_es","label")
        col2 <- cand[cand %in% names(s)]
        if (length(col2)) {
          v <- s[[col2[1]]]; lab2 <- ifelse(!is.na(i), as.character(v[i]), NA_character_)
          lab <- ifelse(is.na(lab) & !is.na(lab2) & nzchar(lab2), lab2, lab)
        }
      }
      lab[is.na(lab) | !nzchar(lab)] <- nm[is.na(lab) | !nzchar(lab)]
      lab
    }
    fam_ok$parent_label_es <- label_es_from_inst(fam_ok$parent %||% fam_ok$parent_col, inst)
  }

  # dividir por tipo (como ya lo tienes)
  sm  <- fam_ok[fam_ok$tipo == "select_multiple", , drop = FALSE]
  so  <- fam_ok[fam_ok$tipo == "select_one",      , drop = FALSE]
  tx  <- fam_ok[fam_ok$tipo == "text",            , drop = FALSE]

  # catálogo de choices (igual que tu versión)
  choices_usadas <- NULL
  if (nrow(fam_ok)) {
    with_ln <- fam_ok %>% dplyr::filter(!is.na(.data$list_norm) & nzchar(.data$list_norm))
    if (nrow(with_ln)) {
      choices_usadas <- with_ln %>%
        dplyr::select(parent, parent_col, list_norm, tipo) %>%
        dplyr::distinct() %>%
        dplyr::left_join(
          inst$choices %>%
            dplyr::transmute(
              list_norm,
              code     = as.character(.data$name),
              label_es = as.character(
                if ("label_spanish_es" %in% names(inst$choices)) .data$label_spanish_es else
                  if ("label" %in% names(inst$choices)) .data$label else .data$name
              )
            ),
          by = "list_norm"
        ) %>%
        dplyr::arrange(.data$parent, .data$code)
    }
  }
  if (is.null(choices_usadas)) {
    choices_usadas <- tibble::tibble(parent = character(), parent_col = character(),
                                     list_norm = character(), tipo = character(),
                                     code = character(), label_es = character())
  }

  # Resumen + mini-avisos
  resumen <- tibble::tibble(
    total_filas_excel = nrow(fam),
    aceptadas_total   = nrow(fam_ok),
    aceptadas_sm      = nrow(sm),
    aceptadas_so      = nrow(so),
    aceptadas_text    = nrow(tx),
    excluidas         = nrow(fam) - nrow(fam_ok),
    textos_adoptados  = nrow(adopciones),
    textos_huerfanos  = nrow(textos_huerfanos)
  )

  if (isTRUE(verbose)) {
    cat("\n[Familias] SO aceptadas:", nrow(so),
        "| SM aceptadas:", nrow(sm),
        "| TEXT finales:", nrow(tx), "\n")
    cat("[Adopciones] text_col adoptadas:", nrow(adopciones), "\n")
    if (nrow(adopciones)) {
      print(utils::head(adopciones, 5))
    }
    cat("[Huérfanas] text sin adopción:", nrow(textos_huerfanos), "\n")
    if (nrow(textos_huerfanos)) {
      print(utils::head(textos_huerfanos, 5))
      cat("→ Sugerencia: asigna estas 'text' en la columna 'text_col' de alguna SO/SM y marca 'use = TRUE'.\n")
    }
  }

  list(
    familias_filtradas   = fam_ok,
    select_multiple      = sm,
    select_one           = so,
    text                 = tx,
    familias_enriquecidas= fam_ok %>% dplyr::mutate(
      falta_dummy_sm = (.data$tipo == "select_multiple") & !.data$exists_dummy_col,
      falta_text     = (.data$tipo %in% c("select_multiple","select_one")) & !.data$exists_text_col
    ),
    choices_usadas       = choices_usadas,
    # NUEVO:
    adopciones           = adopciones,
    textos_huerfanos     = textos_huerfanos,
    resumen              = resumen
  )
}

# ---------------------------------------------------------------------------
# 5) CONSTRUCTOR DE PLANTILLA (prioriza labels desde inst)
# ---------------------------------------------------------------------------

#' @title Construir plantilla (datos + metas) desde familias validadas (sin normalizar inst)
#' @description
#' Constructor robusto que NO normaliza `inst`. Lee labels en español desde
#' `inst$survey_raw` (columna `label::Spanish (ES)` o variantes) y, en su defecto,
#' desde `inst$survey` (p.ej. `label_spanish_es` o `label`). Soporta como `split`
#' el objeto devuelto por `leer_familias_clasificar()` (recomendado).
#'
#' Incluye:
#' - Orden por `q_order` en navegación y creación de hojas.
#' - Labels correctos (survey y choices).
#' - No crea hojas TEXT que ya tengan padre (excluye text adoptadas).
#' - DICCIONARIO con `variable`, `etiqueta`, `tipo_base`, `list_name`.
#' - CHOICES con `parent_col`, `list_name`, `tipo`, `choice_code`, `choice_label`.
#' - INTEGER integradas como hojas con `_recod` y color morado pastel.
#' - En SM no se genera `Seleccionadas_recod`; fila 2 etiqueta bajo "Seleccionadas".
#'
#' @param inst Lista con `survey`, `survey_raw`, `choices`, `choices_raw`.
#' @param dat  Lista de `leer_datos()` con al menos `raw`.
#' @param split Lista con `select_one`, `select_multiple`, `text` (y opcional `integer`).
#'              Si proviene de `leer_familias_clasificar()`, puede traer `choices_usadas`
#'              y `adopciones`.
#' @return lista con: `diccionario`, `choices`, `familias`, `navegacion`, `sheets`
#' @export
construir_plantilla_desde_familias <- function(inst, dat, split){
  stopifnot(is.list(inst), is.list(dat), is.list(split))

  survey  <- inst$survey  %||% inst$survey_raw
  choices <- inst$choices %||% inst$choices_raw
  if (is.null(survey) || is.null(choices)) rlang::abort("inst debe traer $survey y $choices.")

  # --- DICCIONARIO (metadatos directos del XLSForm) --------------------------
  survey_dic <- inst$survey %||% inst$survey_raw
  if (!"q_order" %in% names(survey_dic) || all(is.na(survey_dic$q_order))) {
    survey_dic$q_order <- seq_len(nrow(survey_dic))
  }
  survey_dic$type_base <- sub("\\s.*$", "", as.character(survey_dic$type %||% ""))
  if (!"list_name" %in% names(survey_dic) || all(is.na(survey_dic$list_name))) {
    survey_dic$list_name <- trimws(sub("^\\S+\\s+","", as.character(survey_dic$type %||% "")))
  }

  diccionario <- tibble::tibble(
    q_order   = as.integer(survey_dic$q_order),
    variable  = as.character(survey_dic$name),
    etiqueta  = s_lab_from_original(survey_dic$name, inst),
    tipo_base = as.character(survey_dic$type_base),
    list_name = as.character(survey_dic$list_name)
  )

  # --- FAMILIAS (apilan split) -----------------------------------------------
  fam_all <- dplyr::bind_rows(
    (split$select_one      %||% tibble::tibble()) %>% dplyr::mutate(tipo = "select_one"),
    (split$select_multiple %||% tibble::tibble()) %>% dplyr::mutate(tipo = "select_multiple"),
    (split$text            %||% tibble::tibble()) %>% dplyr::mutate(tipo = "text"),
    (split$integer         %||% tibble::tibble()) %>% dplyr::mutate(tipo = "integer")
  )

  for (cc in c("parent","parent_label","list_name","list_norm",
               "parent_col","other_dummy_col","text_col","tipo")) {
    if (!cc %in% names(fam_all)) fam_all[[cc]] <- ""
    fam_all[[cc]] <- trimws(as.character(fam_all[[cc]]))
    fam_all[[cc]][is.na(fam_all[[cc]])] <- ""
  }

  if (!nrow(fam_all)) {
    return(list(diccionario=diccionario, choices=tibble::tibble(),
                familias=tibble::tibble(), navegacion=tibble::tibble(), sheets=list()))
  }

  # columnas mínimas
  for (k in c("q_order","list_name","list_norm","parent_label","parent",
              "parent_col","other_dummy_col","text_col")) {
    if (!k %in% names(fam_all)) fam_all[[k]] <- NA
  }

  # parent limpio (para cruce)
  fam_all$parent_clean <- janitor::make_clean_names(
    ifelse(!is.na(fam_all$parent) & nzchar(fam_all$parent), fam_all$parent, fam_all$parent_col)
  )

  # completar parent_label desde inst si falta
  missing_pl <- is.na(fam_all$parent_label) | !nzchar(fam_all$parent_label)
  if (any(missing_pl)) {
    pref <- ifelse(!is.na(fam_all$parent) & nzchar(fam_all$parent), fam_all$parent, fam_all$parent_col)
    fam_all$parent_label[missing_pl] <- s_lab_from_original(pref[missing_pl], inst)
  }

  # completar list_name / list_norm desde diccionario si faltan
  if (any(is.na(fam_all$list_name) | !nzchar(fam_all$list_name))) {
    j <- match(fam_all$parent_clean, janitor::make_clean_names(diccionario$variable))
    fam_all$list_name <- dplyr::coalesce(as.character(fam_all$list_name), as.character(diccionario$list_name[j]))
  }
  if (any(is.na(fam_all$list_norm) | !nzchar(fam_all$list_norm))) {
    fam_all$list_norm <- tolower(gsub("[^a-z0-9_]", "_", gsub("\\s+","_", as.character(fam_all$list_name))))
  }

  # excluir TEXT adoptadas
  assigned_texts <- unique(na.omit(c(
    tryCatch(split$adopciones$text_col, error = function(e) NULL),
    tryCatch(split$select_one$text_col, error = function(e) NULL),
    tryCatch(split$select_multiple$text_col, error = function(e) NULL)
  )))
  fam_all <- fam_all %>%
    dplyr::filter(!(tipo == "text" & !is.na(text_col) & nzchar(text_col) & text_col %in% assigned_texts))

  # familias (para auditoría/export)
  familias_tbl <- fam_all %>%
    dplyr::select(
      tipo, parent = parent_clean, parent_label, q_order,
      list_name, list_norm, parent_col, other_dummy_col, text_col
    )

  # --- CHOICES ES canónicas desde inst (fuente de verdad) --------------------
  # choices_es_tbl(inst) -> list_norm, list_name, code, label_es
  choices_es <- choices_es_tbl(inst) %>%
    dplyr::mutate(
      label_es = dplyr::coalesce(.data$label_es, .data$code)
    ) %>%
    dplyr::select(list_norm, list_name, code, label_es)

  # Catálogo de choices por familia SO/SM
  make_choices_tbl <- function(fam_tbl){
    fam_so_sm <- fam_tbl %>%
      dplyr::filter(.data$tipo %in% c("select_one","select_multiple"),
                    !is.na(list_norm) & nzchar(list_norm)) %>%
      dplyr::distinct(parent_col, list_norm, tipo)

    # Si vienen choices_usadas, úsalo sólo como filtro (no para labels)
    if (!is.null(split$choices_usadas) && nrow(split$choices_usadas)) {
      allowed <- split$choices_usadas %>%
        dplyr::select(list_norm, code) %>% dplyr::distinct()
      choices_es_use <- dplyr::inner_join(choices_es, allowed, by = c("list_norm","code"))
    } else {
      choices_es_use <- choices_es
    }

    # mapear list_name por list_norm sin crear .x/.y
    map_ln <- choices_es %>% dplyr::distinct(list_norm, list_name)
    fam_so_sm %>%
      dplyr::mutate(
        variable_base  = janitor::make_clean_names(parent_col),
        variable_label = s_lab_from_original(parent_col, inst)
      ) %>%
      # completar list_name vía vectorizado (sin join con sufijos)
      dplyr::mutate(
        variable_base  = janitor::make_clean_names(parent_col),
        variable_label = s_lab_from_original(parent_col, inst),
        list_name      = map_ln$list_name[match(list_norm, map_ln$list_norm)]
      ) %>%
      # expandir a todas las opciones de esa lista
      dplyr::left_join(
        choices_es_use %>% dplyr::transmute(
          list_norm, choice_code = code, choice_label = label_es
        ),
        by = "list_norm"
      ) %>%
      dplyr::mutate(choice_label = dplyr::coalesce(choice_label, choice_code)) %>%
      dplyr::select(parent_col, variable_base, variable_label,
                    tipo, list_name, list_norm, choice_code, choice_label)
  }

  choices_tbl <- make_choices_tbl(familias_tbl)

  # --- IDs base --------------------------------------------------------------
  resolve_ids <- function(dat_raw){
    as_chr <- function(x){
      if (is.null(x)) return(rep(NA_character_, nrow(dat_raw)))
      if (is.factor(x)) x <- as.character(x)
      as.character(x)
    }
    as_int <- function(x){
      if (is.null(x)) return(rep(NA_integer_, nrow(dat_raw)))
      suppressWarnings(as.integer(x))
    }
    uuid_out <- Reduce(dplyr::coalesce, lapply(list(
      dat_raw[["_uuid"]], dat_raw[["uuid"]], dat_raw[["meta_instance_id"]],
      dat_raw[["instanceid"]], dat_raw[["_id"]]
    ), as_chr))
    idx_out  <- as_int(dat_raw[["_index"]])
    pulso_out<- Reduce(dplyr::coalesce, lapply(list(
      dat_raw[["mand_location_details_pulso_code"]],
      dat_raw[["Pulso_code"]], dat_raw[["pulso_code"]]
    ), as_chr))
    tibble::tibble(`_uuid` = uuid_out, `_index` = idx_out, `Código pulso` = pulso_out)
  }
  dat_raw <- dat$raw
  id_base <- resolve_ids(dat_raw)

  # --- expandir dummies SM ---------------------------------------------------
  expand_sm_dummies <- function(dat_raw, parent_col, opts){
    n <- nrow(dat_raw)
    out <- matrix(NA_integer_, nrow=n, ncol=nrow(opts))
    colnames(out) <- opts$choice_code
    # slash
    for (j in seq_len(nrow(opts))){
      slash <- paste0(parent_col, "/", opts$choice_code[j])
      if (slash %in% names(dat_raw)) {
        v <- dat_raw[[slash]]
        vv <- suppressWarnings(as.integer(as.character(v)))
        if (all(is.na(vv))){
          vv <- ifelse(tolower(as.character(v)) %in% c("true","t","1"), 1L,
                       ifelse(tolower(as.character(v)) %in% c("false","f","0"), 0L, NA_integer_))
        }
        out[, j] <- vv
      }
    }
    # tokens
    miss <- which(apply(out, 2, function(z) all(is.na(z))))
    if (length(miss) && parent_col %in% names(dat_raw)){
      toks <- strsplit(ifelse(is.na(dat_raw[[parent_col]]), "", as.character(dat_raw[[parent_col]])), "\\s+")
      for (j in miss){
        code <- opts$choice_code[j]
        out[, j] <- vapply(toks, function(tt) as.integer(code %in% tt), integer(1))
      }
    }
    tibble::as_tibble(out)
  }

  # ---------- orden de creación por q_order ----------------------------------
  qord_by_var <- diccionario %>% dplyr::transmute(var = variable, q_order)
  ord_val <- function(x){
    v <- janitor::make_clean_names(x)
    qord_by_var$q_order[match(v, janitor::make_clean_names(qord_by_var$var))]
  }

  sel1 <- (split$select_one %||% tibble::tibble())
  if (nrow(sel1)) sel1 <- sel1 %>% dplyr::mutate(.ord = ord_val(parent_col)) %>%
    dplyr::arrange(.ord, parent_col) %>% dplyr::select(-.ord)
  selm <- (split$select_multiple %||% tibble::tibble())
  if (nrow(selm)) selm <- selm %>% dplyr::mutate(.ord = ord_val(parent_col)) %>%
    dplyr::arrange(.ord, parent_col) %>% dplyr::select(-.ord)
  sint <- (split$integer %||% tibble::tibble())
  if (nrow(sint)) sint <- sint %>% dplyr::mutate(.ord = ord_val(parent_col %||% parent)) %>%
    dplyr::arrange(.ord, dplyr::coalesce(parent_col, parent)) %>% dplyr::select(-.ord)
  stxt <- (split$text %||% tibble::tibble())
  if (nrow(stxt)) {
    stxt <- stxt %>% dplyr::mutate(.ord = ord_val(parent_col)) %>%
      dplyr::arrange(.ord, parent_col) %>% dplyr::select(-.ord)
  }

  sheets_list <- list(); nav_rows <- list()

  # ---------- SELECT ONE ----------
  if (!is.null(sel1) && nrow(sel1)){
    for (i in seq_len(nrow(sel1))){
      row <- sel1[i, ]
      parent_col <- row$parent_col
      text_col   <- row$text_col
      tipo <- "select_one"

      opts <- choices_tbl %>%
        dplyr::filter(parent_col == !!parent_col) %>%
        dplyr::distinct(choice_code, choice_label) %>%
        dplyr::transmute(code = choice_code, label = choice_label)

      parent_code  <- dat_raw[[parent_col]] %||% NA_character_
      parent_label <- if (nrow(opts)) opts$label[match(parent_code, opts$code)] else NA_character_
      text_vec     <- if (!is.na(text_col) && nzchar(text_col) && text_col %in% names(dat_raw)) as.character(dat_raw[[text_col]]) else NA_character_

      base <- id_base %>%
        dplyr::mutate(`Selección (código)` = as.character(parent_code),
                      `Selección (label)`  = as.character(parent_label)) %>%
        dplyr::mutate(!!text_col := text_vec)

      if (!is.na(text_col) && nzchar(text_col)) base[[paste0(text_col,"_recod")]] <- NA_character_
      base[["Control"]] <- NA_character_

      cn <- colnames(base)
      base_cols <- cn[!grepl("_recod$", cn)]
      rec_cols  <- cn[grepl("_recod$", cn)]
      order_cols <- c()
      for (b in base_cols){
        order_cols <- c(order_cols, b)
        r <- paste0(b, "_recod")
        if (r %in% rec_cols) order_cols <- c(order_cols, r)
      }
      orphan_rec <- setdiff(rec_cols, paste0(base_cols, "_recod"))
      base <- base[, unique(c(order_cols, orphan_rec)), drop = FALSE]

      hdr_raw <- names(base)
      hdr_lab <- hdr_raw
      hdr_lab[hdr_raw=="_uuid"] <- "UUID"
      hdr_lab[hdr_raw=="_index"] <- "Índice"
      hdr_lab[hdr_raw=="Código pulso"] <- "Código pulso"
      hdr_lab[hdr_raw=="Selección (código)"] <- "Selección (código)"
      hdr_lab[hdr_raw=="Selección (label)"]  <- s_lab_from_original(parent_col, inst)
      if (!is.na(text_col) && nzchar(text_col)) hdr_lab[hdr_raw==text_col] <- s_lab_from_original(text_col, inst)
      hdr_lab[grepl("_recod$", hdr_raw)] <- "Recodificación"

      stitle <- parent_col
      sheets_list[[stitle]] <- structure(base, header_raw = hdr_raw, label_row = hdr_lab, tipo = tipo)
      nav_rows[[length(nav_rows)+1]] <- tibble::tibble(hoja = stitle, tipo = tipo, n = nrow(base))
    }
  }

  # ---------- SELECT MULTIPLE ----------
  if (!is.null(selm) && nrow(selm)){
    for (i in seq_len(nrow(selm))){
      row <- selm[i, ]
      parent_col     <- row$parent_col
      other_dummy_col<- row$other_dummy_col
      text_col       <- row$text_col
      tipo <- "select_multiple"

      opts <- choices_tbl %>%
        dplyr::filter(parent_col == !!parent_col) %>%
        dplyr::distinct(choice_code, choice_label) %>%
        dplyr::transmute(choice_code, choice_label)

      dmm <- if (nrow(opts)) expand_sm_dummies(dat_raw, parent_col, opts) else tibble::tibble()
      if (nrow(opts) && ncol(dmm)) {
        lab <- as.character(opts$choice_label)
        lab[is.na(lab) | !nzchar(lab)] <- as.character(opts$choice_code[is.na(lab) | !nzchar(lab)])
        names(dmm) <- lab
      }

      # dummy Other
      if (!is.na(other_dummy_col) && nzchar(other_dummy_col) && other_dummy_col %in% names(dat_raw)) {
        other_dummy <- dat_raw[[other_dummy_col]]
        other01 <- suppressWarnings(as.integer(as.character(other_dummy)))
        if (all(is.na(other01)))
          other01 <- ifelse(tolower(as.character(other_dummy)) %in% c("true","t","1"), 1L,
                            ifelse(tolower(as.character(other_dummy)) %in% c("false","f","0"), 0L, NA_integer_))
        dmm[["Otro, por favor especificar"]] <- other01
      } else if (nrow(opts)) {
        tmp_txt <- if (!is.na(text_col) && nzchar(text_col) && text_col %in% names(dat_raw)) as.character(dat_raw[[text_col]]) else NA_character_
        dmm[["Otro, por favor especificar"]] <- ifelse(!is.na(tmp_txt) & nzchar(tmp_txt), 1L, 0L)
      }

      # sin "Otro"
      nm <- names(dmm); nm <- nm[!is.na(nm) & nzchar(nm)]
      no_other <- setdiff(nm, "Otro, por favor especificar")

      # seleccionadas (labels y códigos)
      sel_labels <- if (length(no_other)) apply(dmm[, no_other, drop = FALSE], 1, function(r){
        idx <- which(r == 1); if (!length(idx)) "" else paste(names(r)[idx], collapse = "; ")
      }) else rep("", nrow(dat_raw))

      sel_codes <- if (length(no_other) && nrow(opts)) {
        lab2code <- opts$choice_code; names(lab2code) <- as.character(opts$choice_label)
        names(lab2code)[is.na(names(lab2code)) | !nzchar(names(lab2code))] <- as.character(opts$choice_code[is.na(opts$choice_label) | !nzchar(opts$choice_label)])
        apply(dmm[, no_other, drop = FALSE], 1, function(r){
          idx <- which(r == 1); if (!length(idx)) "" else paste(lab2code[names(r)[idx]], collapse = "; ")
        })
      } else rep("", nrow(dat_raw))

      text_vec <- if (!is.na(text_col) && nzchar(text_col) && text_col %in% names(dat_raw)) as.character(dat_raw[[text_col]]) else NA_character_

      base <- id_base %>%
        dplyr::mutate(Seleccionadas     = sel_labels,
                      Seleccionadas_cod = sel_codes) %>%
        dplyr::bind_cols(dmm) %>%
        dplyr::mutate(!!text_col := text_vec)

      # recods intercalados (no para Seleccionadas / Seleccionadas_cod)
      skip_cols <- c("_uuid","_index","Código pulso","Seleccionadas","Seleccionadas_cod")
      for (dc in setdiff(names(base), skip_cols)) base[[paste0(dc,"_recod")]] <- NA_character_
      base[["Control"]] <- NA_character_

      cn <- colnames(base)
      base_cols <- cn[!grepl("_recod$", cn)]
      rec_cols  <- cn[grepl("_recod$", cn)]
      order_cols <- c()
      for (b in base_cols){
        order_cols <- c(order_cols, b)
        r <- paste0(b, "_recod")
        if (r %in% rec_cols) order_cols <- c(order_cols, r)
      }
      orphan_rec <- setdiff(rec_cols, paste0(base_cols, "_recod"))
      base <- base[, unique(c(order_cols, orphan_rec)), drop = FALSE]

      hdr_raw <- names(base); hdr_lab <- hdr_raw
      hdr_lab[hdr_raw=="_uuid"]             <- "UUID"
      hdr_lab[hdr_raw=="_index"]            <- "Índice"
      hdr_lab[hdr_raw=="Código pulso"]      <- "Código pulso"
      hdr_lab[hdr_raw=="Seleccionadas"]     <- s_lab_from_original(parent_col, inst)
      hdr_lab[hdr_raw=="Seleccionadas_cod"] <- "Seleccionadas (código)"
      if (!is.na(text_col) && nzchar(text_col)) hdr_lab[hdr_raw==text_col] <- s_lab_from_original(text_col, inst)
      hdr_lab[grepl("_recod$", hdr_raw)]    <- "Recodificación"

      stitle <- parent_col
      sheets_list[[stitle]] <- structure(base, header_raw = hdr_raw, label_row = hdr_lab, tipo = tipo)
      nav_rows[[length(nav_rows)+1]] <- tibble::tibble(hoja = stitle, tipo = tipo, n = nrow(base))
    }
  }

  # ---------- INTEGER ----------
  if (!is.null(sint) && nrow(sint)){
    for (i in seq_len(nrow(sint))){
      row <- sint[i, ]
      var_col <- if (!is.null(row$parent_col) && nzchar(row$parent_col)) row$parent_col else row$parent
      if (!nzchar(var_col) || !(var_col %in% names(dat$raw))) next
      tipo <- "integer"

      val <- dat_raw[[var_col]]
      base <- id_base %>%
        dplyr::mutate(!!var_col := val)
      base[[paste0(var_col,"_recod")]] <- NA_character_
      base[["Control"]] <- NA_character_

      hdr_raw <- names(base); hdr_lab <- hdr_raw
      hdr_lab[hdr_raw=="_uuid"]        <- "UUID"
      hdr_lab[hdr_raw=="_index"]       <- "Índice"
      hdr_lab[hdr_raw=="Código pulso"] <- "Código pulso"
      hdr_lab[hdr_raw==var_col]        <- s_lab_from_original(var_col, inst)
      hdr_lab[grepl("_recod$", hdr_raw)] <- "Recodificación"

      stitle <- var_col
      sheets_list[[stitle]] <- structure(base, header_raw = hdr_raw, label_row = hdr_lab, tipo = tipo)
      nav_rows[[length(nav_rows)+1]] <- tibble::tibble(hoja = stitle, tipo = tipo, n = nrow(base))
    }
  }

  # ---------- TEXT (huérfanas) -----------------------------------------------
  if (!is.null(stxt) && nrow(stxt)){
    for (i in seq_len(nrow(stxt))){
      row <- stxt[i, ]
      txt_col <- row$parent_col
      tipo <- "text"
      if (!(txt_col %in% names(dat$raw))) next
      if (txt_col %in% assigned_texts) next

      txt <- as.character(dat$raw[[txt_col]])
      base <- id_base %>%
        dplyr::mutate(!!txt_col := txt)
      base[[paste0(txt_col,"_recod")]] <- NA_character_
      base[["Control"]] <- NA_character_

      cn <- colnames(base)
      base <- base[, c(cn[!grepl("_recod$", cn)], cn[grepl("_recod$", cn)]), drop = FALSE]

      hdr_raw <- names(base); hdr_lab <- hdr_raw
      hdr_lab[hdr_raw=="_uuid"]            <- "UUID"
      hdr_lab[hdr_raw=="_index"]           <- "Índice"
      hdr_lab[hdr_raw=="Código pulso"]     <- "Código pulso"
      hdr_lab[hdr_raw==txt_col]            <- s_lab_from_original(txt_col, inst)
      hdr_lab[grepl("_recod$", hdr_raw)]   <- "Recodificación"

      stitle <- txt_col
      sheets_list[[stitle]] <- structure(base, header_raw = hdr_raw, label_row = hdr_lab, tipo = tipo)
      nav_rows[[length(nav_rows)+1]] <- tibble::tibble(hoja = stitle, tipo = tipo, n = nrow(base))
    }
  }

  # ---------- NAVEGACIÓN -----------------------------------------------------
  nav_df <- dplyr::bind_rows(nav_rows)
  dic2 <- diccionario %>% dplyr::mutate(variable_raw = survey_dic$name)
  nav_df <- nav_df %>%
    dplyr::left_join(dic2 %>% dplyr::select(variable, variable_raw, q_order),
                     by = c("hoja" = "variable")) %>%
    dplyr::mutate(q_order = dplyr::coalesce(q_order,
                                            dic2$q_order[match(hoja, dic2$variable_raw)])) %>%
    dplyr::arrange(q_order,
                   factor(tipo, levels=c("select_one","select_multiple","integer","text")),
                   hoja)

  list(
    diccionario = diccionario,
    choices     = choices_tbl,
    familias    = familias_tbl,
    navegacion  = nav_df %>% dplyr::select(hoja, tipo, n),
    sheets      = sheets_list
  )
}


# ------ 6) Exportador a Excel con formato -----------------------------------

suppressPackageStartupMessages({
  library(openxlsx)
  library(dplyr)
  library(rlang)
})

`%||%` <- function(x, y) if (is.null(x) || (length(x) == 1 && is.na(x))) y else x

safe_sheet_name <- function(x, used = character(0)) {
  x <- as.character(x %||% "Hoja")
  x <- trimws(x)
  x <- gsub("[\\[\\]\\*\\:\\?/\\\\]", "_", x)
  x <- gsub("^'+|'+$", "", x)
  if (!nzchar(x)) x <- "Hoja"
  if (nchar(x) > 31) x <- substr(x, 1, 31)
  base <- x; k <- 1L
  while (x %in% used) {
    suf <- paste0(" (", k, ")")
    maxlen <- 31 - nchar(suf)
    x <- paste0(substr(base, 1, maxlen), suf)
    k <- k + 1L
  }
  x
}

# Paleta por tipo (incluye integer morado)
tipo_hex <- function(tipo) {
  t <- tolower(as.character(tipo %||% ""))
  if (t == "select_multiple") return("#E2F0D9") # verde suave
  if (t == "select_one")      return("#D9E1F2") # azul suave
  if (t == "text")            return("#FFF2CC") # amarillo suave
  if (t == "integer")         return("#E6D9F2") # morado pastel
  "#EEEEEE"
}

.set_widths_smart <- function(wb, sheet, ncols, nrows) {
  if (isTRUE(nrows > 2000L || ncols > 60L)) {
    openxlsx::setColWidths(wb, sheet, cols = 1:ncols, widths = 12)
  } else {
    openxlsx::setColWidths(wb, sheet, cols = 1:ncols, widths = "auto")
  }
}

width_for <- function(ncols){
  if (is.na(ncols) || ncols <= 0) return(18)
  if (ncols <= 20) return("auto")
  if (ncols <= 60) return(22)
  18
}

#' Exportar plantilla a Excel con formato (dos filas de encabezado)
#'
#' - Colorea por tipo en NAVEGACION, FAMILIAS, DICCIONARIO y CHOICES
#'   (incluye morado para integer).
#' - Las hojas de variables tienen fila 1 (código/crudo) y fila 2 (labels),
#'   sin `Seleccionadas_recod` en SM.
#'
#' @param plantilla lista de `construir_plantilla_desde_familias()`
#' @param path_xlsx ruta de salida (default "PPRA_Plantilla_Codificacion.xlsx")
#' @param inst (opcional) objeto instrumento (con survey_raw y choices_raw) para
#'             reforzar labels en DICCIONARIO si aplica.
#' @param autofiltro TRUE para habilitar AutoFilter en fila 2
#' @param congelar_encabezado TRUE para freeze pane (fila 3)
#' @export
exportar_plantilla_codificacion_xlsx <- function(plantilla,
                                                 path_xlsx = "PPRA_Plantilla_Codificacion.xlsx",
                                                 inst = NULL,
                                                 autofiltro = TRUE,
                                                 congelar_encabezado = TRUE){
  stopifnot(is.list(plantilla),
            all(c("diccionario","choices","familias","navegacion","sheets") %in% names(plantilla)))

  wb <- openxlsx::createWorkbook()
  used <- character(0)
  add_sheet <- function(title){
    nm <- safe_sheet_name(title, used); used <<- c(used, nm)
    openxlsx::addWorksheet(wb, nm); nm
  }

  # Estilos
  style_hdr1 <- openxlsx::createStyle(textDecoration = "bold", halign = "center", valign = "center",
                                      border = "TopBottomLeftRight")
  style_hdr2 <- openxlsx::createStyle(textDecoration = "italic", wrapText = TRUE, halign = "center", valign = "center",
                                      border = "TopBottomLeftRight")
  fill_tipo <- function(hex) openxlsx::createStyle(fgFill = hex)
  wrap_left <- openxlsx::createStyle(wrapText = TRUE, halign = "left", valign = "top")
  border_all_black <- openxlsx::createStyle(border = "TopBottomLeftRight", borderColour = "black")

  # ===== 1) NAVEGACION =====
  st_nav <- add_sheet("NAVEGACION")
  nav <- plantilla$navegacion %>% dplyr::select(hoja, tipo, n)
  openxlsx::writeData(wb, st_nav, nav, startRow = 1, colNames = TRUE)

  if (nrow(nav)){
    links <- vapply(nav$hoja, function(h){
      tgt <- safe_sheet_name(h, character(0))
      sprintf('HYPERLINK("#\'%s\'!A1","%s")', tgt, h)
    }, FUN.VALUE = character(1))
    openxlsx::writeFormula(wb, st_nav, x = links, startCol = 1, startRow = 2)

    if ("tipo" %in% names(nav)){
      idx_by_tipo <- split(seq_len(nrow(nav)) + 1L, nav$tipo)
      for (tp in names(idx_by_tipo)){
        openxlsx::addStyle(wb, st_nav, fill_tipo(tipo_hex(tp)),
                           rows = idx_by_tipo[[tp]], cols = 1:3, gridExpand = TRUE, stack = TRUE)
      }
    }
  }
  .set_widths_smart(wb, st_nav, 3, nrow(nav) + 1L)
  openxlsx::freezePane(wb, st_nav, firstActiveRow = 2, firstActiveCol = 2)
  openxlsx::addStyle(wb, st_nav, border_all_black,
                     rows = 1:(nrow(nav)+1), cols = 1:3, gridExpand = TRUE, stack = TRUE)

  # ===== 2) FAMILIAS =====
  st_fam <- add_sheet("FAMILIAS")
  fam_tbl <- plantilla$familias
  openxlsx::writeData(wb, st_fam, fam_tbl, startRow = 1, colNames = TRUE)
  if (nrow(fam_tbl) && "tipo" %in% names(fam_tbl)){
    idx_by_tipo <- split(seq_len(nrow(fam_tbl)) + 1L, fam_tbl$tipo)
    for (tp in names(idx_by_tipo)){
      openxlsx::addStyle(wb, st_fam, fill_tipo(tipo_hex(tp)),
                         rows = idx_by_tipo[[tp]], cols = 1:ncol(fam_tbl), gridExpand = TRUE, stack = TRUE)
    }
  }
  .set_widths_smart(wb, st_fam, ncol(fam_tbl), nrow(fam_tbl) + 1L)
  openxlsx::freezePane(wb, st_fam, firstActiveRow = 2)
  openxlsx::addStyle(wb, st_fam, border_all_black,
                     rows = 1:(nrow(fam_tbl)+1), cols = 1:ncol(fam_tbl), gridExpand = TRUE, stack = TRUE)

  # ===== 3) CHOICES =====
  st_ch <- add_sheet("CHOICES")

  # completar list_name SIN crear .x/.y
  choices_out <- {
    ch <- plantilla$choices
    if (!is.null(inst)) {
      ch_map <- choices_es_tbl(inst) %>% dplyr::distinct(list_norm, list_name)
      # vectorizado
      map_vec <- ch_map$list_name[match(ch$list_norm, ch_map$list_norm)]
      ch$list_name <- dplyr::coalesce(ch$list_name, map_vec)
    }
    # asegurar labels y list_name
    ch$choice_label <- dplyr::coalesce(ch$choice_label, ch$choice_code)
    ch$list_name    <- ch$list_name %||% NA_character_
    ch
  }

  ch <- choices_out %>%
    dplyr::transmute(
      parent_col,
      list_name,
      tipo,
      variable_base,
      variable_label,
      code  = choice_code,
      label = choice_label
    )

  openxlsx::writeData(wb, st_ch, ch, startRow = 1, colNames = TRUE)

  if (nrow(ch) && "tipo" %in% names(ch)) {
    idx_by_tipo <- split(seq_len(nrow(ch)) + 1L, ch$tipo)
    for (tp in names(idx_by_tipo)) {
      openxlsx::addStyle(wb, st_ch,
                         openxlsx::createStyle(fgFill = tipo_hex(tp)),
                         rows = idx_by_tipo[[tp]], cols = 1:ncol(ch),
                         gridExpand = TRUE, stack = TRUE)
    }
  }
  .set_widths_smart(wb, st_ch, ncol(ch), nrow(ch) + 1L)
  openxlsx::freezePane(wb, st_ch, firstActiveRow = 2)
  openxlsx::addStyle(wb, st_ch, border_all_black,
                     rows = 1:(nrow(ch)+1), cols = 1:ncol(ch), gridExpand = TRUE, stack = TRUE)

  # ===== 4) DICCIONARIO =====
  st_dic <- add_sheet("DICCIONARIO")
  dic <- plantilla$diccionario

  if (!is.null(inst) && !is.null(inst$survey_raw)) {
    nms <- names(inst$survey_raw)
    idx <- which(grepl("^label::spanish", tolower(nms)))[1]
    if (!is.na(idx)) {
      col_span <- nms[idx]  # <- usa el nombre ORIGINAL, no el tolower
      s_min <- inst$survey_raw %>%
        dplyr::select(name, label_es_dic = !!rlang::sym(col_span))
      dic <- dic %>%
        dplyr::left_join(s_min, by = c("variable" = "name")) %>%
        dplyr::mutate(etiqueta = dplyr::coalesce(label_es_dic, etiqueta)) %>%
        dplyr::select(-label_es_dic)
    }
  }
  openxlsx::writeData(wb, st_dic, dic, startRow = 1, colNames = TRUE)
  if (nrow(dic) && "tipo_base" %in% names(dic)) {
    idx_by_tipo <- split(seq_len(nrow(dic)) + 1L, dic$tipo_base)
    for (tp in names(idx_by_tipo)) {
      openxlsx::addStyle(wb, st_dic,
                         openxlsx::createStyle(fgFill = tipo_hex(tp)),
                         rows = idx_by_tipo[[tp]], cols = 1:ncol(dic),
                         gridExpand = TRUE, stack = TRUE)
    }
  }
  .set_widths_smart(wb, st_dic, ncol(dic), nrow(dic) + 1L)
  openxlsx::freezePane(wb, st_dic, firstActiveRow = 2)
  openxlsx::addStyle(wb, st_dic, border_all_black,
                     rows = 1:(nrow(dic)+1), cols = 1:ncol(dic), gridExpand = TRUE, stack = TRUE)

  # ===== 5) Hojas por VARIABLE ==============================================
  orden_hojas <- plantilla$navegacion$hoja
  for (nm in orden_hojas) {
    df <- plantilla$sheets[[nm]]
    if (is.null(df) || !ncol(df)) next

    # reordenar recods al lado de base
    cn <- colnames(df)
    base_cols <- cn[!grepl("_recod$", cn)]
    rec_cols  <- cn[grepl("_recod$", cn)]
    order_cols <- unlist(lapply(base_cols, function(b){
      c(b, if (paste0(b,"_recod") %in% rec_cols) paste0(b,"_recod"))
    }))
    df <- df[, unique(c(order_cols, rec_cols)), drop = FALSE]

    # encabezados
    tipo_hoja <- attr(plantilla$sheets[[nm]], "tipo") %||% "text"
    hdr_base  <- colnames(df)
    hdr_lab   <- attr(df, "label_row") %||% hdr_base

    # Labels fijos en fila 2
    hdr_lab[hdr_base == "_uuid"]        <- "UUID"
    hdr_lab[hdr_base == "_index"]       <- "Índice"
    hdr_lab[hdr_base == "Código pulso"] <- "Código pulso"
    hdr_lab[grepl("_recod$", hdr_base)] <- "Recodificación"

    # Fila 1 (crudo/código) según tipo
    especiales <- c("_uuid","_index","Código pulso","Codigo pulso",
                    "Seleccionadas","Seleccionadas_cod")

    if (identical(tolower(tipo_hoja), "select_multiple")) {
      parent_col <- nm
      map_lab_code <- plantilla$choices %>%
        dplyr::filter(.data$parent_col == parent_col) %>%
        dplyr::distinct(choice_code, choice_label)

      map_no_recod <- function(cc){
        if (cc %in% especiales) return(cc)
        if (identical(cc, "Otro, por favor especificar")) {
          return(paste0(parent_col, "/Other"))
        }
        if (nrow(map_lab_code)) {
          i <- which(map_lab_code$choice_label == cc)[1]
          if (length(i) == 1 && !is.na(i)) {
            return(paste0(parent_col, "/", map_lab_code$choice_code[i]))
          }
        }
        cc
      }

      hdr_raw <- vapply(hdr_base, function(cc){
        if (grepl("_recod$", cc)) {
          base0 <- sub("_recod$", "", cc)
          paste0(map_no_recod(base0), "_recod")
        } else {
          map_no_recod(cc)
        }
      }, FUN.VALUE = character(1))

    } else if (identical(tolower(tipo_hoja), "select_one")) {
      parent_col <- nm

      map_no_recod <- function(cc){
        if (cc %in% c("_uuid","_index","Código pulso","Codigo pulso")) return(cc)
        if (cc == "Selección (código)") return(parent_col)
        if (cc == "Selección (label)")  return(paste0(parent_col, "_label"))
        cc
      }

      hdr_raw <- vapply(hdr_base, function(cc){
        if (grepl("_recod$", cc)) {
          base0 <- sub("_recod$", "", cc)
          paste0(map_no_recod(base0), "_recod")
        } else {
          map_no_recod(cc)
        }
      }, FUN.VALUE = character(1))

    } else {  # TEXT / INTEGER
      map_no_recod <- function(cc){
        if (cc %in% c("_uuid","_index","Código pulso","Codigo pulso")) return(cc)
        cc
      }
      hdr_raw <- vapply(hdr_base, function(cc){
        if (grepl("_recod$", cc)) {
          base0 <- sub("_recod$", "", cc)
          paste0(map_no_recod(base0), "_recod")
        } else {
          map_no_recod(cc)
        }
      }, FUN.VALUE = character(1))
    }

    # escribir hoja
    st <- add_sheet(nm)
    openxlsx::writeData(wb, st, t(hdr_raw), startRow = 1, colNames = FALSE)
    openxlsx::writeData(wb, st, t(hdr_lab), startRow = 2, colNames = FALSE)
    if (nrow(df)) openxlsx::writeData(wb, st, df, startRow = 3, colNames = FALSE, borders = "none")

    openxlsx::addStyle(wb, st, style_hdr1, rows = 1, cols = 1:ncol(df), gridExpand = TRUE)
    openxlsx::addStyle(wb, st, style_hdr2, rows = 2, cols = 1:ncol(df), gridExpand = TRUE)
    if (isTRUE(autofiltro)) openxlsx::addFilter(wb, st, rows = 2, cols = 1:ncol(df))

    id_cols <- intersect(c("_uuid","_index","Código pulso","Codigo pulso","pulso_code"), colnames(df))
    firstActiveCol <- if (length(id_cols)) (max(match(id_cols, colnames(df))) + 1L) else 1L
    if (isTRUE(congelar_encabezado)) openxlsx::freezePane(wb, st, firstActiveRow = 3, firstActiveCol = firstActiveCol)

    openxlsx::addStyle(wb, st, fill_tipo(tipo_hex(tipo_hoja)), rows = 1:2, cols = 1:ncol(df), gridExpand = TRUE, stack = TRUE)

    nrows <- max(2, nrow(df) + 2L)
    openxlsx::addStyle(wb, st, border_all_black, rows = 1:nrows, cols = 1:ncol(df), gridExpand = TRUE, stack = TRUE)

    openxlsx::setColWidths(wb, st, cols = 1:ncol(df), widths = "auto")
    openxlsx::addStyle(wb, st, wrap_left, rows = 3:(nrow(df)+3), cols = 1:ncol(df), gridExpand = TRUE, stack = TRUE)
  }

  openxlsx::saveWorkbook(wb, path_xlsx, overwrite = TRUE)
  invisible(path_xlsx)
}

