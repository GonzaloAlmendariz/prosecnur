# ============================================================
# MÓDULO: Carga de plan, evaluación y reporte HTML (G-aware + saneo regex)
# ============================================================

suppressPackageStartupMessages({
  library(dplyr)
  library(stringr)
  library(tibble)
})

`%||%` <- function(a, b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a

# -------------------------------------------------------------------
# 0) Helpers de normalización de texto y Sí/No robusto
# -------------------------------------------------------------------

# Normalizador mínimo de "Procesamiento" (por si viene con comillas raras)
normalizar_procesamiento <- function(x) {
  if (is.null(x)) return(x)
  x <- as.character(x)
  x <- gsub("\u201C|\u201D", "\"", x, perl = TRUE)  # “ ”
  x <- gsub("\u2018|\u2019", "'",  x, perl = TRUE)  # ‘ ’
  x <- gsub("[\u00A0\u2007\u202F]", " ", x, perl = TRUE) # espacios duros
  x <- gsub("(?<!<|>|!|<-|=)=(?!=)", "==", x, perl = TRUE) # "=" suelto -> "=="
  x <- gsub("={3,}", "==", x, perl = TRUE)

  # balanceo simple de paréntesis por si el plan trae alguno desbalanceado
  n_open  <- stringr::str_count(x, "\\(")
  n_close <- stringr::str_count(x, "\\)")
  need    <- n_open > n_close
  x[need] <- paste0(
    x[need],
    vapply((n_open[need] - n_close[need]), function(k) paste0(rep(")", k), collapse=""), "")
  )
  x
}

# “Sí/No” robusto en crudos: yes/Yes/si/Si/Sí/s/Y/1/true, etc. (por compatibilidad)
# Devuelve TRUE/FALSE/NA. (El plan ya trae %in% de YES/SÍ; esto queda disponible por si acaso.)
is_yes <- function(x) {
  if (is.logical(x)) return(x)
  if (is.numeric(x)) return(ifelse(is.na(x), NA, x != 0))
  s <- trimws(as.character(x))
  s <- iconv(s, from = "", to = "ASCII//TRANSLIT")   # quita acentos
  s <- tolower(s)
  s <- gsub("\\s+", " ", s)
  yes_set <- c("yes", "y", "si", "s", "1", "true", "verdadero")
  out <- ifelse(!nzchar(s), NA, s %in% yes_set)
  as.logical(out)
}

is_no <- function(x) {
  if (is.logical(x)) return(!x)
  if (is.numeric(x)) return(ifelse(is.na(x), NA, x == 0))
  s <- trimws(as.character(x))
  s <- iconv(s, from = "", to = "ASCII//TRANSLIT")
  s <- tolower(s)
  s <- gsub("\\s+", " ", s)
  no_set <- c("no", "n", "0", "false", "falso")
  out <- ifelse(!nzchar(s), NA, s %in% no_set)
  as.logical(out)
}

# Sanea expresiones del "Procesamiento" para regex con \s y grepl perl
sanear_procesamiento_regex <- function(x) {
  if (is.null(x) || is.na(x) || !nzchar(x)) return(x)
  out <- as.character(x)

  # Asegurar que \s dentro de cadenas quede como \\s (escape correcto en R)
  # (Si ya está \\s, NO lo vuelve a duplicar gracias al lookbehind.)
  out <- gsub("(?<!\\\\)\\\\s", "\\\\\\\\s", out, perl = TRUE)

  out
}

# -------------------------------------------------------------------
# 1) Cargar plan desde Excel (hoja "Plan")
#    Acepta Variable 3 si está presente (opcional)
# -------------------------------------------------------------------
#' @export
cargar_plan_excel <- function(path, sheet = "Plan") {
  stopifnot(file.exists(path))
  plan <- readxl::read_excel(path, sheet = sheet)

  req <- c("ID","Tipo de observación","Objetivo",
           "Variable 1","Variable 1 - Etiqueta",
           "Variable 2","Variable 2 - Etiqueta",
           "Nombre de regla","Procesamiento")
  faltan <- setdiff(req, names(plan))
  if (length(faltan)) {
    stop("La hoja '", sheet, "' carece de columnas: ",
         paste(faltan, collapse = ", "), call. = FALSE)
  }

  # Columnas opcionales para variable 3 (si el plan las trae)
  if (!"Variable 3" %in% names(plan)) plan[["Variable 3"]] <- NA_character_
  if (!"Variable 3 - Etiqueta" %in% names(plan)) plan[["Variable 3 - Etiqueta"]] <- NA_character_

  plan[["Procesamiento"]] <- vapply(plan[["Procesamiento"]],
                                    function(z) sanear_procesamiento_regex(normalizar_procesamiento(z)),
                                    character(1))
  plan
}

# -------------------------------------------------------------------
# 2) (OBSOLETA) Reescrituras ODK → R
#     Se deja por compatibilidad pero NO se usa (el plan ya viene listo).
# -------------------------------------------------------------------
reescribir_odk_fix <- function(rhs, var1 = NULL) rhs

# -------------------------------------------------------------------
# 3) Evaluación de consistencia (TRUE = OK; FALSE = inconsistencia)
#     - Respeta Variable 1/2/3 en meta
#     - is_yes()/is_no() disponibles en el entorno de evaluación
# -------------------------------------------------------------------
#' @export
evaluar_consistencia <- function(datos,
                                 plan,
                                 excluir_ids = NULL,
                                 contar_na_como_inconsistencia = FALSE) {
  stopifnot(is.data.frame(datos), is.data.frame(plan))

  safe_col <- function(df, col, type = "character") {
    if (col %in% names(df)) df[[col]] else {
      if (type == "character") rep(NA_character_, nrow(df))
      else if (type == "numeric") rep(NA_real_, nrow(df))
      else if (type == "integer") rep(NA_integer_, nrow(df))
      else rep(NA, nrow(df))
    }
  }

  req <- c("ID","Nombre de regla","Procesamiento")
  if (!all(req %in% names(plan))) {
    stop("El plan no tiene columnas requeridas: ", paste(req, collapse = ", "))
  }

  plan2 <- tibble(
    id_regla            = plan[["ID"]],
    nombre_regla        = plan[["Nombre de regla"]],
    procesamiento       = plan[["Procesamiento"]],
    objetivo            = safe_col(plan, "Objetivo", "character"),
    tipo_observacion    = safe_col(plan, "Tipo de observación", "character"),
    variable_1          = safe_col(plan, "Variable 1", "character"),
    variable_1_etiqueta = safe_col(plan, "Variable 1 - Etiqueta", "character"),
    variable_2          = safe_col(plan, "Variable 2", "character"),
    variable_2_etiqueta = safe_col(plan, "Variable 2 - Etiqueta", "character"),
    variable_3          = safe_col(plan, "Variable 3", "character"),
    variable_3_etiqueta = safe_col(plan, "Variable 3 - Etiqueta", "character")
  )

  if (!is.null(excluir_ids) && length(excluir_ids)) {
    plan2 <- dplyr::filter(plan2, !.data$id_regla %in% excluir_ids)
  }
  if (nrow(plan2) == 0L) {
    return(list(datos = datos, resumen = tibble(), reglas_meta = tibble()))
  }

  .parsear_regla <- function(proc, nombre_fallback = NA_character_) {
    if (is.na(proc) || !nzchar(proc)) return(NULL)
    if (grepl("<-", proc, fixed = TRUE)) {
      partes <- strsplit(proc, "<-", fixed = TRUE)[[1]]
      flag <- trimws(partes[1]); rhs <- trimws(paste(partes[-1], collapse = "<-"))
    } else {
      if (is.na(nombre_fallback) || !nzchar(nombre_fallback)) return(NULL)
      flag <- nombre_fallback; rhs <- trimws(proc)
    }
    if (!nzchar(flag) || !nzchar(rhs)) return(NULL)
    list(flag = flag, rhs = rhs)
  }

  base_ok <- if (contar_na_como_inconsistencia) FALSE else TRUE
  df <- datos
  resumen <- vector("list", nrow(plan2))


  # entorno seguro con utilidades disponibles dentro de eval()
  grepl_perl_default <- function(pattern, x, ..., perl = TRUE) {
    base::grepl(pattern, x, ..., perl = perl)
  }

  eval_env <- rlang::env(
    rlang::base_env(),   # <- antes estaba rlang::empty_env()
    is_yes = is_yes,
    is_no  = is_no,
    grepl  = grepl_perl_default
  )

  # expón las columnas de df como símbolos dentro del entorno
  for (nm in names(df)) rlang::env_bind(eval_env, !!nm := df[[nm]])

  for (i in seq_len(nrow(plan2))) {
    r <- plan2[i, ]
    pr <- .parsear_regla(sanear_procesamiento_regex(normalizar_procesamiento(r$procesamiento)),
                         r$nombre_regla)
    if (is.null(pr)) next

    rhs2 <- pr$rhs
    ok <- TRUE
    expr_parsed <- tryCatch(rlang::parse_expr(rhs2), error = function(e){ ok <<- FALSE; e })
    if (!ok) {
      warning(sprintf("No se pudo parsear la regla %s (%s): %s", r$id_regla, r$nombre_regla, rhs2))
      next
    }

    # evalúa y añade flag a df
    vals <- tryCatch(rlang::eval_bare(expr_parsed, env = eval_env),
                     error = function(e) { warning(sprintf("Error evaluando %s: %s", r$nombre_regla, e$message)); rep(NA, nrow(df)) })
    if (!is.logical(vals)) {
      warning(sprintf("La evaluación de %s no devolvió lógico. Marcando NA.", r$nombre_regla))
      vals <- rep(NA, nrow(df))
    }
    df[[pr$flag]] <- vals
    # mantener también en el env para posibles dependencias posteriores
    rlang::env_bind(eval_env, !!pr$flag := vals)

    n_inc <- sum(!dplyr::coalesce(vals, base_ok), na.rm = TRUE)

    resumen[[i]] <- tibble(
      id_regla            = r$id_regla,
      nombre_regla        = r$nombre_regla,
      objetivo            = r$objetivo,
      tipo_observacion    = r$tipo_observacion,
      variable_1          = r$variable_1,
      variable_1_etiqueta = r$variable_1_etiqueta,
      variable_2          = r$variable_2,
      variable_2_etiqueta = r$variable_2_etiqueta,
      variable_3          = r$variable_3,
      variable_3_etiqueta = r$variable_3_etiqueta,
      flag                = pr$flag,
      n_inconsistencias   = n_inc
    )
  }

  resumen <- dplyr::bind_rows(resumen) |>
    dplyr::mutate(porcentaje = if (nrow(df) > 0) n_inconsistencias/nrow(df) else NA_real_) |>
    dplyr::arrange(dplyr::desc(n_inconsistencias))

  reglas_meta <- dplyr::select(
    plan2,
    id_regla, nombre_regla, objetivo, tipo_observacion,
    variable_1, variable_1_etiqueta,
    variable_2, variable_2_etiqueta,
    variable_3, variable_3_etiqueta,
    procesamiento
  )

  list(datos = df, resumen = resumen, reglas_meta = reglas_meta)
}

# -------------------------------------------------------------------
# 4) Atajo: evaluar cargando el plan directo desde Excel
# -------------------------------------------------------------------
#' @export
evaluar_desde_excel <- function(datos,
                                path_xlsx,
                                excluir_ids = NULL,
                                contar_na_como_inconsistencia = FALSE) {
  plan <- cargar_plan_excel(path_xlsx, sheet = "Plan")
  evaluar_consistencia(
    datos = datos,
    plan = plan,
    excluir_ids = excluir_ids,
    contar_na_como_inconsistencia = contar_na_como_inconsistencia
  )
}

# -------------------------------------------------------------------
# 5) Casos por regla (solo _uuid, código pulso y variables de la regla)
#     – ahora incluye variable_3 si existe
# -------------------------------------------------------------------
#' @export
observaciones_regla <- function(evaluacion,
                                regla,
                                por = c("id_regla","nombre_regla","flag"),
                                incluir_flag = FALSE,
                                contar_na_como_inconsistencia = FALSE) {
  por <- match.arg(por)
  stopifnot(is.list(evaluacion),
            all(c("datos","resumen","reglas_meta") %in% names(evaluacion)))

  datos <- evaluacion$datos
  res   <- evaluacion$resumen
  meta  <- evaluacion$reglas_meta
  if (!nrow(res)) return(tibble())

  idx_res <- switch(
    por,
    id_regla     = which(res$id_regla     %in% regla),
    nombre_regla = which(res$nombre_regla %in% regla),
    flag         = which(res$flag         %in% regla)
  )
  if (!length(idx_res)) return(tibble())
  i <- idx_res[[1L]]

  id_regla  <- res$id_regla[i]
  flag_name <- res$flag[i]

  mrow <- meta[match(id_regla, meta$id_regla), , drop = FALSE]
  var1 <- mrow$variable_1 %||% NA_character_
  var2 <- mrow$variable_2 %||% NA_character_
  var3 <- mrow$variable_3 %||% NA_character_

  if (!nzchar(flag_name) || !flag_name %in% names(datos)) return(tibble())
  base_ok <- if (isTRUE(contar_na_como_inconsistencia)) FALSE else TRUE
  vals    <- datos[[flag_name]]
  inc     <- !dplyr::coalesce(vals, base_ok)
  if (!any(inc, na.rm = TRUE)) return(tibble())

  # Aliases de código pulso
  codigo_alias <- c("Codigo pulso","Código pulso","Codigo_pulso","codigo_pulso","codigo.pulso")
  codigo_col <- codigo_alias[codigo_alias %in% names(datos)][1] %||% NULL

  cols <- c("_uuid", "_id", codigo_col, var1, var2, var3)
  cols <- unique(na.omit(cols))
  cols <- cols[cols %in% names(datos)]

  out <- tibble::as_tibble(datos[inc, cols, drop = FALSE])
  if (isTRUE(incluir_flag)) out[[flag_name]] <- datos[[flag_name]][inc]
  out
}

# -------------------------------------------------------------------
# 6) Total de inconsistencias (tibble o kable HTML)
# -------------------------------------------------------------------
#' @export
total_inconsistencias <- function(evaluacion,
                                  as_kable = FALSE,
                                  titulo = "Totales de inconsistencia por regla") {
  stopifnot(all(c("datos","resumen") %in% names(evaluacion)))
  res <- evaluacion$resumen
  if (!nrow(res)) {
    tb <- tibble(Reglas = 0L, Total_inconsistencias = 0L)
    if (!as_kable) return(tb)
    if (!requireNamespace("kableExtra", quietly = TRUE)) return(tb)
    return(
      knitr::kable(tb, format = "html", caption = titulo) |>
        kableExtra::kable_styling(full_width = FALSE)
    )
  }
  tb <- dplyr::transmute(
    res,
    id_regla, nombre_regla, n_inconsistencias,
    porcentaje = round(100*porcentaje, 1)
  )
  tot <- sum(res$n_inconsistencias %||% 0L, na.rm = TRUE)
  cab <- tibble(Reglas = nrow(res), Total_inconsistencias = tot)

  if (!as_kable) return(list(cabecera = cab, detalle = tb))

  if (!requireNamespace("kableExtra", quietly = TRUE)) return(list(cabecera = cab, detalle = tb))
  list(
    cabecera = knitr::kable(cab, format = "html", caption = titulo) |>
      kableExtra::kable_styling(full_width = FALSE),
    detalle = knitr::kable(tb, format = "html") |>
      kableExtra::kable_styling(full_width = FALSE, bootstrap_options = c("striped","condensed"))
  )
}

# -------------------------------------------------------------------
# 7) Agrupar info por regla (bloques) para el reporte
# -------------------------------------------------------------------
#' @export
reporte_bloques <- function(evaluacion,
                            familias = NULL,
                            ids = NULL,
                            solo_relevantes = FALSE,
                            incluir_solo_inconsistentes = FALSE,
                            incluir_reglas_sin_casos = TRUE,
                            contar_na_como_inconsistencia = FALSE,
                            ordenar = c("n_desc","id_asc")) {
  ordenar <- match.arg(ordenar)
  res  <- evaluacion$resumen
  meta <- evaluacion$reglas_meta
  if (!nrow(res)) return(tibble())

  # filtrar por familia/ids opcionales
  if (!is.null(familias) && length(familias)) {
    keep <- vapply(res$tipo_observacion, function(x) any(stringr::str_detect(x %||% "", paste(familias, collapse = "|"))), TRUE)
    res <- res[keep, , drop = FALSE]
  }
  if (!is.null(ids) && length(ids)) {
    res <- res[res$id_regla %in% ids | res$nombre_regla %in% ids, , drop = FALSE]
  }
  if (solo_relevantes) {
    res <- res[res$n_inconsistencias > 0, , drop = FALSE]
  }

  # armar tibble de bloques
  join_cols <- c("id_regla","nombre_regla")
  bloques <- dplyr::left_join(res, meta, by = join_cols, suffix = c("", ""))
  stopifnot("procesamiento" %in% names(bloques))

  # anexar casos (lista-columna)
  bloques$casos <- lapply(seq_len(nrow(bloques)), function(i) {
    if (!incluir_solo_inconsistentes && bloques$n_inconsistencias[i] == 0) return(tibble())
    observaciones_regla(
      evaluacion = evaluacion,
      regla = bloques$id_regla[i],
      por = "id_regla",
      incluir_flag = FALSE,
      contar_na_como_inconsistencia = contar_na_como_inconsistencia
    )
  })

  if (!incluir_reglas_sin_casos) {
    keep <- vapply(bloques$casos, function(df) nrow(df) > 0, TRUE)
    bloques <- bloques[keep, , drop = FALSE]
  }

  if (ordenar == "n_desc") {
    bloques <- dplyr::arrange(bloques, dplyr::desc(.data$n_inconsistencias))
  } else {
    bloques <- dplyr::arrange(bloques, .data$id_regla)
  }

  bloques
}


# --- helpers locales de sanitización y utilidades --------------------
`%||%` <- function(a, b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a
.nz <- function(x) is.character(x) && length(x)==1 && !is.na(x) && nzchar(trimws(x))

.san_label <- function(s) {
  s <- as.character(s %||% "")
  s <- gsub("\\*\\*", "", s, perl = TRUE)                     # quita **markdown**
  s <- gsub("\\$\\{([A-Za-z0-9_]+)\\}", "«\\1»", s, perl=TRUE) # ${var} -> «var»
  s <- gsub("[\\r\\n]+", " ", s, perl = TRUE)                 # sin saltos
  trimws(gsub("\\s+", " ", s, perl = TRUE))
}

.safe_html <- function(x) htmltools::htmlEscape(as.character(x %||% ""))

# -------------------------------------------------------------------
# 8) Renderizador kable para bloques (solo Var1/2/3 no-NA + Procesamiento)
# -------------------------------------------------------------------
#' @export
render_bloques_kable <- function(
    bloques,
    max_casos = 10,
    mostrar_aleatorios_si_cero = 0,
    fallback_df = NULL,
    cols_id = c("_uuid","_index","Codigo pulso","Código pulso"),
    cols_interes = NULL,
    seed = NULL,
    alto_resumen = "280px",
    ancho_resumen = "100%",
    alto_casos = "420px",
    ancho_casos = "100%",
    map_clean_to_original = NULL,
    usar_mapeo = FALSE,
    quitar_vacias_en_casos = FALSE
){
  `%||%` <- function(a,b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a
  nz  <- function(x) is.character(x) && length(x)==1 && !is.na(x) && nzchar(trimws(x))
  esc <- function(x) htmltools::htmlEscape(as.character(x %||% ""))

  .san_label <- function(s){
    s <- as.character(s %||% "")
    s <- gsub("\\*\\*", "", s, perl = TRUE)
    s <- gsub("\\$\\{([A-Za-z0-9_]+)\\}", "«\\1»", s, perl=TRUE)
    s <- gsub("[\\r\\n]+", " ", s, perl = TRUE)
    trimws(gsub("\\s+", " ", s, perl = TRUE))
  }
  .exact_cols <- function(df, wanted){
    if (is.null(df) || !length(wanted)) return(character(0))
    intersect(as.character(wanted), names(df))
  }
  .drop_all_na_cols <- function(df){
    if (is.null(df) || !nrow(df) || !ncol(df)) return(df)
    keep <- vapply(df, function(col) any(!is.na(col) & trimws(as.character(col)) != ""), logical(1))
    df[, keep, drop = FALSE]
  }
  maybe_map <- function(vars){
    if (!usar_mapeo || is.null(map_clean_to_original)) return(vars)
    if (!all(c("clean","original") %in% names(map_clean_to_original))) return(vars)
    vapply(vars, function(v){
      if (is.na(v) || !nzchar(v)) return(v)
      hit <- map_clean_to_original$original[match(v, map_clean_to_original$clean)]
      ifelse(!is.na(hit), hit, v)
    }, FUN.VALUE = character(1))
  }

  if (is.null(bloques) || !nrow(bloques)) {
    cat("<p><em>No hay reglas para mostrar.</em></p>")
    return(invisible(NULL))
  }
  if (!requireNamespace("kableExtra", quietly = TRUE)) stop("Necesitas {kableExtra}.")
  if (!is.null(seed)) set.seed(seed)

  html_nodes <- list()

  for (i in seq_len(nrow(bloques))) {
    b <- bloques[i,]

    # ===== Encabezado: SOLO ID =====
    cab <- paste0("[", b$id_regla %||% "", "]")

    objetivo <- .san_label(b$objetivo)
    tipo     <- .san_label(b$tipo_observacion)

    header_html <- htmltools::tagList(
      htmltools::tags$div(style="margin:18px 0;border-top:1px solid #ddd;padding-top:14px"),
      htmltools::tags$h3(esc(cab)),
      if (nz(objetivo)) htmltools::tags$div(
        htmltools::tags$strong("Objetivo:"), " ", esc(objetivo)
      ),
      if (nz(tipo)) htmltools::tags$div(
        htmltools::tags$strong("Tipo:"), " ", esc(tipo)
      )
    )

    # ===== Resumen =====
    lab1 <- .san_label(b$variable_1_etiqueta)
    lab2 <- .san_label(b$variable_2_etiqueta)
    lab3 <- .san_label(b$variable_3_etiqueta)

    resumen_raw <- tibble::tibble(
      `Id regla`   = b$id_regla %||% NA_character_,
      `Variable 1` = b$variable_1 %||% NA_character_,
      `Etiqueta 1` = if (nz(lab1)) lab1 else NA_character_,
      `Variable 2` = b$variable_2 %||% NA_character_,
      `Etiqueta 2` = if (nz(lab2)) lab2 else NA_character_,
      `Variable 3` = b$variable_3 %||% NA_character_,
      `Etiqueta 3` = if (nz(lab3)) lab3 else NA_character_,
      `# inconsist.` = b$n_inconsistencias %||% 0L,
      `%`           = if (!is.na(b$porcentaje)) sprintf("%.1f%%", 100*b$porcentaje) else NA_character_
    )

    # 🔍 Ocultar Var2/Etq2 y Var3/Etq3 si son NA / "" / "NA" (texto)
    blankish <- function(v) {
      v0 <- as.character(v)
      is.na(v0) | trimws(v0) == "" | trimws(v0) == "NA"
    }
    cols_chk  <- c("Variable 2","Etiqueta 2","Variable 3","Etiqueta 3")
    drop_cols <- vapply(resumen_raw[cols_chk], function(col) all(blankish(col)), logical(1))
    keep_cols <- c(setdiff(names(resumen_raw), cols_chk[drop_cols]))
    resumen_df <- resumen_raw[, keep_cols, drop = FALSE]

    k_res <- knitr::kable(resumen_df, format = "html", align = "l", escape = TRUE) |>
      kableExtra::kable_styling(full_width = FALSE,
                                bootstrap_options = c("striped","condensed"),
                                font_size = 13) |>
      kableExtra::scroll_box(
        height = alto_resumen, width = ancho_resumen,
        box_css = "overflow-x:auto; overflow-y:auto; border:1px solid #ddd; padding:6px; border-radius:6px;"
      )

    # ===== Casos =====
    casos <- b$casos[[1]]
    vars_regla <- unique(na.omit(c(b$variable_1, b$variable_2, b$variable_3)))
    vars_regla <- vars_regla[nzchar(vars_regla)]
    vars_regla <- maybe_map(vars_regla)

    var_labels <- paste0(vars_regla, "_label")
    desired_cols <- unique(c(cols_id, as.vector(rbind(vars_regla, var_labels)), cols_interes))
    desired_cols <- desired_cols[!is.na(desired_cols) & nzchar(desired_cols)]

    usar_fallback <- (is.null(casos) || !nrow(casos)) && mostrar_aleatorios_si_cero > 0 &&
      !is.null(fallback_df) && nrow(fallback_df) > 0

    if (usar_fallback) {
      n_take <- min(mostrar_aleatorios_si_cero, nrow(fallback_df))
      idx <- sample.int(nrow(fallback_df), n_take)
      keep <- .exact_cols(fallback_df, desired_cols)
      if (!length(keep)) keep <- .exact_cols(fallback_df, cols_id)
      if (!length(keep)) keep <- head(names(fallback_df), 6)
      casos <- tibble::as_tibble(fallback_df[idx, keep, drop = FALSE])
    } else if (!is.null(casos) && nrow(casos)) {
      keep <- .exact_cols(casos, desired_cols)
      if (!length(keep)) keep <- .exact_cols(casos, c(cols_id, vars_regla))
      if (!length(keep)) keep <- .exact_cols(casos, cols_id)
      if (!length(keep)) keep <- head(names(casos), 6)
      casos <- casos[, keep, drop = FALSE]
    }

    if (isTRUE(quitar_vacias_en_casos)) {
      casos <- .drop_all_na_cols(casos)
    }

    casos_node <- if (!is.null(casos) && nrow(casos) && ncol(casos)) {
      casos_show <- utils::head(casos, max_casos)
      k_casos <- knitr::kable(casos_show, format = "html", align = "l", escape = TRUE) |>
        kableExtra::kable_styling(full_width = FALSE, bootstrap_options = c("hover","condensed"), font_size = 12) |>
        kableExtra::scroll_box(
          height = alto_casos, width = ancho_casos,
          box_css = "overflow-x:auto; overflow-y:auto; border:1px solid #ddd; padding:6px; border-radius:6px;"
        )
      htmltools::HTML(as.character(k_casos))
    } else {
      htmltools::tags$div(style="margin-top:6px;color:#666;", htmltools::tags$em("Casos: (ninguno)"))
    }

    # ===== Procesamiento =====
    proc <- unname(as.character(b$procesamiento %||% ""))
    proc_node <- if (nz(proc)) {
      htmltools::tags$div(
        style="margin-top:8px;",
        htmltools::tags$strong("Procesamiento utilizado en R:"),
        htmltools::tags$pre(
          style="white-space:pre-wrap; border:1px solid #ddd; padding:8px; border-radius:6px; background:#fafafa; font-family: ui-monospace, Menlo, Consolas, 'Liberation Mono', monospace;",
          htmltools::HTML(proc)  # ⬅️ sin escape, muestra "<-", "&&", "<=", etc.
        )
      )
    } else htmltools::HTML("")

    html_nodes[[length(html_nodes)+1]] <- htmltools::tagList(
      htmltools::tags$div(style="margin:18px 0;border-top:1px solid #ddd;padding-top:14px"),
      header_html,
      htmltools::HTML(as.character(k_res)),
      casos_node,
      proc_node
    )
  }

  out <- htmltools::tagList(html_nodes)
  cat(as.character(out))
  invisible(out)
}

# -------------------------------------------------------------------
# 9) Reporte HTML completo (usa los mismos nodos HTML seguros)
# -------------------------------------------------------------------
#' @export
reporte_html_limpieza <- function(evaluacion,
                                  path_html,
                                  familias = NULL,
                                  ids = NULL,
                                  solo_relevantes = FALSE,
                                  incluir_solo_inconsistentes = FALSE,
                                  max_casos = 20,
                                  mostrar_aleatorios_si_cero = 0,
                                  fallback_df = NULL,
                                  cols_id = c("_uuid","_index","Codigo pulso","Código pulso"),
                                  cols_interes = NULL,
                                  seed = NULL,
                                  titulo = "Reporte de limpieza – ACNUR",
                                  autor  = NULL,
                                  section_map = NULL,
                                  ancho_contenido = "1100px") {

  stopifnot(is.list(evaluacion), is.character(path_html), length(path_html)==1)
  if (!requireNamespace("kableExtra", quietly = TRUE)) stop("Necesitas {kableExtra}.")
  if (!requireNamespace("htmltools", quietly = TRUE)) stop("Necesitas {htmltools}.")

  # Totales
  tot <- total_inconsistencias(evaluacion, as_kable = FALSE)
  if (is.list(tot)) {
    cab <- tot$cabecera
    det <- tot$detalle
  } else {
    cab <- tot
    det <- tibble::tibble()
  }

  k_cab <- knitr::kable(cab, format = "html", caption = NULL) |>
    kableExtra::kable_styling(full_width = FALSE, bootstrap_options = c("condensed"))
  k_det <- if (nrow(det)) {
    knitr::kable(det, format = "html", caption = "Inconsistencias por regla") |>
      kableExtra::kable_styling(full_width = FALSE, bootstrap_options = c("striped","condensed")) |>
      kableExtra::scroll_box(width = "100%", height = "260px",
                             box_css = "overflow:auto;border:1px solid #ddd;border-radius:6px;padding:6px;")
  } else {
    htmltools::HTML("<em>No hay detalle de reglas.</em>")
  }

  # Bloques
  bloques <- reporte_bloques(
    evaluacion = evaluacion,
    familias = familias,
    ids = ids,
    solo_relevantes = solo_relevantes,
    incluir_solo_inconsistentes = incluir_solo_inconsistentes,
    incluir_reglas_sin_casos = TRUE,
    contar_na_como_inconsistencia = FALSE,
    ordenar = "n_desc"
  )

  # Generar los nodos de bloques con la función de arriba
  bloques_nodes <- render_bloques_kable(
    bloques = bloques,
    max_casos = max_casos,
    mostrar_aleatorios_si_cero = mostrar_aleatorios_si_cero,
    fallback_df = fallback_df %||% evaluacion$datos,
    cols_id = cols_id,
    cols_interes = cols_interes,
    seed = seed,
    alto_resumen = "220px",
    ancho_resumen = "100%",
    alto_casos   = "360px",
    ancho_casos  = "100%"
  )

  # Armar HTML
  css <- htmltools::HTML(sprintf("
  <style>
    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, 'Helvetica Neue', Arial, 'Noto Sans', 'Liberation Sans', sans-serif; }
    .wrap { max-width: %s; margin: 24px auto; padding: 6px 14px; }
    h1 { font-size: 1.6rem; margin: 0 0 6px 0; }
    .meta { color:#666; margin-bottom: 14px; }
    h2 { font-size: 1.2rem; margin-top: 20px; }
    h3 { font-size: 1.05rem; }
  </style>", ancho_contenido))

  head <- htmltools::tags$head(css)
  header <- htmltools::div(class="wrap",
                           htmltools::tags$h1(.safe_html(titulo)),
                           htmltools::div(class="meta",
                                          .safe_html(paste("Generado:", format(Sys.time(), "%Y-%m-%d %H:%M"),
                                                           if (!is.null(autor) && nzchar(autor)) paste("—", autor) else "")))
  )

  cuerpo <- htmltools::div(class="wrap",
                           htmltools::tags$h2("Totales"),
                           htmltools::HTML(as.character(k_cab)),
                           htmltools::tags$h2("Inconsistencias por regla"),
                           if (inherits(k_det, "shiny.tag")) k_det else htmltools::HTML(as.character(k_det)),
                           htmltools::tags$h2("Detalle por regla"),
                           if (inherits(bloques_nodes, "shiny.tag.list"))
                             bloques_nodes
                           else htmltools::HTML(as.character(bloques_nodes))
  )

  html <- htmltools::tagList(head, htmltools::tags$body(header, cuerpo))
  dir.create(dirname(path_html), recursive = TRUE, showWarnings = FALSE)
  htmltools::save_html(html, file = path_html, background = "white")
  invisible(path_html)
}



