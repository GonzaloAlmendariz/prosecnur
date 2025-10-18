# ============================================================
# MÓDULO: Carga de plan, evaluación multi-tabla y reporte HTML
#          (robusto a estudios con y sin repeats)
# ============================================================

suppressPackageStartupMessages({
  library(dplyr)
  library(stringr)
  library(tibble)
  library(rlang)
})

#' @title Evaluación de reglas de limpieza (multi-tabla)
#' @description
#' Evalúa un plan de limpieza sobre:
#' - un `data.frame` (estudios sin repeats),
#' - una **lista** devuelta por `lector_limpieza()` (estudios con repeats), o
#' - una **ruta a .xlsx** (se usa una hoja principal).
#'
#' Usa la columna **Tabla** del plan para elegir el `data.frame` correcto.
#' Devuelve las tablas con flags añadidos y un resumen por regla.
#'
#' @return lista con:
#' \itemize{
#' \item \code{datos}: data.frame principal (compatibilidad hacia atrás)
#' \item \code{datos_tablas}: lista nombrada de data.frames (principal y repeats)
#' \item \code{resumen}: tibble de conteo de inconsistencias por regla
#' \item \code{reglas_meta}: metadatos de reglas (incluye Tabla, Sección, Tipo, Procesamiento)
#' }
#' @export

`%||%` <- function(a, b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a
.nz     <- function(x) is.character(x) && length(x)==1 && !is.na(x) && nzchar(trimws(x))

# --- ALIAS DE VARIABLES: enlaza Intro04 ~ intro04, MMR01 ~ mmr01, etc. ----
.var_aliases <- function(nm) {
  # genera claves alternativas para cada nombre de columna
  raw <- nm
  lo  <- tolower(nm)
  up  <- toupper(nm)
  # versión "simple": solo alfanumérico y "_"
  simp <- function(x){
    x <- gsub("[^A-Za-z0-9_]", "", x)
    x
  }
  data.frame(
    alias = unique(c(raw, lo, up, simp(raw), simp(lo), simp(up))),
    stringsAsFactors = FALSE
  )$alias
}

.bind_vars_with_aliases <- function(env, df) {
  nms <- names(df)
  # ligamos cada columna en el entorno con varias claves: original, lower, upper, "simple"
  for (nm in nms) {
    vals <- df[[nm]]
    keys <- unique(c(nm, tolower(nm), toupper(nm),
                     gsub("[^A-Za-z0-9_]", "", nm),
                     gsub("[^A-Za-z0-9_]", "", tolower(nm)),
                     gsub("[^A-Za-z0-9_]", "", toupper(nm))))
    for (k in keys) {
      # si ya existe una ligadura con el mismo nombre y es idéntica, no hace falta re-ligar
      rlang::env_bind(env, !!k := vals)
    }
  }
  env
}


# -------------------------------------------------------------------
# 0) Helpers de comparación NA-segura y Sí/No
# -------------------------------------------------------------------
.eq_num_na <- function(a, b){
  aa <- suppressWarnings(as.numeric(a))
  bb <- suppressWarnings(as.numeric(b))
  ( (is.na(aa) & is.na(bb)) | (aa == bb) )
}
.eq_chr_na <- function(a, b){
  aa <- as.character(a); bb <- as.character(b)
  ( (is.na(aa) & is.na(bb)) | (aa == bb) )
}
eq_num_na <- .eq_num_na
eq_chr_na <- .eq_chr_na

is_yes <- function(x) {
  if (is.logical(x)) return(x)
  if (is.numeric(x)) return(ifelse(is.na(x), NA, x != 0))
  s <- trimws(as.character(x))
  s <- iconv(s, from = "", to = "ASCII//TRANSLIT")
  s <- tolower(s)
  s <- gsub("\\s+", " ", s)
  yes_set <- c("yes","y","si","s","1","true","verdadero")
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
  no_set <- c("no","n","0","false","falso")
  out <- ifelse(!nzchar(s), NA, s %in% no_set)
  as.logical(out)
}

# grepl por defecto con perl=TRUE (para reglas con \\s, lookarounds, etc.)
.grepl_perl <- function(pattern, x, ..., perl = TRUE) base::grepl(pattern, x, ..., perl = perl)

# Normalizador fuerte de nombres de tablas (quita acentos, (principal), espacios raros, símbolos)
.norm_tab <- function(s){
  s <- as.character(s %||% "")
  s <- tolower(trimws(s))
  s <- gsub("^\\(|\\)$", "", s)               # "(principal)" -> "principal"
  s <- gsub("[\u00A0\u2007\u202F]", " ", s)   # espacios duros -> " "
  s <- gsub("\\s+", "_", s)                   # espacios -> "_"
  s <- iconv(s, to = "ASCII//TRANSLIT")       # perú -> peru
  gsub("[^a-z0-9_]", "", s, perl = TRUE)      # limpia símbolos
}

# -------------------------------------------------------------------
# 1) Normalizadores de “Procesamiento”
# -------------------------------------------------------------------
.normalizar_procesamiento <- function(x) {
  if (is.null(x)) return(x)
  x <- as.character(x)
  x <- gsub("\u201C|\u201D", "\"", x, perl = TRUE)  # “ ”
  x <- gsub("\u2018|\u2019", "'",  x, perl = TRUE)  # ‘ ’
  x <- gsub("[\u00A0\u2007\u202F]", " ", x, perl = TRUE) # espacios duros
  x <- gsub("(?<!<|>|!|<-|=)=(?!=)", "==", x, perl = TRUE) # "=" suelto -> "=="
  x <- gsub("={3,}", "==", x, perl = TRUE)
  # balance rápido de paréntesis
  n_open  <- stringr::str_count(x, "\\(")
  n_close <- stringr::str_count(x, "\\)")
  need    <- n_open > n_close
  x[need] <- paste0(
    x[need],
    vapply((n_open[need] - n_close[need]), function(k) paste0(rep(")", k), collapse=""), "")
  )
  x
}

.sanear_regex_en_procesamiento <- function(x) {
  if (is.null(x) || is.na(x) || !nzchar(x)) return(x)
  out <- as.character(x)
  # asegurar \\s literal
  out <- gsub("(?<!\\\\)\\\\s", "\\\\\\\\s", out, perl = TRUE)
  out
}

# -------------------------------------------------------------------
# 2) Cargar plan con TUS columnas (incluye “Tabla”)
# -------------------------------------------------------------------
#' @title Cargar plan de limpieza desde Excel
#' @param path ruta al .xlsx
#' @param sheet nombre de hoja (default "Plan")
#' @export
cargar_plan_excel <- function(path, sheet = "Plan"){
  stopifnot(file.exists(path))
  plan <- readxl::read_excel(path, sheet = sheet)

  # normalizar encabezados (espacios raros → espacio simple)
  nms_raw <- names(plan)
  nms <- gsub("[   ]+", " ", trimws(nms_raw), perl = TRUE)
  names(plan) <- nms

  req <- c("ID","Tabla","Sección","Categoría","Tipo",
           "Nombre de regla","Objetivo",
           "Variable 1","Variable 1 - Etiqueta",
           "Variable 2","Variable 2 - Etiqueta",
           "Variable 3","Variable 3 - Etiqueta",
           "Procesamiento")
  faltan <- setdiff(req, names(plan))
  if (length(faltan)) {
    stop("La hoja '", sheet, "' carece de columnas: ",
         paste(faltan, collapse = ", "))
  }

  plan[["Procesamiento"]] <- vapply(
    plan[["Procesamiento"]],
    function(z) .sanear_regex_en_procesamiento(.normalizar_procesamiento(z)),
    character(1)
  )

  plan
}

# -------------------------------------------------------------------
# 3) Resolver origen de datos (df, lista de lector o .xlsx)
# -------------------------------------------------------------------
# Acepta:
# - data.frame
# - lista con $datos_tablas (ya normalizada)
# - lista con $data (como entrega lector_limpieza): lista de data.frames
# - lista “plana” cuyos elementos son data.frames
# - ruta a xlsx (+ hoja_principal)
.resolver_datos_multitabla <- function(datos, hoja_principal = NULL) {

  # 1) data.frame simple
  if (is.data.frame(datos)) {
    return(list(principal = datos, tablas = list(principal = datos)))
  }

  # 2) lista con $datos_tablas (compatibilidad)
  if (is.list(datos) && !is.null(datos$datos_tablas) && is.list(datos$datos_tablas)) {
    tablas_list <- Filter(is.data.frame, datos$datos_tablas)
    if (!length(tablas_list)) stop("La lista 'datos$datos_tablas' no tiene data.frames.")
    # elegir principal por hoja_principal si existe, si no, el primero
    principal <- if (!is.null(hoja_principal) && hoja_principal %in% names(tablas_list)) {
      tablas_list[[hoja_principal]]
    } else {
      tablas_list[[1]]
    }
    if (!"principal" %in% names(tablas_list)) {
      tablas_list <- c(list(principal = principal), tablas_list)
    }
    return(list(principal = principal, tablas = tablas_list))
  }

  # 3) lista con $data (salida de lector_limpieza())
  if (is.list(datos) && !is.null(datos$data) && is.list(datos$data)) {
    tablas_list <- Filter(is.data.frame, datos$data)
    if (!length(tablas_list)) stop("La lista 'datos$data' no tiene data.frames.")
    # elegir principal por hoja_principal si existe (nombre exacto), si no, el primero
    principal <- if (!is.null(hoja_principal) && hoja_principal %in% names(tablas_list)) {
      tablas_list[[hoja_principal]]
    } else {
      tablas_list[[1]]
    }
    if (!"principal" %in% names(tablas_list)) {
      tablas_list <- c(list(principal = principal), tablas_list)
    }
    return(list(principal = principal, tablas = tablas_list))
  }

  # 4) lista “plana” de data.frames
  if (is.list(datos)) {
    tablas_list <- Filter(is.data.frame, datos)
    if (length(tablas_list)) {
      principal <- if (!is.null(hoja_principal) && hoja_principal %in% names(tablas_list)) {
        tablas_list[[hoja_principal]]
      } else {
        tablas_list[[1]]
      }
      if (!"principal" %in% names(tablas_list)) {
        tablas_list <- c(list(principal = principal), tablas_list)
      }
      return(list(principal = principal, tablas = tablas_list))
    }
  }

  # 5) ruta a .xlsx (una sola hoja principal)
  if (is.character(datos) && length(datos) == 1 && grepl("\\.xlsx?$", datos, ignore.case = TRUE)) {
    sheets <- readxl::excel_sheets(datos)
    hoja <- hoja_principal %||% (sheets[1] %||% "Sheet1")
    df <- suppressMessages(readxl::read_excel(datos, sheet = hoja))
    return(list(principal = df, tablas = list(principal = df)))
  }

  stop("No pude convertir 'datos' a data.frame/lista/ruta. ",
       "Pásame un data.frame, una lista (con $data o $datos_tablas) o la ruta a un .xlsx.")
}

.construir_indice_tablas <- function(tablas_list){
  claves <- names(tablas_list)
  norm   <- .norm_tab(claves)
  # asegurar 'principal'
  if (!"principal" %in% names(tablas_list)) {
    tablas_list <- c(list(principal = tablas_list[[1]]), tablas_list)
    claves <- names(tablas_list)
    norm   <- .norm_tab(claves)
  }
  map_norm_to_key <- stats::setNames(claves, norm)  # e.g. "s1" -> "S1"
  list(map_norm_to_key = map_norm_to_key, tablas = tablas_list)
}

.encontrar_tabla_para <- function(tabla_plan, idx){
  tnorm <- .norm_tab(tabla_plan)
  if (!nzchar(tnorm) || tnorm %in% c("principal","(principal)")) return("principal")
  if (tnorm %in% names(idx$map_norm_to_key)) return(idx$map_norm_to_key[[tnorm]])
  warning("No encontré la tabla '", tabla_plan, "'. Se evaluará en 'principal'.")
  "principal"
}

# -------------------------------------------------------------------
# 4) Evaluación de reglas (TRUE en flag = inconsistencia)
# -------------------------------------------------------------------
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

#' @title Evaluar consistencia (multi-tabla)
#' @param datos data.frame, lista de lector_limpieza(), o ruta a .xlsx
#' @param plan tibble de plan (ver \code{cargar_plan_excel()})
#' @param hoja_principal si \code{datos} es ruta a .xlsx, hoja a leer
#' @param contar_na_como_inconsistencia lógico (default FALSE)
#' @export
evaluar_consistencia <- function(datos,
                                 plan,
                                 hoja_principal = NULL,
                                 contar_na_como_inconsistencia = FALSE) {
  stopifnot(is.data.frame(plan))

  # Resolver datos y construir índice
  base <- .resolver_datos_multitabla(datos, hoja_principal = hoja_principal)
  idx  <- .construir_indice_tablas(base$tablas)
  tablas <- idx$tablas

  # mapear plan a nombres internos
  plan2 <- tibble::tibble(
    id_regla            = plan[["ID"]],
    tabla_dest          = plan[["Tabla"]] %||% "(principal)",
    seccion             = plan[["Sección"]] %||% NA_character_,
    categoria           = plan[["Categoría"]] %||% NA_character_,
    tipo_observacion    = plan[["Tipo"]] %||% NA_character_,
    nombre_regla        = plan[["Nombre de regla"]],
    objetivo            = plan[["Objetivo"]],
    variable_1          = plan[["Variable 1"]],
    variable_1_etiqueta = plan[["Variable 1 - Etiqueta"]],
    variable_2          = plan[["Variable 2"]],
    variable_2_etiqueta = plan[["Variable 2 - Etiqueta"]],
    variable_3          = plan[["Variable 3"]],
    variable_3_etiqueta = plan[["Variable 3 - Etiqueta"]],
    procesamiento       = plan[["Procesamiento"]]
  )

  base_ok <- if (contar_na_como_inconsistencia) FALSE else TRUE
  resumen_rows <- vector("list", nrow(plan2))

  # Evaluamos regla por regla en su tabla
  for (i in seq_len(nrow(plan2))) {
    r <- plan2[i,]

    pr <- .parsear_regla(.sanear_regex_en_procesamiento(.normalizar_procesamiento(r$procesamiento)),
                         r$nombre_regla)
    if (is.null(pr)) next

    key <- .encontrar_tabla_para(r$tabla_dest, idx)
    df  <- tablas[[key]]
    if (!is.data.frame(df)) {
      warning("La tabla '", key, "' no es un data.frame. Se omite la regla ", r$id_regla)
      next
    }

    # Entorno para eval
    eval_env <- rlang::env(
      rlang::base_env(),
      is_yes = is_yes,
      is_no  = is_no,
      grepl  = .grepl_perl,
      eq_num_na = eq_num_na,
      eq_chr_na = eq_chr_na
    )
    eval_env <- .bind_vars_with_aliases(eval_env, df)

    rhs2 <- pr$rhs
    ok <- TRUE
    # --- justo antes del parseo ---
    rhs2 <- gsub("\\bregex\\s*\\(", "grepl(", rhs2)   # alias por si algo se coló
    rhs2 <- gsub("\\bcount-selected\\s*\\(", "count_selected(", rhs2)
    # selected(var,"k") -> grepl(...)
    rhs2 <- gsub(
      "\\bselected\\s*\\(\\s*([A-Za-z0-9_\\.]+)\\s*,\\s*'([^']+)'\\s*\\)",
      "grepl('(^|\\\\s)\\2(\\\\s|$)', \\1, perl = TRUE)", rhs2, perl = TRUE
    )
    rhs2 <- gsub(
      "\\bselected\\s*\\(\\s*([A-Za-z0-9_\\.]+)\\s*,\\s*\"([^\"]+)\"\\s*\\)",
      "grepl(\"(^|\\\\s)\\2(\\\\s|$)\", \\1, perl = TRUE)", rhs2, perl = TRUE
    )
    expr_parsed <- tryCatch(rlang::parse_expr(rhs2), error = function(e){ ok <<- FALSE; e })
    if (!ok) {
      warning(sprintf("No se pudo parsear la regla %s (%s): %s",
                      r$id_regla, r$nombre_regla, rhs2))
      next
    }
    vals <- tryCatch(rlang::eval_bare(expr_parsed, env = eval_env),
                     error = function(e) {
                       warning(sprintf("Error evaluando %s: %s", r$nombre_regla, e$message))
                       rep(NA, nrow(df))
                     })
    if (!is.logical(vals)) {
      warning(sprintf("La evaluación de %s no devolvió lógico. Marcando NA.", r$nombre_regla))
      vals <- rep(NA, nrow(df))
    }

    # guardar flag y permitir dependencias
    df[[pr$flag]] <- vals
    rlang::env_bind(eval_env, !!pr$flag := vals)
    tablas[[key]] <- df

    n_inc <- sum(!dplyr::coalesce(vals, base_ok), na.rm = TRUE)

    resumen_rows[[i]] <- tibble::tibble(
      id_regla          = r$id_regla,
      nombre_regla      = r$nombre_regla,
      tabla             = key,
      seccion           = r$seccion,
      categoria         = r$categoria,
      tipo_observacion  = r$tipo_observacion,
      flag              = pr$flag,
      n_inconsistencias = n_inc
    )
  }

  resumen <- dplyr::bind_rows(resumen_rows)
  # base para porcentaje por tabla
  tam <- tibble::tibble(tabla = names(tablas), n = vapply(tablas, nrow, integer(1)))
  resumen <- resumen %>%
    dplyr::left_join(tam, by = "tabla") %>%
    dplyr::mutate(porcentaje = ifelse(n > 0, n_inconsistencias / n, NA_real_)) %>%
    dplyr::arrange(dplyr::desc(n_inconsistencias))

  reglas_meta <- dplyr::select(
    plan2,
    id_regla, nombre_regla,
    tabla_dest, seccion, categoria, tipo_observacion,
    variable_1, variable_1_etiqueta,
    variable_2, variable_2_etiqueta,
    variable_3, variable_3_etiqueta,
    procesamiento
  ) %>% dplyr::rename(tabla = tabla_dest)

  list(
    datos          = base$principal,   # compat
    datos_tablas   = tablas,           # principal y repeats
    resumen        = resumen,
    reglas_meta    = reglas_meta
  )
}

# -------------------------------------------------------------------
# 5) Atajo: evaluar cargando el plan directo desde Excel
# -------------------------------------------------------------------
#' @title Evaluar desde Excel
#' @param datos data.frame, lista lector_limpieza() o ruta a .xlsx
#' @param path_xlsx ruta al plan .xlsx
#' @param hoja_principal si \code{datos} es ruta a .xlsx, hoja a leer
#' @export
evaluar_desde_excel <- function(datos,
                                path_xlsx,
                                hoja_principal = NULL,
                                contar_na_como_inconsistencia = FALSE){
  plan <- cargar_plan_excel(path_xlsx, sheet = "Plan")
  evaluar_consistencia(
    datos = datos,
    plan  = plan,
    hoja_principal = hoja_principal,
    contar_na_como_inconsistencia = contar_na_como_inconsistencia
  )
}

# -------------------------------------------------------------------
# 6) Casos por regla (multi-tabla; respeta la columna Tabla)
# -------------------------------------------------------------------
#' @title Extraer observaciones (casos) para una regla
#' @description
#' Devuelve filas **de la tabla correcta** donde la regla marca inconsistencia.
#' @param evaluacion list devuelto por \code{evaluar_consistencia()}
#' @param regla identificador a buscar (por id_regla, nombre_regla o flag)
#' @param por campo de búsqueda: "id_regla", "nombre_regla" o "flag"
#' @param incluir_flag si TRUE añade la columna del flag en la salida
#' @param contar_na_como_inconsistencia NA cuenta como inconsistencia si TRUE
#' @export
observaciones_regla <- function(evaluacion,
                                regla,
                                por = c("id_regla","nombre_regla","flag"),
                                incluir_flag = FALSE,
                                contar_na_como_inconsistencia = FALSE) {
  por <- match.arg(por)
  stopifnot(is.list(evaluacion),
            all(c("datos_tablas","resumen","reglas_meta") %in% names(evaluacion)))

  tablas <- evaluacion$datos_tablas
  res    <- evaluacion$resumen
  meta   <- evaluacion$reglas_meta
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
  tabla_key <- res$tabla[i] %||% "principal"

  mrow <- meta[match(id_regla, meta$id_regla), , drop = FALSE]
  var1 <- mrow$variable_1 %||% NA_character_
  var2 <- mrow$variable_2 %||% NA_character_
  var3 <- mrow$variable_3 %||% NA_character_

  df <- tablas[[tabla_key]] %||% evaluacion$datos
  if (!is.data.frame(df) || !nzchar(flag_name) || !flag_name %in% names(df)) return(tibble())

  base_ok <- if (isTRUE(contar_na_como_inconsistencia)) FALSE else TRUE
  vals    <- df[[flag_name]]
  inc     <- !dplyr::coalesce(vals, base_ok)
  if (!any(inc, na.rm = TRUE)) return(tibble())

  # Aliases de código pulso
  codigo_alias <- c("Codigo pulso","Código pulso","Codigo_pulso","codigo_pulso","codigo.pulso")
  codigo_col <- codigo_alias[codigo_alias %in% names(df)][1] %||% NULL

  cols <- c("_uuid", "_id", codigo_col, var1, var2, var3)
  cols <- unique(na.omit(cols))
  cols <- cols[cols %in% names(df)]

  out <- tibble::as_tibble(df[inc, cols, drop = FALSE])
  if (isTRUE(incluir_flag)) out[[flag_name]] <- df[[flag_name]][inc]
  out
}

# -------------------------------------------------------------------
# 7) Totales de inconsistencias (igual API, multi-tabla under the hood)
# -------------------------------------------------------------------
#' @title Totales de inconsistencia
#' @description Resumen global de inconsistencias por regla (con porcentaje por tabla).
#' @export
total_inconsistencias <- function(evaluacion,
                                  as_kable = FALSE,
                                  titulo = "Totales de inconsistencia por regla") {
  stopifnot(all(c("resumen") %in% names(evaluacion)))
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
    id_regla, nombre_regla, tabla, n_inconsistencias,
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
# 8) Bloques para reporte (llama observaciones_regla multi-tabla)
# -------------------------------------------------------------------
#' @title Construir bloques de reporte por regla
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

  bloques <- dplyr::left_join(res, meta, by = c("id_regla","nombre_regla")) # meta ya guarda 'tabla'
  stopifnot("procesamiento" %in% names(bloques))

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

# Helpers HTML
.san_label <- function(s) {
  s <- as.character(s %||% "")
  s <- gsub("\\*\\*", "", s, perl = TRUE)
  s <- gsub("\\$\\{([A-Za-z0-9_]+)\\}", "«\\1»", s, perl=TRUE)
  s <- gsub("[\\r\\n]+", " ", s, perl = TRUE)
  trimws(gsub("\\s+", " ", s, perl = TRUE))
}
.safe_html <- function(x) htmltools::htmlEscape(as.character(x %||% ""))

# -------------------------------------------------------------------
# 9) Render de bloques kable/HTML (igual API)
# -------------------------------------------------------------------
#' @title Renderizar bloques con kable/HTML
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

  .san_label2 <- function(s){
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

    cab <- paste0("[", b$id_regla %||% "", "]")

    objetivo <- .san_label2(b$objetivo)
    tipo     <- .san_label2(b$tipo_observacion)

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

    lab1 <- .san_label2(b$variable_1_etiqueta)
    lab2 <- .san_label2(b$variable_2_etiqueta)
    lab3 <- .san_label2(b$variable_3_etiqueta)

    resumen_raw <- tibble::tibble(
      `Id regla`   = b$id_regla %||% NA_character_,
      `Tabla`      = b$tabla %||% NA_character_,
      `Variable 1` = b$variable_1 %||% NA_character_,
      `Etiqueta 1` = if (nz(lab1)) lab1 else NA_character_,
      `Variable 2` = b$variable_2 %||% NA_character_,
      `Etiqueta 2` = if (nz(lab2)) lab2 else NA_character_,
      `Variable 3` = b$variable_3 %||% NA_character_,
      `Etiqueta 3` = if (nz(lab3)) lab3 else NA_character_,
      `# inconsist.` = b$n_inconsistencias %||% 0L,
      `%`           = if (!is.na(b$porcentaje)) sprintf("%.1f%%", 100*b$porcentaje) else NA_character_
    )

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

    # Casos
    casos <- b$casos[[1]]
    vars_regla <- unique(na.omit(c(b$variable_1, b$variable_2, b$variable_3)))
    vars_regla <- vars_regla[nzchar(vars_regla)]
    vars_regla <- maybe_map(vars_regla)

    var_labels <- paste0(vars_regla, "_label")
    desired_cols <- unique(c(c("_uuid","_id"), as.vector(rbind(vars_regla, var_labels)), cols_interes))
    desired_cols <- desired_cols[!is.na(desired_cols) & nzchar(desired_cols)]

    usar_fallback <- (is.null(casos) || !nrow(casos)) && mostrar_aleatorios_si_cero > 0 &&
      !is.null(fallback_df) && nrow(fallback_df) > 0

    if (usar_fallback) {
      n_take <- min(mostrar_aleatorios_si_cero, nrow(fallback_df))
      idx <- sample.int(nrow(fallback_df), n_take)
      keep <- .exact_cols(fallback_df, desired_cols)
      if (!length(keep)) keep <- .exact_cols(fallback_df, c("_uuid","_id"))
      if (!length(keep)) keep <- head(names(fallback_df), 6)
      casos <- tibble::as_tibble(fallback_df[idx, keep, drop = FALSE])
    } else if (!is.null(casos) && nrow(casos)) {
      keep <- .exact_cols(casos, desired_cols)
      if (!length(keep)) keep <- .exact_cols(casos, c("_uuid","_id", vars_regla))
      if (!length(keep)) keep <- .exact_cols(casos, c("_uuid","_id"))
      if (!length(keep)) keep <- head(names(casos), 6)
      casos <- casos[, keep, drop = FALSE]
    }

    if (isTRUE(quitar_vacias_en_casos)) {
      if (!is.null(casos) && ncol(casos)) {
        keep <- vapply(casos, function(col) any(!is.na(col) & trimws(as.character(col)) != ""), logical(1))
        casos <- casos[, keep, drop = FALSE]
      }
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

    proc <- unname(as.character(b$procesamiento %||% ""))
    proc_node <- if (.nz(proc)) {
      htmltools::tags$div(
        style="margin-top:8px;",
        htmltools::tags$strong("Procesamiento utilizado en R:"),
        htmltools::tags$pre(
          style="white-space:pre-wrap; border:1px solid #ddd; padding:8px; border-radius:6px; background:#fafafa; font-family: ui-monospace, Menlo, Consolas, 'Liberation Mono', monospace;",
          htmltools::HTML(proc)
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
# 10) Reporte HTML completo
# -------------------------------------------------------------------
#' @title Generar reporte HTML de limpieza
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

  # Nodos detallados
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

  # HTML
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

