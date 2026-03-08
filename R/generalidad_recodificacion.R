#' Recodificar items categóricos a escala estandarizada 0-100
#'
#' Toma una base procesada con [reporte_data()] y recodifica variables
#' `select_one` usando la estructura del instrumento de [reporte_instrumento()].
#' La recodificación se hace por orden de categorías del `list_name` y convierte
#' respuestas sustantivas a una escala 0-100.
#'
#' Reglas para categorías especiales:
#' \itemize{
#'   \item `codigos_missing` se tratan como `missing`.
#'   \item `codigos_no_aplica` se tratan como `no_aplica`.
#'   \item Tanto `missing` como `no_aplica` se recodifican a `NA_real_`.
#' }
#'
#' @param data Objeto devuelto por [reporte_data()] (clase
#'   `"prosecnur_reporte_tbl"`) o un `data.frame` con estructura equivalente.
#' @param instrumento Objeto devuelto por [reporte_instrumento()]. Si es `NULL`,
#'   se toma desde `attr(data, "instrumento_reporte")`.
#' @param vars Variables a recodificar. Si es `NULL`, usa todas las `select_one`
#'   del instrumento presentes en `data`.
#' @param excluir_vars Variables a excluir de la recodificación.
#' @param orden_por_lista Configuración de orden ascendente (de menor a mayor)
#'   por `list_name`, usando códigos (`choices$name`). Debe ser una lista
#'   nombrada por `list_name`, donde cada elemento es un vector de códigos en
#'   el orden deseado.
#' @param codigos_missing Códigos a tratar como `missing`, usando solo códigos
#'   (`choices$name`). Acepta vector global de códigos o lista nombrada por
#'   `list_name` (opcionalmente con `.default`).
#' @param codigos_no_aplica Códigos a tratar como `no_aplica`, usando solo
#'   códigos (`choices$name`). Acepta vector global de códigos o lista nombrada
#'   por `list_name` (opcionalmente con `.default`).
#' @param prefijo Prefijo para nuevas columnas recodificadas cuando
#'   `reemplazar = FALSE`.
#' @param reemplazar Si `TRUE`, reemplaza la variable original. Si `FALSE`,
#'   crea columnas nuevas con prefijo.
#'
#' @return El mismo `data.frame` con variables recodificadas a 0-100. Además
#'   agrega atributo `recodificacion_items_meta` con trazabilidad por variable
#'   (list_name, mapeo aplicado y conteos de calidad).
#' @export
reporte_recodificar_items <- function(
    data,
    instrumento = NULL,
    vars = NULL,
    excluir_vars = NULL,
    orden_por_lista = NULL,
    codigos_missing = character(0),
    codigos_no_aplica = character(0),
    prefijo = "r100_",
    reemplazar = FALSE
) {
  if (!is.data.frame(data)) {
    stop("`data` debe ser un data.frame o tibble.", call. = FALSE)
  }

  if (is.null(instrumento)) {
    instrumento <- attr(data, "instrumento_reporte", exact = TRUE)
  }
  if (is.null(instrumento) || !is.list(instrumento)) {
    stop(
      "No se pudo resolver `instrumento`. Pásalo explícitamente o usa salida de `reporte_data()`.",
      call. = FALSE
    )
  }

  survey <- instrumento$survey
  choices <- instrumento$choices
  orders_list <- instrumento$orders_list

  if (!is.data.frame(survey) || !all(c("name", "type") %in% names(survey))) {
    stop("`instrumento$survey` no tiene la estructura esperada.", call. = FALSE)
  }
  if (!is.data.frame(choices) || !all(c("list_name", "name") %in% names(choices))) {
    stop("`instrumento$choices` no tiene la estructura esperada.", call. = FALSE)
  }

  .rr_norm_token <- function(x) {
    y <- as.character(x)
    y <- iconv(y, from = "", to = "ASCII//TRANSLIT")
    y <- tolower(trimws(y))
    y <- gsub("\\s+", " ", y)
    y
  }

  .rr_cfg_by_list <- function(cfg, list_name) {
    if (is.null(cfg)) return(character(0))
    if (is.list(cfg)) {
      nms <- names(cfg)
      if (!is.null(nms) && length(nms)) {
        if (!is.na(list_name) && nzchar(list_name) && list_name %in% nms) {
          return(as.character(cfg[[list_name]]))
        }
        if (".default" %in% nms) {
          return(as.character(cfg[[".default"]]))
        }
      }
      return(character(0))
    }
    as.character(cfg)
  }

  if (is.null(vars)) {
    is_so <- grepl("^select_one(\\s|$)", as.character(survey$type))
    vars <- unique(as.character(survey$name[is_so]))
    vars <- vars[vars %in% names(data)]
  } else {
    vars <- as.character(vars)
    vars <- vars[!is.na(vars) & nzchar(trimws(vars))]
    vars <- unique(vars)
    vars <- vars[vars %in% names(data)]
  }

  if (!is.null(excluir_vars)) {
    excluir_vars <- as.character(excluir_vars)
    excluir_vars <- excluir_vars[!is.na(excluir_vars) & nzchar(trimws(excluir_vars))]
    vars <- setdiff(vars, unique(excluir_vars))
  }

  if (!length(vars)) {
    stop("No hay variables candidatas para recodificar.", call. = FALSE)
  }

  out <- data
  meta <- list()

  for (v in vars) {
    idx <- which(as.character(survey$name) == v)
    ln <- NA_character_
    if (length(idx) && "list_name" %in% names(survey)) {
      ln_vals <- as.character(survey$list_name[idx])
      ln_vals <- ln_vals[!is.na(ln_vals) & nzchar(ln_vals)]
      if (length(ln_vals)) ln <- ln_vals[1]
    }

    ord_codes <- character(0)
    ord_labels <- character(0)

    if (!is.null(orders_list) && v %in% names(orders_list)) {
      ent <- orders_list[[v]]
      ord_codes <- if (!is.null(ent$names)) as.character(ent$names) else character(0)
      ord_labels <- if (!is.null(ent$labels)) as.character(ent$labels) else character(0)
    }

    if (!length(ord_codes) && !is.na(ln)) {
      ch <- choices[as.character(choices$list_name) == ln, , drop = FALSE]
      if (nrow(ch)) {
        ord_codes <- as.character(ch$name)
        ord_labels <- if ("label" %in% names(ch)) as.character(ch$label) else ord_codes
      }
    }

    if (!length(ord_codes)) {
      warning("Se omite `", v, "`: no se pudo resolver categorías desde instrumento.", call. = FALSE)
      next
    }

    if (length(ord_labels) != length(ord_codes)) {
      ord_labels <- ord_codes
    }

    ord_user <- .rr_cfg_by_list(orden_por_lista, ln)
    ord_user <- ord_user[!is.na(ord_user) & nzchar(trimws(ord_user))]
    if (length(ord_user)) {
      ord_user_tok <- .rr_norm_token(ord_user)
      ord_codes_tok <- .rr_norm_token(ord_codes)
      use_tok <- intersect(ord_user_tok, ord_codes_tok)
      if (length(use_tok)) {
        rest_tok <- setdiff(ord_codes_tok, use_tok)
        new_tok <- c(use_tok, rest_tok)
        idx_new <- match(new_tok, ord_codes_tok)
        ord_codes <- ord_codes[idx_new]
        ord_labels <- ord_labels[idx_new]
      }
    }

    keep <- !is.na(ord_codes) & nzchar(ord_codes)
    ord_codes <- ord_codes[keep]
    ord_labels <- ord_labels[keep]
    if (!length(ord_codes)) {
      warning("Se omite `", v, "`: categorías vacías tras limpieza.", call. = FALSE)
      next
    }

    tok_code <- .rr_norm_token(ord_codes)

    tok_missing <- unique(.rr_norm_token(.rr_cfg_by_list(codigos_missing, ln)))
    tok_no_aplica <- unique(.rr_norm_token(.rr_cfg_by_list(codigos_no_aplica, ln)))
    tok_missing <- tok_missing[nzchar(tok_missing)]
    tok_no_aplica <- tok_no_aplica[nzchar(tok_no_aplica)]

    is_no_aplica_cat <- tok_code %in% tok_no_aplica
    is_missing_cat <- tok_code %in% tok_missing
    is_substantive <- !(is_no_aplica_cat | is_missing_cat)

    tok_missing_v <- unique(c(tok_missing, tok_code[is_missing_cat]))
    tok_no_aplica_v <- unique(c(tok_no_aplica, tok_code[is_no_aplica_cat]))

    sub_codes <- ord_codes[is_substantive]
    sub_labels <- ord_labels[is_substantive]
    sub_tok_code <- tok_code[is_substantive]

    n_sub <- length(sub_codes)
    if (n_sub == 0L) {
      warning("Se omite `", v, "`: sin categorías sustantivas para recodificar.", call. = FALSE)
      next
    }

    scores <- if (n_sub == 1L) 100 else seq(0, 100, length.out = n_sub)
    map_tok <- stats::setNames(as.numeric(scores), sub_tok_code)

    x_raw <- out[[v]]
    x_chr <- as.character(x_raw)
    x_tok <- .rr_norm_token(x_chr)

    res <- rep(NA_real_, length(x_tok))
    motivo <- rep("no_mapeado", length(x_tok))

    idx_na_entrada <- is.na(x_raw) | !nzchar(x_tok) | x_tok == "na"
    idx_no_aplica <- !idx_na_entrada & (x_tok %in% tok_no_aplica_v)
    idx_missing <- !idx_na_entrada & !idx_no_aplica & (x_tok %in% tok_missing_v)
    idx_valid <- !idx_na_entrada & !idx_no_aplica & !idx_missing & (x_tok %in% names(map_tok))

    if (any(idx_valid)) {
      res[idx_valid] <- unname(map_tok[x_tok[idx_valid]])
    }

    motivo[idx_na_entrada] <- "na_entrada"
    motivo[idx_no_aplica] <- "no_aplica"
    motivo[idx_missing] <- "missing"
    motivo[idx_valid] <- "valido"

    out_name <- if (isTRUE(reemplazar)) v else paste0(prefijo, v)
    out[[out_name]] <- res

    var_label <- attr(out[[v]], "label", exact = TRUE)
    var_label <- if (!is.null(var_label) && nzchar(trimws(as.character(var_label)))) {
      as.character(var_label)
    } else {
      v
    }
    attr(out[[out_name]], "label") <- paste0(var_label, " [0-100]")
    attr(out[[out_name]], "measure") <- "scale"

    meta[[v]] <- list(
      variable = v,
      variable_salida = out_name,
      list_name = ln,
      mapeo = tibble::tibble(
        codigo = sub_codes,
        etiqueta = sub_labels,
        score_0_100 = as.numeric(scores)
      ),
      conteos = c(
        n_total = length(x_tok),
        n_valido = sum(motivo == "valido"),
        n_no_aplica = sum(motivo == "no_aplica"),
        n_missing = sum(motivo == "missing"),
        n_no_mapeado = sum(motivo == "no_mapeado"),
        n_na_entrada = sum(motivo == "na_entrada")
      )
    )
  }

  attr(out, "recodificacion_items_meta") <- meta
  out
}

#' Construir índices jerárquicos desde variables recodificadas (0-100)
#'
#' Calcula promedios simples por bloques (por ejemplo, conductores) y luego
#' índices de segundo nivel a partir de esos bloques u otras variables numéricas.
#' La función es agnóstica al estudio: solo usa las declaraciones entregadas.
#'
#' @param data `data.frame` con ítems recodificados (0-100).
#' @param bloques Lista nombrada. Cada elemento es un vector de variables que
#'   componen un bloque de primer nivel.
#' @param indices Lista nombrada opcional. Cada elemento es un vector de
#'   referencias para construir índices de segundo nivel. Puede referir a:
#'   \itemize{
#'     \item nombres de bloque (clave en `bloques`), o
#'     \item variables/columnas ya existentes en `data`.
#'   }
#' @param min_prop_valid Proporción mínima de ítems válidos requerida por fila.
#'   Puede ser un número único (ej. `0.5`) o un vector nombrado por bloque/índice
#'   (opcionalmente con `.default`).
#' @param prefijo_bloque Prefijo de columnas creadas para bloques.
#' @param prefijo_indice Prefijo de columnas creadas para índices.
#'
#' @return El mismo `data.frame` con columnas agregadas para bloques e índices.
#'   Agrega además `attr(x, "indices_meta")` con trazabilidad de cálculos.
#' @export
reporte_construir_indices <- function(
    data,
    bloques,
    indices = NULL,
    min_prop_valid = 0.5,
    prefijo_bloque = "bloq_",
    prefijo_indice = "idx_"
) {
  if (!is.data.frame(data)) {
    stop("`data` debe ser un data.frame o tibble.", call. = FALSE)
  }
  if (!is.list(bloques) || is.null(names(bloques)) || any(!nzchar(names(bloques)))) {
    stop("`bloques` debe ser una lista nombrada.", call. = FALSE)
  }

  .ri_min_prop <- function(id, cfg) {
    if (length(cfg) == 1L && is.numeric(cfg) && is.finite(cfg)) {
      return(as.numeric(cfg)[1])
    }
    if (!is.null(names(cfg)) && length(names(cfg))) {
      if (id %in% names(cfg)) return(as.numeric(cfg[[id]])[1])
      if (".default" %in% names(cfg)) return(as.numeric(cfg[[".default"]])[1])
    }
    0.5
  }

  .ri_row_mean_min <- function(df_num, min_prop = 0.5) {
    n_total <- ncol(df_num)
    if (n_total == 0L) return(rep(NA_real_, nrow(df_num)))
    n_valid <- rowSums(!is.na(df_num))
    out <- rowMeans(df_num, na.rm = TRUE)
    out[n_valid == 0L] <- NA_real_
    out[(n_valid / n_total) < min_prop] <- NA_real_
    out
  }

  out <- data
  meta_bloques <- list()
  meta_indices <- list()

  for (id in names(bloques)) {
    vars <- as.character(bloques[[id]])
    vars <- vars[!is.na(vars) & nzchar(trimws(vars))]
    vars <- unique(vars)
    vars_ok <- vars[vars %in% names(out)]
    if (!length(vars_ok)) {
      warning("Bloque `", id, "` sin variables disponibles en `data`.", call. = FALSE)
      next
    }

    X <- as.data.frame(out[, vars_ok, drop = FALSE])
    X[] <- lapply(X, function(v) suppressWarnings(as.numeric(v)))
    mp <- .ri_min_prop(id, min_prop_valid)
    score <- .ri_row_mean_min(X, min_prop = mp)

    out_name <- paste0(prefijo_bloque, id)
    out[[out_name]] <- score
    attr(out[[out_name]], "label") <- id
    attr(out[[out_name]], "measure") <- "scale"

    meta_bloques[[id]] <- list(
      salida = out_name,
      vars = vars_ok,
      min_prop_valid = mp,
      n_vars = length(vars_ok)
    )
  }

  if (!is.null(indices)) {
    if (!is.list(indices) || is.null(names(indices)) || any(!nzchar(names(indices)))) {
      stop("`indices` debe ser una lista nombrada.", call. = FALSE)
    }

    for (id in names(indices)) {
      refs <- as.character(indices[[id]])
      refs <- refs[!is.na(refs) & nzchar(trimws(refs))]
      refs <- unique(refs)

      cols <- character(0)
      for (r in refs) {
        c1 <- paste0(prefijo_bloque, r)
        if (c1 %in% names(out)) {
          cols <- c(cols, c1)
        } else if (r %in% names(out)) {
          cols <- c(cols, r)
        }
      }
      cols <- unique(cols)
      if (!length(cols)) {
        warning("Índice `", id, "` sin referencias disponibles en `data`.", call. = FALSE)
        next
      }

      X <- as.data.frame(out[, cols, drop = FALSE])
      X[] <- lapply(X, function(v) suppressWarnings(as.numeric(v)))
      mp <- .ri_min_prop(id, min_prop_valid)
      score <- .ri_row_mean_min(X, min_prop = mp)

      out_name <- paste0(prefijo_indice, id)
      out[[out_name]] <- score
      attr(out[[out_name]], "label") <- id
      attr(out[[out_name]], "measure") <- "scale"

      meta_indices[[id]] <- list(
        salida = out_name,
        refs = refs,
        refs_resueltas = cols,
        min_prop_valid = mp,
        n_refs = length(cols)
      )
    }
  }

  attr(out, "indices_meta") <- list(
    bloques = meta_bloques,
    indices = meta_indices
  )
  out
}

#' Construir configuración de dimensiones para `reporte_interactivo()`
#'
#' Genera una configuración lista para el Tab 4 (Dimensiones) a partir de los
#' metadatos producidos por [reporte_recodificar_items()] y
#' [reporte_construir_indices()]. El objetivo es separar los nombres técnicos
#' (`idx_*`, `bloq_*`, `r100_*`) de las etiquetas orientadas al usuario.
#'
#' @param data `data.frame` que contiene resultados recodificados/índices.
#' @param labels_indices Vector/lista nombrada opcional para rotular índices.
#'   Acepta claves por nombre lógico (`indice_general`) o por columna de salida
#'   (`idx_indice_general`).
#' @param labels_bloques Vector/lista nombrada opcional para rotular bloques.
#'   Acepta claves por nombre lógico (`trato`) o por columna (`bloq_trato`).
#' @param labels_indicadores Vector/lista nombrada opcional para rotular
#'   indicadores (`r100_*`).
#' @param semaforo_cortes Numeric(2) con cortes del semáforo en escala 0-100.
#'   Por defecto `c(50, 75)`.
#' @param semaforo_colores Vector nombrado de 3 colores (`rojo`, `ambar`,
#'   `verde`) para el heatmap semafórico.
#' @param radar_min_ejes Número mínimo de ejes para usar radar (si no se
#'   cumple, Tab 4 puede usar barras comparativas).
#' @param incluir_total_default Si `TRUE`, Tab 4 inicia mostrando `Total`.
#' @param iteracion_habilitada_default Si `TRUE`, Tab 4 puede iniciar con
#'   iteración habilitada (si hay variable disponible).
#' @param max_categorias_principal Máximo de categorías visibles para variable
#'   principal.
#' @param max_niveles_iteracion Máximo de niveles visibles de iteración.
#' @param paleta_radar Paleta cualitativa por defecto del radar (`"okabe_ito"`
#'   o `"ipe"`).
#'
#' @return Una lista con:
#' \itemize{
#'   \item `catalog_general`: catálogo de objetivos de vista General.
#'   \item `catalog_indicadores`: catálogo de objetivos de vista Indicadores.
#'   \item `labels_indices`, `labels_bloques`, `labels_indicadores`.
#'   \item `semaforo`: cortes y colores.
#'   \item `visual`: reglas del motor visual.
#' }
#' @export
reporte_dimensiones_config <- function(
    data,
    labels_indices = NULL,
    labels_bloques = NULL,
    labels_indicadores = NULL,
    semaforo_cortes = c(50, 75),
    semaforo_colores = c(rojo = "#D84B55", ambar = "#E0B44C", verde = "#3A9A5B"),
    radar_min_ejes = 3L,
    incluir_total_default = TRUE,
    iteracion_habilitada_default = FALSE,
    max_categorias_principal = 8L,
    max_niveles_iteracion = 12L,
    paleta_radar = c("okabe_ito", "ipe")
) {
  `%||%` <- function(x, y) if (!is.null(x)) x else y

  if (!is.data.frame(data)) {
    stop("`data` debe ser un data.frame o tibble.", call. = FALSE)
  }

  paleta_radar <- match.arg(paleta_radar)

  .as_named_chr <- function(x) {
    if (is.null(x)) return(stats::setNames(character(0), character(0)))
    v <- as.character(unlist(x, use.names = TRUE))
    n <- names(v)
    if (is.null(n)) return(stats::setNames(character(0), character(0)))
    ok <- !is.na(n) & nzchar(trimws(n)) & !is.na(v) & nzchar(trimws(v))
    stats::setNames(v[ok], n[ok])
  }

  .pretty <- function(x) {
    x <- as.character(x %||% "")
    x <- gsub("^idx_", "", x)
    x <- gsub("^bloq_", "", x)
    x <- gsub("^r100_", "", x)
    x <- gsub("[_\\.]+", " ", x)
    x <- trimws(x)
    if (!nzchar(x)) return("Variable")
    paste0(toupper(substring(x, 1, 1)), substring(x, 2))
  }

  .label_data <- function(v) {
    if (!(v %in% names(data))) return(.pretty(v))
    lb <- attr(data[[v]], "label", exact = TRUE)
    lb <- as.character(lb %||% "")
    lb <- gsub("\\s*\\[0-100\\]$", "", lb)
    if (nzchar(trimws(lb))) trimws(lb) else .pretty(v)
  }

  labels_indices <- .as_named_chr(labels_indices)
  labels_bloques <- .as_named_chr(labels_bloques)
  labels_indicadores <- .as_named_chr(labels_indicadores)

  .nm_get <- function(x, key) {
    key <- as.character(key %||% "")[1]
    if (!nzchar(key)) return(NULL)
    nms <- names(x)
    if (is.null(nms)) return(NULL)
    i <- match(key, nms)
    if (is.na(i)) return(NULL)
    as.character(x[i])[1]
  }

  semaforo_cortes <- suppressWarnings(as.numeric(semaforo_cortes))
  semaforo_cortes <- semaforo_cortes[is.finite(semaforo_cortes) & !is.na(semaforo_cortes)]
  if (length(semaforo_cortes) < 2L) semaforo_cortes <- c(50, 75)
  semaforo_cortes <- sort(unique(semaforo_cortes))[1:2]
  semaforo_cortes <- pmax(0, pmin(100, semaforo_cortes))
  if (length(semaforo_cortes) < 2L || semaforo_cortes[1] >= semaforo_cortes[2]) {
    semaforo_cortes <- c(50, 75)
  }

  semaforo_colores <- as.character(semaforo_colores %||% character(0))
  nmsc <- names(semaforo_colores %||% character(0))
  if (is.null(nmsc)) nmsc <- character(0)
  col_rojo <- if ("rojo" %in% nmsc) semaforo_colores[["rojo"]] else "#D84B55"
  col_amb <- if ("ambar" %in% nmsc) semaforo_colores[["ambar"]] else "#E0B44C"
  col_ver <- if ("verde" %in% nmsc) semaforo_colores[["verde"]] else "#3A9A5B"

  radar_min_ejes <- suppressWarnings(as.integer(radar_min_ejes)[1])
  if (!is.finite(radar_min_ejes) || is.na(radar_min_ejes) || radar_min_ejes < 1L) radar_min_ejes <- 3L

  max_categorias_principal <- suppressWarnings(as.integer(max_categorias_principal)[1])
  if (!is.finite(max_categorias_principal) || is.na(max_categorias_principal) || max_categorias_principal < 1L) {
    max_categorias_principal <- 8L
  }

  max_niveles_iteracion <- suppressWarnings(as.integer(max_niveles_iteracion)[1])
  if (!is.finite(max_niveles_iteracion) || is.na(max_niveles_iteracion) || max_niveles_iteracion < 1L) {
    max_niveles_iteracion <- 12L
  }

  idx_meta <- attr(data, "indices_meta", exact = TRUE)
  rec_meta <- attr(data, "recodificacion_items_meta", exact = TRUE)
  meta_bloques <- if (is.list(idx_meta) && is.list(idx_meta$bloques)) idx_meta$bloques else list()
  meta_indices <- if (is.list(idx_meta) && is.list(idx_meta$indices)) idx_meta$indices else list()

  block_key_to_var <- stats::setNames(
    vapply(meta_bloques, function(x) as.character(x$salida %||% NA_character_)[1], character(1)),
    names(meta_bloques)
  )
  block_key_to_var <- block_key_to_var[!is.na(block_key_to_var) & nzchar(block_key_to_var)]
  block_var_to_key <- stats::setNames(names(block_key_to_var), as.character(block_key_to_var))

  rec_var_to_source <- stats::setNames(character(0), character(0))
  if (is.list(rec_meta) && length(rec_meta)) {
    rec_df <- data.frame(
      src = names(rec_meta),
      out = vapply(rec_meta, function(x) as.character(x$variable_salida %||% NA_character_)[1], character(1)),
      stringsAsFactors = FALSE
    )
    rec_df <- rec_df[!is.na(rec_df$out) & nzchar(rec_df$out), , drop = FALSE]
    if (nrow(rec_df)) {
      rec_var_to_source <- stats::setNames(as.character(rec_df$src), as.character(rec_df$out))
    }
  }

  catalog_general <- list()
  for (id in names(meta_indices)) {
    it <- meta_indices[[id]]
    idx_var <- as.character(it$salida %||% NA_character_)[1]
    if (is.na(idx_var) || !nzchar(idx_var) || !(idx_var %in% names(data))) next

    refs <- unique(c(
      as.character(it$refs_resueltas %||% character(0)),
      as.character(it$refs %||% character(0))
    ))

    axis_vars <- character(0)
    axis_labels <- character(0)
    for (r in refs) {
      rv <- if (r %in% names(data)) {
        r
      } else if (r %in% names(block_key_to_var)) {
        as.character(block_key_to_var[[r]])
      } else {
        NA_character_
      }
      if (is.na(rv) || !nzchar(rv) || !(rv %in% names(data)) || rv %in% axis_vars) next

      axis_vars <- c(axis_vars, rv)

      bkey <- if (rv %in% names(block_var_to_key)) as.character(block_var_to_key[[rv]]) else rv
      lb <- .nm_get(labels_bloques, bkey) %||% .nm_get(labels_bloques, rv) %||% .label_data(rv)
      axis_labels <- c(axis_labels, as.character(lb))
    }
    if (!length(axis_vars)) next

    ilab <- .nm_get(labels_indices, id) %||% .nm_get(labels_indices, idx_var) %||% .label_data(idx_var)
    catalog_general[[idx_var]] <- list(
      id = idx_var,
      key = id,
      label = as.character(ilab),
      axis_vars = axis_vars,
      axis_labels = axis_labels
    )
  }

  if (!length(catalog_general)) {
    idx_vars <- grep("^idx_", names(data), value = TRUE)
    bloq_vars <- grep("^bloq_", names(data), value = TRUE)
    if (length(idx_vars) && length(bloq_vars)) {
      axis_labels <- vapply(bloq_vars, function(v) {
        bkey <- if (v %in% names(block_var_to_key)) as.character(block_var_to_key[[v]]) else v
        as.character(.nm_get(labels_bloques, bkey) %||% .nm_get(labels_bloques, v) %||% .label_data(v))
      }, character(1))
      for (v in idx_vars) {
        ilab <- .nm_get(labels_indices, v) %||% .label_data(v)
        catalog_general[[v]] <- list(
          id = v,
          key = v,
          label = as.character(ilab),
          axis_vars = bloq_vars,
          axis_labels = axis_labels
        )
      }
    }
  }

  catalog_indicadores <- list()
  for (bk in names(meta_bloques)) {
    bl <- meta_bloques[[bk]]
    bvar <- as.character(bl$salida %||% NA_character_)[1]
    vars <- unique(as.character(bl$vars %||% character(0)))
    vars <- vars[vars %in% names(data)]
    if (!length(vars)) next

    blab <- .nm_get(labels_bloques, bk) %||% .nm_get(labels_bloques, bvar) %||% .pretty(bk)
    ilabs <- vapply(vars, function(v) {
      src <- .nm_get(rec_var_to_source, v) %||% v
      as.character(.nm_get(labels_indicadores, v) %||% .nm_get(labels_indicadores, src) %||% .label_data(v))
    }, character(1))

    catalog_indicadores[[bk]] <- list(
      id = bk,
      key = bk,
      label = as.character(blab),
      block_var = bvar,
      axis_vars = vars,
      axis_labels = ilabs
    )
  }

  list(
    version = 1L,
    catalog_general = catalog_general,
    catalog_indicadores = catalog_indicadores,
    labels_indices = labels_indices,
    labels_bloques = labels_bloques,
    labels_indicadores = labels_indicadores,
    semaforo = list(
      cortes = as.numeric(semaforo_cortes),
      colores = c(rojo = col_rojo, ambar = col_amb, verde = col_ver)
    ),
    visual = list(
      radar_min_ejes = as.integer(radar_min_ejes),
      incluir_total_default = isTRUE(incluir_total_default),
      iteracion_habilitada_default = isTRUE(iteracion_habilitada_default),
      max_categorias_principal = as.integer(max_categorias_principal),
      max_niveles_iteracion = as.integer(max_niveles_iteracion),
      paleta_radar = as.character(paleta_radar)
    )
  )
}

#' Exportar tablas SPSS de promedios (0-100) con cruces horizontales
#'
#' Genera tablas de promedios ponderados para variables recodificadas e índices
#' (sin desviación estándar ni percentiles), con estilo SPSS y encabezados
#' humanos mergeados. Permite:
#' \itemize{
#'   \item Filas por una variable de agrupación (`fila`), o solo total.
#'   \item Columnas por una o varias variables de cruce (`cruzar_con`).
#'   \item Redondeo configurable con criterio escolar (half-up).
#' }
#'
#' Estructura de columnas:
#' \itemize{
#'   \item Nivel 1: variable de cruce (ej. Sexo, Distrito).
#'   \item Nivel 2: categoría del cruce (ej. Mujer, Hombre).
#'   \item Nivel 3: indicador (variables declaradas en `secciones`).
#' }
#'
#' @param data `data.frame` con indicadores numéricos (0-100).
#' @param secciones Vector o lista nombrada de indicadores. Si es vector, se usa
#'   una sección única definida por `titulo_default`. Se ignora cuando se usa
#'   `tablas`.
#' @param tablas Lista opcional de tablas a exportar en un mismo Excel. Cada
#'   elemento puede declarar: `titulo`, `indicadores`, `fila`, `cruzar_con`,
#'   `nombres_indicadores`, `agregar_brecha`, `etiqueta_brecha` e
#'   `incluir_total`.
#' @param titulo_default Título de sección por defecto cuando `secciones` se pasa
#'   como vector (no lista).
#' @param modo_cruces Cómo tratar múltiples cruces:
#'   `juntos` (un solo bloque horizontal) o `separados` (una tabla por cruce).
#' @param separacion_filas Número de filas en blanco entre tablas dentro de la
#'   misma hoja.
#' @param fila Variable opcional para filas (por ejemplo, `servicio` o
#'   `municipio`). Si es `NULL`, se usa una sola fila `"Total"`.
#' @param cruzar_con Vector opcional de variables para columnas (una o varias).
#' @param path_xlsx Ruta de salida Excel.
#' @param hoja Nombre de hoja de salida.
#' @param dic_vars Diccionario opcional de etiquetas (`name`, `label`).
#' @param labels_override Lista opcional de etiquetas por variable.
#' @param nombres_indicadores Vector/lista nombrada para renombrar indicadores
#'   en el encabezado de columnas (clave = nombre de variable, valor = etiqueta).
#' @param survey `survey` del instrumento (opcional).
#' @param orders_list `orders_list` del instrumento (opcional).
#' @param fuente Texto de fuente al pie de tabla.
#' @param weight_col Variable de peso. Si es `NULL`, intenta
#'   `attr(data, "var_peso")` y luego `peso`.
#' @param digits Número de decimales para el promedio.
#' @param incluir_total Si `TRUE`, agrega bloque de columnas `"Total"` aun cuando
#'   existan cruces.
#' @param mostrar_todo Si `TRUE`, muestra todas las categorías definidas en
#'   instrumento para `fila`/`cruzar_con`, aunque no tengan casos.
#' @param agregar_brecha Si `TRUE`, agrega fila `"Brecha"` calculada como
#'   `max - min` entre categorías de fila para cada columna de indicador.
#' @param etiqueta_brecha Texto de la etiqueta de la fila de brecha.
#'
#' @return Ruta normalizada del archivo Excel generado (invisible).
#' @export
reporte_tablas_recodificadas <- function(
    data,
    secciones,
    tablas = NULL,
    titulo_default = "INDICES",
    modo_cruces = c("juntos", "separados"),
    separacion_filas = 3,
    fila = NULL,
    cruzar_con = NULL,
    path_xlsx = "tablas_indices_spss.xlsx",
    hoja = "Indices",
    dic_vars = NULL,
    labels_override = NULL,
    nombres_indicadores = NULL,
    survey = NULL,
    orders_list = NULL,
    fuente = "Pulso PUCP",
    weight_col = NULL,
    digits = 1,
    incluir_total = TRUE,
    mostrar_todo = FALSE,
    agregar_brecha = FALSE,
    etiqueta_brecha = "Brecha"
) {
  if (!is.data.frame(data)) {
    stop("`data` debe ser un data.frame o tibble.", call. = FALSE)
  }
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("Se requiere el paquete 'openxlsx'.", call. = FALSE)
  }

  digits <- suppressWarnings(as.integer(digits)[1])
  if (!is.finite(digits) || is.na(digits) || digits < 0) {
    stop("`digits` debe ser un entero mayor o igual a 0.", call. = FALSE)
  }

  modo_cruces <- match.arg(modo_cruces)
  separacion_filas <- suppressWarnings(as.integer(separacion_filas)[1])
  if (!is.finite(separacion_filas) || is.na(separacion_filas) || separacion_filas < 1L) {
    stop("`separacion_filas` debe ser un entero mayor o igual a 1.", call. = FALSE)
  }

  usar_tablas_plan <- !is.null(tablas)
  if (!usar_tablas_plan) {
    if (is.character(secciones)) {
      vars <- as.character(secciones)
      vars <- vars[!is.na(vars) & nzchar(trimws(vars))]
      secciones <- list(unique(vars))
      names(secciones) <- as.character(titulo_default)[1]
    }
    if (!is.list(secciones) || !length(secciones)) {
      stop("`secciones` debe ser un vector o una lista cuando `tablas` es NULL.", call. = FALSE)
    }
    if (is.null(names(secciones))) {
      names(secciones) <- paste0("SECCION_", seq_along(secciones))
    }
    names(secciones)[!nzchar(names(secciones))] <- paste0(
      "SECCION_",
      which(!nzchar(names(secciones)))
    )
  } else {
    if (!is.list(tablas) || !length(tablas)) {
      stop("`tablas` debe ser una lista no vacía.", call. = FALSE)
    }
  }

  if (!is.null(fila)) {
    fila <- as.character(fila)[1]
    if (!nzchar(fila) || !(fila %in% names(data))) {
      stop("`fila` debe ser una variable existente en `data`.", call. = FALSE)
    }
  }

  cruzar_con <- if (is.null(cruzar_con)) character(0) else as.character(cruzar_con)
  cruzar_con <- cruzar_con[!is.na(cruzar_con) & nzchar(trimws(cruzar_con))]
  cruzar_con <- unique(cruzar_con[cruzar_con %in% names(data)])
  if (!is.null(fila)) {
    cruzar_con <- setdiff(cruzar_con, fila)
  }

  if (is.null(weight_col) || !nzchar(as.character(weight_col))) {
    weight_col <- attr(data, "var_peso", exact = TRUE)
  }
  if (is.null(weight_col) || !nzchar(as.character(weight_col))) {
    weight_col <- if ("peso" %in% names(data)) "peso" else NA_character_
  }
  if (!is.na(weight_col) && !(weight_col %in% names(data))) {
    weight_col <- NA_character_
  }

  .rt_round_half_up <- function(x, dg = 0L) {
    if (!length(x)) return(x)
    s <- 10^dg
    out <- ifelse(
      is.na(x),
      NA_real_,
      ifelse(x >= 0, floor(x * s + 0.5), ceiling(x * s - 0.5)) / s
    )
    as.numeric(out)
  }

  .rt_numfmt <- function(dg = 0L) {
    if (dg <= 0) return("#,##0")
    paste0("#,##0.", paste(rep("0", dg), collapse = ""))
  }

  .rt_var_label <- function(var, map_local = NULL) {
    mapa <- if (!is.null(map_local)) map_local else nombres_indicadores
    if (!is.null(mapa) && var %in% names(mapa)) {
      return(as.character(mapa[[var]]))
    }
    if (!is.null(labels_override) && var %in% names(labels_override)) {
      return(as.character(labels_override[[var]]))
    }
    if (var %in% names(data)) {
      vl <- attr(data[[var]], "label", exact = TRUE)
      if (!is.null(vl) && nzchar(trimws(as.character(vl)))) {
        return(as.character(vl))
      }
    }
    if (!is.null(dic_vars) && all(c("name", "label") %in% names(dic_vars))) {
      lb <- dic_vars$label[dic_vars$name == var]
      if (length(lb) && !all(is.na(lb))) return(as.character(lb[1]))
    }
    as.character(var)
  }

  .rt_list_name <- function(var) {
    if (is.null(survey) || !is.data.frame(survey) ||
        !all(c("name", "list_name") %in% names(survey))) {
      return(NA_character_)
    }
    ln <- as.character(survey$list_name[as.character(survey$name) == var])
    ln <- ln[!is.na(ln) & nzchar(ln)]
    if (!length(ln)) return(NA_character_)
    ln[1]
  }

  .rt_levels <- function(var) {
    if (!(var %in% names(data))) {
      return(list(keys = character(0), labels = character(0)))
    }

    v_raw <- as.character(data[[var]])
    v_clean <- trimws(v_raw)
    v_clean <- v_clean[!is.na(v_clean) & nzchar(v_clean) & v_clean != "NA"]
    present <- unique(v_clean)

    ln <- .rt_list_name(var)
    ord_obj <- NULL
    if (!is.null(orders_list)) {
      if (var %in% names(orders_list)) {
        ord_obj <- orders_list[[var]]
      } else if (!is.na(ln) && ln %in% names(orders_list)) {
        ord_obj <- orders_list[[ln]]
      }
    }

    codes <- character(0)
    labels <- character(0)

    if (!is.null(ord_obj)) {
      if (!is.null(ord_obj$names))  codes <- as.character(ord_obj$names)
      if (!is.null(ord_obj$labels)) labels <- as.character(ord_obj$labels)
    }

    if (!length(codes)) {
      lb <- attr(data[[var]], "labels", exact = TRUE)
      if (!is.null(lb) && length(lb)) {
        codes <- as.character(names(lb))
        labels <- as.character(unname(lb))
      }
    }

    if (!length(codes)) {
      codes <- sort(unique(present))
      labels <- codes
    }

    keep <- !is.na(codes) & nzchar(trimws(codes))
    codes <- trimws(codes[keep])
    labels <- labels[keep]
    if (length(labels) != length(codes)) labels <- codes

    match_codes  <- if (length(codes))  sum(v_clean %in% codes) else 0L
    match_labels <- if (length(labels)) sum(v_clean %in% labels) else 0L

    use_labels_as_keys <- length(labels) && (match_labels > match_codes)
    keys <- if (use_labels_as_keys) labels else codes
    labs <- labels

    dedup <- !duplicated(keys)
    keys <- keys[dedup]
    labs <- labs[dedup]

    if (!isTRUE(mostrar_todo)) {
      in_data <- keys %in% present
      keys <- keys[in_data]
      labs <- labs[in_data]
    }

    extras <- setdiff(present, keys)
    if (length(extras)) {
      keys <- c(keys, extras)
      labs <- c(labs, extras)
    }

    list(keys = as.character(keys), labels = as.character(labs))
  }

  .rt_get_weights <- function() {
    if (is.na(weight_col)) return(rep(1, nrow(data)))
    w <- suppressWarnings(as.numeric(data[[weight_col]]))
    w[!is.finite(w) | is.na(w)] <- 0
    w
  }

  .rt_weighted_mean <- function(x_num, mask, w, dg = 1L) {
    idx <- mask & !is.na(mask) &
      is.finite(x_num) & !is.na(x_num) &
      is.finite(w) & !is.na(w) & w > 0

    if (!any(idx)) return(NA_real_)
    mu <- sum(x_num[idx] * w[idx], na.rm = TRUE) / sum(w[idx], na.rm = TRUE)
    .rt_round_half_up(mu, dg)
  }

  .rt_merge_runs <- function(v) {
    if (!length(v)) return(list())
    out <- list()
    ini <- 1L
    for (i in seq_along(v)) {
      if (i == length(v) || !identical(v[i], v[i + 1])) {
        out[[length(out) + 1L]] <- c(ini, i)
        ini <- i + 1L
      }
    }
    out
  }

  .rt_clean_vars <- function(x) {
    y <- as.character(x)
    y <- y[!is.na(y) & nzchar(trimws(y))]
    unique(y)
  }

  .rt_split_cruces <- function(plan) {
    if (modo_cruces != "separados") return(plan)
    out <- list()
    for (tb in plan) {
      cc <- .rt_clean_vars(tb$cruzar_con)
      if (is.null(tb$fila) || !nzchar(tb$fila)) {
        fila_tb <- NULL
      } else {
        fila_tb <- tb$fila
      }
      if (!is.null(fila_tb)) cc <- setdiff(cc, fila_tb)
      if (length(cc) <= 1L) {
        tb$cruzar_con <- cc
        out[[length(out) + 1L]] <- tb
      } else {
        for (cv in cc) {
          tb2 <- tb
          tb2$cruzar_con <- cv
          tb2$titulo <- paste0(tb$titulo, " - ", .rt_var_label(cv, tb$nombres_indicadores))
          out[[length(out) + 1L]] <- tb2
        }
      }
    }
    out
  }

  .rt_plan_from_secciones <- function() {
    plan <- list()
    for (sec in names(secciones)) {
      inds <- .rt_clean_vars(secciones[[sec]])
      if (!length(inds)) next
      plan[[length(plan) + 1L]] <- list(
        titulo = as.character(sec),
        indicadores = inds,
        fila = if (is.null(fila)) NULL else as.character(fila),
        cruzar_con = .rt_clean_vars(cruzar_con),
        nombres_indicadores = nombres_indicadores,
        agregar_brecha = isTRUE(agregar_brecha),
        etiqueta_brecha = as.character(etiqueta_brecha)[1],
        incluir_total = isTRUE(incluir_total)
      )
    }
    .rt_split_cruces(plan)
  }

  .rt_plan_from_tablas <- function(tablas_in) {
    nms <- names(tablas_in)
    plan <- list()
    for (i in seq_along(tablas_in)) {
      tb <- tablas_in[[i]]
      if (!is.list(tb)) next

      titulo_i <- tb$titulo
      if (is.null(titulo_i) || !nzchar(trimws(as.character(titulo_i)[1]))) {
        titulo_i <- if (!is.null(nms) && nzchar(nms[i])) nms[i] else paste0("TABLA_", i)
      }

      inds_i <- if (!is.null(tb$indicadores)) tb$indicadores else tb$secciones
      inds_i <- .rt_clean_vars(inds_i)
      if (!length(inds_i)) next

      fila_i <- if (!is.null(tb$fila)) as.character(tb$fila)[1] else if (!is.null(fila)) as.character(fila)[1] else NULL
      if (!is.null(fila_i) && (!nzchar(fila_i) || !(fila_i %in% names(data)))) {
        fila_i <- NULL
      }

      cruz_i <- if (!is.null(tb$cruzar_con)) tb$cruzar_con else cruzar_con
      cruz_i <- .rt_clean_vars(cruz_i)
      cruz_i <- cruz_i[cruz_i %in% names(data)]
      if (!is.null(fila_i)) cruz_i <- setdiff(cruz_i, fila_i)

      nind_i <- if (!is.null(tb$nombres_indicadores)) tb$nombres_indicadores else nombres_indicadores
      brecha_i <- if (!is.null(tb$agregar_brecha)) isTRUE(tb$agregar_brecha) else isTRUE(agregar_brecha)
      etb_i <- if (!is.null(tb$etiqueta_brecha)) as.character(tb$etiqueta_brecha)[1] else as.character(etiqueta_brecha)[1]
      inc_total_i <- if (!is.null(tb$incluir_total)) isTRUE(tb$incluir_total) else isTRUE(incluir_total)

      plan[[length(plan) + 1L]] <- list(
        titulo = as.character(titulo_i)[1],
        indicadores = inds_i,
        fila = fila_i,
        cruzar_con = cruz_i,
        nombres_indicadores = nind_i,
        agregar_brecha = brecha_i,
        etiqueta_brecha = etb_i,
        incluir_total = inc_total_i
      )
    }
    .rt_split_cruces(plan)
  }

  plan_tablas <- if (usar_tablas_plan) {
    .rt_plan_from_tablas(tablas)
  } else {
    .rt_plan_from_secciones()
  }
  if (!length(plan_tablas)) {
    stop("No se encontraron tablas válidas para exportar.", call. = FALSE)
  }

  st <- if (exists("mk_styles_cruces", mode = "function")) {
    mk_styles_cruces()
  } else {
    list(
      sec_title = openxlsx::createStyle(fontSize = 18, halign = "center", fontName = "Arial"),
      header = openxlsx::createStyle(fontSize = 10, textDecoration = "bold", border = c("top", "bottom"), fontName = "Arial", halign = "center"),
      header_A = openxlsx::createStyle(fontSize = 10, textDecoration = "bold", border = c("top", "bottom"), fontName = "Arial", halign = "left"),
      body_txt = openxlsx::createStyle(fontSize = 10, halign = "left", fontName = "Arial"),
      table_end = openxlsx::createStyle(border = "bottom", borderStyle = "thin", borderColour = "#000000"),
      note = openxlsx::createStyle(fontSize = 9, fontColour = "#666666", textDecoration = "italic", fontName = "Arial"),
      footer_top = openxlsx::createStyle(border = "top", borderStyle = "thin", borderColour = "#000000")
    )
  }
  style_num <- openxlsx::createStyle(
    fontSize = 10,
    halign = "right",
    valign = "center",
    numFmt = .rt_numfmt(digits),
    fontName = "Arial"
  )

  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, hoja)
  fila_excel <- 1L
  w <- .rt_get_weights()

  for (tb in plan_tablas) {
    sec <- as.character(tb$titulo)[1]
    vars_sec <- .rt_clean_vars(tb$indicadores)
    vars_sec <- vars_sec[vars_sec %in% names(data)]
    if (!length(vars_sec)) next

    fila_tb <- if (!is.null(tb$fila)) as.character(tb$fila)[1] else NULL
    if (!is.null(fila_tb) && (!nzchar(fila_tb) || !(fila_tb %in% names(data)))) {
      fila_tb <- NULL
    }

    cruzar_tb <- .rt_clean_vars(tb$cruzar_con)
    cruzar_tb <- cruzar_tb[cruzar_tb %in% names(data)]
    if (!is.null(fila_tb)) cruzar_tb <- setdiff(cruzar_tb, fila_tb)

    nombres_tb <- tb$nombres_indicadores
    agregar_brecha_tb <- isTRUE(tb$agregar_brecha)
    etiqueta_brecha_tb <- as.character(tb$etiqueta_brecha)[1]
    incluir_total_tb <- isTRUE(tb$incluir_total)

    x_cache <- lapply(vars_sec, function(v) suppressWarnings(as.numeric(data[[v]])))
    names(x_cache) <- vars_sec

    if (!is.null(fila_tb)) {
      lev_fila <- .rt_levels(fila_tb)
      row_keys <- lev_fila$keys
      row_labels <- lev_fila$labels
      if (!length(row_keys)) {
        row_keys <- "__SIN_DATOS__"
        row_labels <- "Sin datos"
      }
      v_fila <- trimws(as.character(data[[fila_tb]]))
      row_masks <- lapply(row_keys, function(k) !is.na(v_fila) & v_fila == k)
    } else {
      row_labels <- "Total"
      row_masks <- list(rep(TRUE, nrow(data)))
    }

    row_header <- if (!is.null(fila_tb)) .rt_var_label(fila_tb, nombres_tb) else "Grupo"
    usar_header_cruce <- length(cruzar_tb) > 0L

    col_plan <- list()
    if (usar_header_cruce) {
      blocks <- list()
      if (incluir_total_tb) {
        blocks[[length(blocks) + 1L]] <- list(
          h1 = "",
          h2 = "Total",
          mask = rep(TRUE, nrow(data))
        )
      }

      for (s in cruzar_tb) {
        lev_s <- .rt_levels(s)
        if (!length(lev_s$keys)) next
        v_s <- trimws(as.character(data[[s]]))
        s_lbl <- .rt_var_label(s, nombres_tb)
        for (j in seq_along(lev_s$keys)) {
          blocks[[length(blocks) + 1L]] <- list(
            h1 = s_lbl,
            h2 = as.character(lev_s$labels[j]),
            mask = !is.na(v_s) & v_s == lev_s$keys[j]
          )
        }
      }

      if (!length(blocks)) {
        usar_header_cruce <- FALSE
      } else {
        h1 <- c("")
        h2 <- c("")
        h3 <- c(row_header)
        for (b in blocks) {
          for (v in vars_sec) {
            h1 <- c(h1, b$h1)
            h2 <- c(h2, b$h2)
            h3 <- c(h3, .rt_var_label(v, nombres_tb))
            col_plan[[length(col_plan) + 1L]] <- list(var = v, mask = b$mask)
          }
        }
      }
    }

    if (!usar_header_cruce) {
      h1 <- NULL
      h2 <- NULL
      h3 <- c(row_header, vapply(vars_sec, .rt_var_label, character(1), map_local = nombres_tb))
      for (v in vars_sec) {
        col_plan[[length(col_plan) + 1L]] <- list(var = v, mask = rep(TRUE, nrow(data)))
      }
    }

    out <- data.frame(.fila = row_labels, stringsAsFactors = FALSE, check.names = FALSE)
    names(out)[1] <- row_header

    for (j in seq_along(col_plan)) {
      pj <- col_plan[[j]]
      xj <- x_cache[[pj$var]]
      vals <- vapply(seq_along(row_masks), function(i) {
        .rt_weighted_mean(
          x_num = xj,
          mask = row_masks[[i]] & pj$mask,
          w = w,
          dg = digits
        )
      }, numeric(1))
      out[[paste0("col_", j)]] <- as.numeric(vals)
    }

    if (agregar_brecha_tb && nrow(out) >= 2) {
      vals_brecha <- vapply(seq_len(ncol(out) - 1L), function(j) {
        vv <- suppressWarnings(as.numeric(out[[j + 1L]]))
        vv <- vv[is.finite(vv) & !is.na(vv)]
        if (!length(vv)) return(NA_real_)
        .rt_round_half_up(max(vv) - min(vv), digits)
      }, numeric(1))
      row_b <- as.list(c(etiqueta_brecha_tb, vals_brecha))
      names(row_b) <- names(out)
      out <- rbind(out, as.data.frame(row_b, stringsAsFactors = FALSE, check.names = FALSE))
      for (jj in 2:ncol(out)) out[[jj]] <- suppressWarnings(as.numeric(out[[jj]]))
    }

    ncols <- length(h3)
    sec_txt <- if (is.na(sec) || !nzchar(trimws(sec))) "TABLA" else sec

    openxlsx::writeData(wb, hoja, toupper(sec_txt), startRow = fila_excel, startCol = 1, colNames = FALSE)
    openxlsx::mergeCells(wb, hoja, rows = fila_excel, cols = 1:ncols)
    openxlsx::addStyle(wb, hoja, st$sec_title, rows = fila_excel, cols = 1:ncols, gridExpand = TRUE, stack = TRUE)
    fila_excel <- fila_excel + 1L

    if (usar_header_cruce) {
      openxlsx::writeData(wb, hoja, t(h1), startRow = fila_excel, startCol = 1, colNames = FALSE)
      openxlsx::writeData(wb, hoja, t(h2), startRow = fila_excel + 1L, startCol = 1, colNames = FALSE)
      openxlsx::writeData(wb, hoja, t(h3), startRow = fila_excel + 2L, startCol = 1, colNames = FALSE)

      openxlsx::addStyle(wb, hoja, st$header_A, rows = fila_excel:(fila_excel + 2L), cols = 1, gridExpand = TRUE, stack = TRUE)
      if (ncols > 1L) {
        openxlsx::addStyle(wb, hoja, st$header, rows = fila_excel:(fila_excel + 2L), cols = 2:ncols, gridExpand = TRUE, stack = TRUE)
      }

      runs_h1 <- .rt_merge_runs(h1)
      for (r in runs_h1) {
        if ((r[2] - r[1] + 1L) > 1L) {
          openxlsx::mergeCells(wb, hoja, rows = fila_excel, cols = r[1]:r[2])
        }
      }

      runs_h2 <- .rt_merge_runs(h2)
      for (r in runs_h2) {
        if ((r[2] - r[1] + 1L) > 1L) {
          openxlsx::mergeCells(wb, hoja, rows = fila_excel + 1L, cols = r[1]:r[2])
        }
      }
      row_ini <- fila_excel + 3L
    } else {
      openxlsx::writeData(wb, hoja, t(h3), startRow = fila_excel, startCol = 1, colNames = FALSE)
      openxlsx::addStyle(wb, hoja, st$header_A, rows = fila_excel, cols = 1, gridExpand = TRUE, stack = TRUE)
      if (ncols > 1L) {
        openxlsx::addStyle(wb, hoja, st$header, rows = fila_excel, cols = 2:ncols, gridExpand = TRUE, stack = TRUE)
      }
      row_ini <- fila_excel + 1L
    }
    openxlsx::writeData(wb, hoja, out, startRow = row_ini, startCol = 1, colNames = FALSE)
    row_fin <- row_ini + nrow(out) - 1L

    openxlsx::addStyle(wb, hoja, st$body_txt, rows = row_ini:row_fin, cols = 1, gridExpand = TRUE, stack = TRUE)
    if (ncols > 1L) {
      openxlsx::addStyle(wb, hoja, style_num, rows = row_ini:row_fin, cols = 2:ncols, gridExpand = TRUE, stack = TRUE)
    }
    openxlsx::addStyle(wb, hoja, st$table_end, rows = row_fin, cols = 1:ncols, gridExpand = TRUE, stack = TRUE)

    row_fuente <- row_fin + 1L
    txt_fuente <- paste0("Fuente: ", fuente)
    openxlsx::writeData(wb, hoja, txt_fuente, startRow = row_fuente, startCol = 1, colNames = FALSE)
    openxlsx::mergeCells(wb, hoja, rows = row_fuente, cols = 1:ncols)
    if (!is.null(st$note)) {
      openxlsx::addStyle(wb, hoja, st$note, rows = row_fuente, cols = 1:ncols, gridExpand = TRUE, stack = TRUE)
    }
    if (!is.null(st$footer_top)) {
      openxlsx::addStyle(wb, hoja, st$footer_top, rows = row_fuente, cols = 1:ncols, gridExpand = TRUE, stack = TRUE)
    }

    openxlsx::setColWidths(wb, hoja, cols = 1, widths = 34)
    if (ncols > 1L) openxlsx::setColWidths(wb, hoja, cols = 2:ncols, widths = 13)

    fila_excel <- row_fuente + separacion_filas + 1L
  }

  openxlsx::saveWorkbook(wb, path_xlsx, overwrite = TRUE)
  invisible(normalizePath(path_xlsx, winslash = "/"))
}
