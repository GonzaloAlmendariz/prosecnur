# =============================================================================
# prosecnur — Sistema unificado de indicadores, dimensiones y tablas analíticas
# Versión consolidada 2.0
# =============================================================================

# =============================================================================
# Operador base
# =============================================================================

`%||%` <- function(x, y) {
  if (is.null(x)) y else x
}

# =============================================================================
# Validadores y utilidades generales
# =============================================================================

.pn_es_texto_scalar <- function(x) {
  is.character(x) && length(x) == 1L && !is.na(x) && nzchar(trimws(x))
}

.pn_assert_texto_scalar <- function(x, arg) {
  if (!.pn_es_texto_scalar(x)) {
    stop("`", arg, "` debe ser un texto escalar no vacío.", call. = FALSE)
  }
  invisible(TRUE)
}

.pn_assert_lista <- function(x, arg) {
  if (!is.list(x)) {
    stop("`", arg, "` debe ser una lista.", call. = FALSE)
  }
  invisible(TRUE)
}

.pn_assert_df <- function(x, arg = "data") {
  if (!is.data.frame(x)) {
    stop("`", arg, "` debe ser un data.frame o tibble.", call. = FALSE)
  }
  invisible(TRUE)
}

.pn_compact_chr <- function(x) {
  x <- as.character(x)
  x <- x[!is.na(x)]
  x <- trimws(x)
  x[nzchar(x)]
}

.pn_null_if_empty <- function(x) {
  if (is.null(x)) return(NULL)
  if (length(x) == 0L) return(NULL)
  x
}

.pn_warn <- function(...) {
  warning(..., call. = FALSE)
}

.pn_msg <- function(...) {
  message(...)
}

.pn_safe_names <- function(x) {
  make.names(x, unique = TRUE)
}

# =============================================================================
# Instrumento / metadata
# =============================================================================

.pn_get_label_col_safe <- function(df) {
  if (is.null(df)) return(NULL)
  if ("label" %in% names(df)) return("label")
  cand <- grep("^label(::|$)", names(df), value = TRUE)
  if (length(cand)) return(cand[1])
  NULL
}

.pn_get_survey <- function(instrumento) {
  if (is.null(instrumento) || !is.list(instrumento)) return(NULL)
  instrumento$survey %||% NULL
}

.pn_get_choices <- function(instrumento) {
  if (is.null(instrumento) || !is.list(instrumento)) return(NULL)
  instrumento$choices %||% NULL
}

.pn_get_orders_list <- function(instrumento) {
  if (is.null(instrumento) || !is.list(instrumento)) return(NULL)
  instrumento$orders_list %||% NULL
}

.pn_get_list_name <- function(survey, var) {
  if (is.null(survey) || !all(c("name", "list_name") %in% names(survey))) {
    return(NA_character_)
  }
  i <- which(!is.na(survey$name) & as.character(survey$name) == var)[1]
  if (is.na(i)) return(NA_character_)
  ln <- as.character(survey$list_name[i])
  if (is.na(ln) || !nzchar(ln)) return(NA_character_)
  ln
}

pn_obtener_label_var <- function(var, instrumento = NULL, data = NULL, labels_override = NULL) {
  var <- trimws(as.character(var)[1])

  if (!is.null(labels_override) && var %in% names(labels_override)) {
    lab <- labels_override[[var]]
    if (!is.null(lab) && nzchar(trimws(as.character(lab)))) {
      return(as.character(lab))
    }
  }

  surv <- .pn_get_survey(instrumento)
  if (!is.null(surv) && "name" %in% names(surv)) {
    label_col <- .pn_get_label_col_safe(surv)
    if (!is.null(label_col) && label_col %in% names(surv)) {
      nm <- trimws(as.character(surv$name))
      i <- which(!is.na(nm) & nm == var)[1]
      if (!is.na(i)) {
        lab <- surv[[label_col]][i]
        if (!is.na(lab) && nzchar(trimws(as.character(lab)))) {
          return(as.character(lab))
        }
      }
    }
  }

  if (!is.null(data) && var %in% names(data)) {
    vl <- attr(data[[var]], "label", exact = TRUE)
    if (!is.null(vl) && nzchar(trimws(as.character(vl)))) {
      return(as.character(vl))
    }
  }

  var
}

.pn_get_dicc_code_to_label <- function(instrumento, var = NULL, list_name = NULL, data = NULL) {
  surv <- .pn_get_survey(instrumento)
  ch   <- .pn_get_choices(instrumento)
  label_col <- .pn_get_label_col_safe(ch)

  if (is.null(list_name) && !is.null(var)) {
    list_name <- .pn_get_list_name(surv, var)
  }

  if (!is.null(ch) &&
      all(c("list_name", "name") %in% names(ch)) &&
      !is.null(label_col) && label_col %in% names(ch) &&
      !is.na(list_name) && nzchar(list_name)) {
    ch_v <- ch[ch$list_name == list_name, , drop = FALSE]
    if (nrow(ch_v)) {
      return(stats::setNames(as.character(ch_v[[label_col]]), as.character(ch_v$name)))
    }
  }

  if (!is.null(data) && !is.null(var) && var %in% names(data)) {
    labs <- attr(data[[var]], "labels", exact = TRUE)
    if (!is.null(labs) && length(labs) > 0) {
      return(stats::setNames(as.character(unname(labs)), as.character(names(labs))))
    }
  }

  NULL
}

.pn_get_order_info <- function(var, instrumento = NULL, data = NULL) {
  ord  <- .pn_get_orders_list(instrumento)
  surv <- .pn_get_survey(instrumento)

  if (!is.null(ord)) {
    if (var %in% names(ord)) return(ord[[var]])
    ln <- .pn_get_list_name(surv, var)
    if (!is.na(ln) && ln %in% names(ord)) return(ord[[ln]])
  }

  dic <- .pn_get_dicc_code_to_label(instrumento = instrumento, var = var, data = data)
  if (!is.null(dic)) {
    return(list(
      names  = names(dic),
      labels = unname(dic),
      label  = pn_obtener_label_var(var, instrumento = instrumento, data = data)
    ))
  }

  NULL
}

# =============================================================================
# Pesos
# =============================================================================

.pn_get_pesos <- function(data, peso_var = NULL) {
  .pn_assert_df(data, "data")

  if (is.null(peso_var) || !nzchar(peso_var) || !(peso_var %in% names(data))) {
    return(rep(1, nrow(data)))
  }

  w <- suppressWarnings(as.numeric(data[[peso_var]]))
  w[!is.finite(w) | is.na(w)] <- 0
  w
}

# =============================================================================
# Helpers de formato / orden / categorías
# =============================================================================

.pn_auto_row_height <- function(text, chars_per_line = 70, base = 24, per_line = 16) {
  if (length(text) == 0 || is.na(text)) return(base)
  txt <- gsub("\\r?\\n", " ", as.character(text))
  lines <- max(1, ceiling(nchar(txt) / chars_per_line))
  base + (lines - 1) * per_line
}

.pn_move_ns_pref_last <- function(tab) {
  if (!nrow(tab) || !("Opciones" %in% names(tab))) return(tab)
  idx <- which(trimws(tab$Opciones) == "No sé / Prefiero no decir")
  if (length(idx) == 0) return(tab)
  dplyr::bind_rows(tab[-idx, , drop = FALSE], tab[idx, , drop = FALSE])
}

.pn_map_from_attr_labels <- function(tab, var, df) {
  if (is.null(df) || !(var %in% names(df))) return(tab)
  lab_attr <- attr(df[[var]], "labels", exact = TRUE)
  if (is.null(lab_attr) || length(lab_attr) == 0) return(tab)
  codes_vec  <- as.character(names(lab_attr))
  labels_vec <- as.character(unname(lab_attr))
  if (!("Opciones" %in% names(tab))) return(tab)

  is_total <- tab$Opciones %in% c("Total", "")
  body  <- if (any(is_total)) tab[!is_total, , drop = FALSE] else tab
  total <- if (any(is_total)) tab[ is_total, , drop = FALSE] else NULL

  idx <- match(as.character(body$Opciones), codes_vec)
  body$Opciones <- ifelse(!is.na(idx), labels_vec[idx], body$Opciones)

  if (!is.null(total) && nrow(total)) dplyr::bind_rows(body, total) else body
}

.pn_map_to_labels <- function(tab, var, orders_list) {
  if (is.null(orders_list)) return(tab)
  if (!("Opciones" %in% names(tab))) return(tab)

  is_total <- tab$Opciones == "Total"
  body  <- if (any(is_total)) tab[!is_total, , drop = FALSE] else tab
  total <- if (any(is_total)) tab[ is_total, , drop = FALSE] else NULL
  if (!nrow(body)) return(tab)

  ord_lbl <- tryCatch(orders_list[[var]]$labels, error = function(e) NULL)
  ord_nam <- tryCatch(orders_list[[var]]$names,  error = function(e) NULL)

  if (!is.null(ord_nam) && !is.null(ord_lbl)) {
    idx_code <- match(body$Opciones, ord_nam)
    body$Opciones <- ifelse(!is.na(idx_code), ord_lbl[idx_code], body$Opciones)
  }

  if (!is.null(total) && nrow(total)) dplyr::bind_rows(body, total) else body
}

.pn_completar_categorias <- function(tab, var, orders_list, denom = NULL,
                                     mostrar_todo = FALSE,
                                     codigos_solo_si_presentes = NULL) {

  if (!isTRUE(mostrar_todo)) return(tab)
  if (is.null(orders_list))  return(tab)
  if (!("Opciones" %in% names(tab))) return(tab)
  if (!(var %in% names(orders_list))) return(tab)

  codigos_cond_chr <- if (is.null(codigos_solo_si_presentes)) {
    character(0)
  } else {
    as.character(codigos_solo_si_presentes)
  }

  is_total <- tab$Opciones == "Total"
  body  <- if (any(is_total)) tab[!is_total, , drop = FALSE] else tab
  total <- if (any(is_total)) tab[ is_total, , drop = FALSE] else NULL

  if (!nrow(body)) return(tab)

  ord_entry <- orders_list[[var]]
  ord_lbl   <- tryCatch(ord_entry$labels, error = function(e) NULL)
  ord_nam   <- tryCatch(ord_entry$names,  error = function(e) NULL)

  if (is.null(ord_lbl)) return(tab)

  full_lbl <- as.character(ord_lbl)
  full_lbl <- full_lbl[!is.na(full_lbl) & nzchar(full_lbl)]

  faltan <- setdiff(full_lbl, body$Opciones)

  if (length(faltan)) {
    if (length(codigos_cond_chr) && !is.null(ord_nam)) {
      ord_lbl_chr <- as.character(ord_lbl)
      ord_nam_chr <- as.character(ord_nam)
      idx_faltan    <- match(faltan, ord_lbl_chr)
      codes_faltan  <- ord_nam_chr[idx_faltan]
      keep <- !(codes_faltan %in% codigos_cond_chr)
      faltan <- faltan[keep]
    }

    if (length(faltan)) {
      add <- tibble::tibble(
        Opciones = faltan,
        n        = 0,
        pct      = if (!is.null(denom) && denom > 0) 0 else NA_real_
      )
      body <- dplyr::bind_rows(body, add)
    }
  }

  body <- body |>
    dplyr::mutate(.orden_aux = match(Opciones, full_lbl)) |>
    dplyr::arrange(.orden_aux) |>
    dplyr::select(-.orden_aux)

  if (!is.null(total) && nrow(total)) dplyr::bind_rows(body, total) else body
}

.pn_reordenar_por_instrumento <- function(tab, var, orders_list) {
  if (is.null(orders_list) || !(var %in% names(orders_list))) return(tab)
  if (!all(c("Opciones", "n", "pct") %in% names(tab))) return(tab)

  is_total <- tab$Opciones == "Total"
  body  <- if (any(is_total)) tab[!is_total, , drop = FALSE] else tab
  total <- if (any(is_total)) tab[ is_total, , drop = FALSE] else NULL
  if (!nrow(body)) return(tab)

  ord_lbl <- tryCatch(orders_list[[var]]$labels, error = function(e) NULL)
  ord_nam <- tryCatch(orders_list[[var]]$names,  error = function(e) NULL)

  if (!is.null(ord_lbl)) {
    body <- dplyr::mutate(body, .orden_aux = match(Opciones, ord_lbl))
  } else {
    body$.orden_aux <- NA_integer_
  }

  if (all(is.na(body$.orden_aux)) && !is.null(ord_nam)) {
    body <- dplyr::mutate(body, .orden_aux = match(Opciones, ord_nam))
  }

  base_max <- suppressWarnings(max(body$.orden_aux, na.rm = TRUE))
  if (!is.finite(base_max)) base_max <- 0

  body <- body |>
    dplyr::mutate(
      .orden_aux = ifelse(
        is.na(.orden_aux),
        base_max + dplyr::row_number(),
        .orden_aux
      )
    ) |>
    dplyr::arrange(.orden_aux) |>
    dplyr::select(-.orden_aux)

  if (!is.null(total) && nrow(total)) dplyr::bind_rows(body, total) else body
}

# =============================================================================
# Detección SO / SM
# =============================================================================

.pn_split_sm_tokens <- function(x) {
  x <- as.character(x)
  lapply(x, function(xx) {
    if (is.na(xx) || !nzchar(xx) || xx == "NA") return(character(0))
    toks <- unlist(strsplit(xx, "\\s*[;\\s]+\\s*"))
    toks <- toks[nzchar(toks)]
    toks
  })
}

.pn_has_var_or_dummies <- function(data, var) {
  if (!is.data.frame(data)) return(FALSE)
  if (var %in% names(data)) return(TRUE)
  var_esc <- gsub("([\\W])", "\\\\\\1", var)
  any(grepl(paste0("^", var_esc, "[/\\.]"), names(data)))
}

.pn_has_only_dummies <- function(data, var) {
  if (!is.data.frame(data)) return(FALSE)
  if (var %in% names(data)) return(FALSE)
  var_esc <- gsub("([\\W])", "\\\\\\1", var)
  any(grepl(paste0("^", var_esc, "[/\\.]"), names(data)))
}

.pn_tipo_pregunta <- function(var, survey = NULL, sm_vars_force = NULL, data = NULL) {
  if (!is.null(sm_vars_force) && var %in% sm_vars_force) return("sm")

  if (!is.null(survey) && all(c("type", "name") %in% names(survey)) && any(survey$name == var)) {
    tps <- unique(stats::na.omit(as.character(survey$type[survey$name == var])))
    if (length(tps)) {
      if (any(grepl("^select_multiple(\\s|$)", tps))) return("sm")
      if (any(grepl("^select_one(\\s|$)", tps)))      return("so")
    }
  }

  if (!is.null(data) && .pn_has_only_dummies(data, var)) {
    return("sm")
  }

  "so"
}

.pn_col_sm_compact <- function(data, var) {
  v_orig <- paste0(var, "_ORIG")
  if (v_orig %in% names(data)) return(v_orig)
  if (var %in% names(data))    return(var)
  NA_character_
}

.pn_sm_compact_to_long <- function(x, id, w) {
  tibble::tibble(
    id    = id,
    valor = as.character(x),
    w     = as.numeric(w)
  ) |>
    tidyr::separate_rows(valor, sep = "\\s*;\\s*", convert = FALSE) |>
    dplyr::mutate(valor = trimws(valor)) |>
    dplyr::filter(!is.na(valor) & nzchar(valor) & valor != "NA")
}

pn_resolver_var_spec <- function(var_madre, instrumento = NULL, df = NULL) {
  data <- df
  inst <- instrumento

  if (is.null(data) || !is.data.frame(data)) {
    return(list(
      var_madre = var_madre,
      cols = character(0),
      map_code_to_label = list(),
      list_name = NA_character_,
      col_compact = NA_character_
    ))
  }

  var_esc <- gsub("([\\W])", "\\\\\\1", var_madre)
  pat_dum <- paste0("^", var_esc, "(\\.|_recod\\.|/)")
  cols <- grep(pat_dum, names(data), value = TRUE)

  col_compact <- NA_character_
  cand1 <- paste0(var_madre, "_ORIG")
  if (cand1 %in% names(data)) {
    col_compact <- cand1
  } else if (var_madre %in% names(data)) {
    col_compact <- var_madre
  }

  surv <- .pn_get_survey(inst)
  ch   <- .pn_get_choices(inst)

  list_name <- NA_character_
  if (!is.null(surv) && all(c("name", "list_name") %in% names(surv))) {
    i <- which(!is.na(surv$name) & surv$name == var_madre)[1]
    if (!is.na(i)) {
      list_name <- as.character(surv$list_name[i])
      if (is.na(list_name) || !nzchar(list_name)) list_name <- NA_character_
    }
  }

  map_code_to_label <- NULL
  label_col <- .pn_get_label_col_safe(ch)

  if (!is.null(ch) &&
      all(c("list_name", "name") %in% names(ch)) &&
      !is.null(label_col) && label_col %in% names(ch)) {
    if (!is.na(list_name) && nzchar(list_name)) {
      ch_v <- ch[ch$list_name == list_name, , drop = FALSE]
      if (nrow(ch_v)) {
        map_code_to_label <- stats::setNames(
          as.character(ch_v[[label_col]]),
          as.character(ch_v$name)
        )
      }
    }
  }

  if (is.null(map_code_to_label)) {
    cand_attr <- NULL
    if (!is.na(col_compact) && col_compact %in% names(data)) cand_attr <- col_compact
    if (is.null(cand_attr) && length(cols)) cand_attr <- cols[1]

    if (!is.null(cand_attr) && cand_attr %in% names(data)) {
      labs <- attr(data[[cand_attr]], "labels", exact = TRUE)
      if (!is.null(labs) && length(labs) > 0) {
        map_code_to_label <- stats::setNames(
          as.character(unname(labs)),
          as.character(names(labs))
        )
      }
    }
  }

  if (is.null(map_code_to_label)) map_code_to_label <- character(0)

  dummy_code <- function(x) {
    sub(paste0("^", var_esc, "(\\.|_recod\\.|/)"), "", x)
  }

  dummy_codes <- if (length(cols)) dummy_code(cols) else character(0)

  codes_order <- character(0)
  if (length(map_code_to_label) > 0) {
    codes_order <- as.character(names(map_code_to_label))
  }

  if (!length(codes_order) && !is.na(col_compact) && col_compact %in% names(data)) {
    x <- as.character(data[[col_compact]])
    x <- x[!is.na(x) & nzchar(x) & x != "NA"]
    if (length(x)) {
      vals <- unlist(strsplit(x, "\\s*;\\s*"), use.names = FALSE)
      vals <- trimws(vals)
      vals <- vals[!is.na(vals) & nzchar(vals) & vals != "NA"]
      codes_order <- unique(vals)
    }
  }

  if (!length(codes_order) && length(dummy_codes)) {
    codes_order <- unique(dummy_codes)
  }

  if (length(codes_order)) {
    suppressWarnings(num <- as.numeric(codes_order))
    if (!all(is.na(num))) {
      ord <- order(is.na(num), num, codes_order)
      codes_order <- codes_order[ord]
    } else {
      codes_order <- sort(codes_order)
    }
  }

  if (length(cols) && length(codes_order)) {
    ord_idx <- match(dummy_codes, codes_order)
    if (all(is.na(ord_idx))) {
      ord_idx <- seq_along(cols)
    } else {
      nf <- is.na(ord_idx)
      if (any(nf)) {
        base <- suppressWarnings(max(ord_idx, na.rm = TRUE))
        if (!is.finite(base)) base <- 0
        ord_idx[nf] <- base + seq_len(sum(nf))
      }
    }
    cols <- cols[order(ord_idx)]
  }

  if (length(dummy_codes)) {
    falt <- setdiff(dummy_codes, names(map_code_to_label))
    if (length(falt)) {
      extra <- stats::setNames(falt, falt)
      map_code_to_label <- c(map_code_to_label, extra)
    }
  }

  list(
    var_madre = var_madre,
    cols = cols,
    map_code_to_label = as.list(map_code_to_label),
    list_name = list_name,
    col_compact = col_compact
  )
}

# =============================================================================
# Resúmenes numéricos
# =============================================================================

.pn_resumen_numerico_w <- function(x, w, probs = c(.25, .5, .75), digits = 1) {
  x <- suppressWarnings(as.numeric(x))
  w <- suppressWarnings(as.numeric(w))
  labs <- c(
    "Casos válidos",
    "Promedio",
    "Desviación estándar",
    "Mínimo",
    "Percentil 25",
    "Mediana (Percentil 50)",
    "Percentil 75",
    "Máximo"
  )

  idx <- is.finite(x) & !is.na(x) & is.finite(w) & !is.na(w) & w > 0
  if (!any(idx)) {
    return(tibble::tibble(
      Estadistico = labs,
      Valor = c(0, rep(NA_real_, 7))
    ))
  }

  x <- x[idx]
  w <- w[idx]
  n_val <- length(x)

  mu <- stats::weighted.mean(x, w, na.rm = TRUE)
  wsum <- sum(w)
  var_w <- if (wsum > 0) sum(w * (x - mu)^2) / wsum else NA_real_
  sd_w  <- sqrt(var_w)

  ord <- order(x)
  x2 <- x[ord]
  w2 <- w[ord]
  cw <- cumsum(w2) / sum(w2)

  wq <- function(p) {
    j <- which(cw >= p)[1]
    if (is.na(j)) NA_real_ else x2[j]
  }

  tibble::tibble(
    Estadistico = labs,
    Valor = c(
      n_val,
      round(mu, digits),
      round(sd_w, digits),
      round(min(x2), digits),
      round(wq(probs[1]), digits),
      round(wq(probs[2]), digits),
      round(wq(probs[3]), digits),
      round(max(x2), digits)
    )
  )
}

.pn_resumen_numerico_w_mask <- function(x, w, mask,
                                        probs = c(.25, .5, .75),
                                        digits = 1) {
  x <- suppressWarnings(as.numeric(x))
  w <- suppressWarnings(as.numeric(w))
  mask <- as.logical(mask)

  idx <- mask & is.finite(x) & !is.na(x) & is.finite(w) & !is.na(w) & w > 0
  if (!any(idx)) {
    return(c(
      N = 0,
      Media = NA_real_,
      SD = NA_real_,
      Min = NA_real_,
      P25 = NA_real_,
      Mediana = NA_real_,
      P75 = NA_real_,
      Max = NA_real_
    ))
  }

  x <- x[idx]
  w <- w[idx]

  n_val <- length(x)
  mu <- stats::weighted.mean(x, w, na.rm = TRUE)
  wsum <- sum(w)
  var_w <- if (wsum > 0) sum(w * (x - mu)^2) / wsum else NA_real_
  sd_w  <- sqrt(var_w)

  ord <- order(x)
  x2 <- x[ord]
  w2 <- w[ord]
  cw <- cumsum(w2) / sum(w2)

  wq <- function(p) {
    j <- which(cw >= p)[1]
    if (is.na(j)) NA_real_ else x2[j]
  }

  c(
    N       = n_val,
    Media   = round(mu, digits),
    SD      = round(sd_w, digits),
    Min     = round(min(x2), digits),
    P25     = round(wq(probs[1]), digits),
    Mediana = round(wq(probs[2]), digits),
    P75     = round(wq(probs[3]), digits),
    Max     = round(max(x2), digits)
  )
}

# =============================================================================
# Constructores declarativos
# =============================================================================

#' Crear una regla de normalización
#'
#' Define cómo normalizar un conjunto de variables dentro de un plan analítico.
#'
#' @param variables Vector de nombres de variables.
#' @param metodo Método de normalización: `"minmax"`, `"z"`, `"rango_teorico"` o `"ninguna"`.
#' @param a,b Límite inferior y superior para escalamiento.
#' @param minimo,maximo Límites teóricos para `metodo = "rango_teorico"`.
#' @param invertir Variables que deben invertirse luego de normalizar.
#' @param aplicar_a_todas Si `TRUE`, aplica a todas las variables elegibles.
#' @param excluir Variables a excluir cuando `aplicar_a_todas = TRUE`.
#' @param solo_numericas Si `TRUE`, restringe a variables numéricas.
#' @param ignorar_faltantes Si `TRUE`, ignora valores faltantes en el cálculo.
#' @param id Identificador de la regla.
#' @param notas Texto libre para documentación.
#'
#' @return Objeto de clase `pn_regla_normalizacion`.
#' @export
pn_regla_normalizacion <- function(
    variables = NULL,
    metodo = c("minmax", "z", "rango_teorico", "ninguna"),
    a = 0,
    b = 100,
    minimo = NULL,
    maximo = NULL,
    invertir = NULL,
    aplicar_a_todas = FALSE,
    excluir = NULL,
    solo_numericas = TRUE,
    ignorar_faltantes = TRUE,
    id = NULL,
    notas = NULL
) {
  metodo <- match.arg(metodo)

  if (!is.null(variables)) variables <- .pn_compact_chr(variables)
  if (!is.null(excluir)) excluir <- .pn_compact_chr(excluir)
  if (!is.null(invertir)) invertir <- .pn_compact_chr(invertir)

  if (!isTRUE(aplicar_a_todas) && (is.null(variables) || !length(variables))) {
    stop("Debe indicarse `variables` o usar `aplicar_a_todas = TRUE`.", call. = FALSE)
  }

  out <- list(
    clase = "regla_normalizacion",
    id = id %||% "normalizacion",
    variables = variables,
    metodo = metodo,
    a = a,
    b = b,
    minimo = minimo,
    maximo = maximo,
    invertir = invertir,
    aplicar_a_todas = isTRUE(aplicar_a_todas),
    excluir = excluir,
    solo_numericas = isTRUE(solo_numericas),
    ignorar_faltantes = isTRUE(ignorar_faltantes),
    notas = notas %||% character(0)
  )
  class(out) <- c("pn_regla_normalizacion", "regla_normalizacion", "list")
  out
}

#' Crear un item de indicador
#'
#' @param variable Nombre de la variable.
#' @param peso Peso del item en la agregación.
#' @param invertir Si `TRUE`, invierte el sentido del item.
#' @param usar_normalizada Si `TRUE`, usa la versión normalizada de la variable.
#' @param incluir Si `TRUE`, incluye el item en el cálculo.
#' @param etiqueta Etiqueta legible del item.
#' @param notas Texto libre para documentación.
#'
#' @return Objeto de clase `pn_item_indicador`.
#' @export
pn_item_indicador <- function(
    variable,
    peso = 1,
    invertir = FALSE,
    usar_normalizada = TRUE,
    incluir = TRUE,
    etiqueta = NULL,
    notas = NULL
) {
  .pn_assert_texto_scalar(variable, "variable")

  out <- list(
    clase = "item_indicador",
    variable = variable,
    peso = as.numeric(peso)[1],
    invertir = isTRUE(invertir),
    usar_normalizada = isTRUE(usar_normalizada),
    incluir = isTRUE(incluir),
    etiqueta = etiqueta,
    notas = notas %||% character(0)
  )
  class(out) <- c("pn_item_indicador", "item_indicador", "list")
  out
}

#' Crear una dimensión o índice
#'
#' @param id Identificador de la dimensión.
#' @param titulo Título legible.
#' @param items Lista de objetos `pn_item_indicador()`.
#' @param agregacion Método de agregación.
#' @param minimo_items Mínimo de items válidos requeridos por fila.
#' @param estandarizar_resultado Si `TRUE`, reescala el resultado al rango indicado.
#' @param rango_resultado Rango objetivo de salida al estandarizar.
#' @param crear_variable Si `TRUE`, crea la variable de salida en el dataset.
#' @param notas Texto libre para documentación.
#'
#' @return Objeto de clase `pn_dimension_indicador`.
#' @export
pn_dimension_indicador <- function(
    id,
    titulo = NULL,
    items,
    agregacion = c("promedio", "suma", "media_ponderada", "suma_ponderada"),
    minimo_items = 1,
    estandarizar_resultado = FALSE,
    rango_resultado = c(0, 100),
    crear_variable = TRUE,
    notas = NULL
) {
  .pn_assert_texto_scalar(id, "id")
  .pn_assert_lista(items, "items")

  if (!length(items)) stop("`items` debe contener al menos un elemento.", call. = FALSE)

  ok_items <- vapply(items, inherits, logical(1), what = "item_indicador")
  if (!all(ok_items)) {
    stop("Todos los elementos de `items` deben ser creados con `pn_item_indicador()`.", call. = FALSE)
  }

  agregacion <- match.arg(agregacion)

  out <- list(
    clase = "dimension_indicador",
    id = id,
    titulo = titulo %||% id,
    items = items,
    agregacion = agregacion,
    minimo_items = as.integer(minimo_items)[1],
    estandarizar_resultado = isTRUE(estandarizar_resultado),
    rango_resultado = as.numeric(rango_resultado),
    crear_variable = isTRUE(crear_variable),
    notas = notas %||% character(0)
  )
  class(out) <- c("pn_dimension_indicador", "dimension_indicador", "list")
  out
}

#' Crear una definición de tabla simple
#'
#' @param id Identificador de la tabla.
#' @param variable Variable principal de análisis.
#' @param tipo Tipo de tabla.
#' @param titulo Título legible.
#' @param filtro Expresión lógica en texto para filtrar filas.
#' @param cruzar_por Vector de variables de cruce (opcional).
#' @param incluir_total Si `TRUE`, agrega fila total.
#' @param mostrar_todo Si `TRUE`, completa categorías no observadas.
#' @param codigos_solo_si_presentes Códigos que solo se muestran si aparecen en los datos.
#' @param opciones_excluir Códigos/etiquetas a excluir.
#' @param notas Texto libre para documentación.
#'
#' @return Objeto de clase `pn_tabla_simple`.
#' @export
pn_tabla_simple <- function(
    id,
    variable,
    tipo = c("frecuencia", "resumen_numerico", "media", "conteo", "proporcion",
             "media_por_grupo", "resumen_por_grupo"),
    titulo = NULL,
    filtro = NULL,
    cruzar_por = NULL,
    incluir_total = TRUE,
    mostrar_todo = FALSE,
    codigos_solo_si_presentes = NULL,
    opciones_excluir = NULL,
    notas = NULL
) {
  .pn_assert_texto_scalar(id, "id")
  .pn_assert_texto_scalar(variable, "variable")
  tipo <- match.arg(tipo)

  out <- list(
    clase = "tabla_simple",
    id = id,
    variable = variable,
    tipo = tipo,
    titulo = titulo %||% id,
    filtro = filtro,
    cruzar_por = .pn_null_if_empty(.pn_compact_chr(cruzar_por)),
    incluir_total = isTRUE(incluir_total),
    mostrar_todo = isTRUE(mostrar_todo),
    codigos_solo_si_presentes = .pn_null_if_empty(as.character(codigos_solo_si_presentes)),
    opciones_excluir = .pn_null_if_empty(as.character(opciones_excluir)),
    notas = notas %||% character(0)
  )
  class(out) <- c("pn_tabla_simple", "tabla_simple", "list")
  out
}

#' Crear una fila para tabla de condiciones
#'
#' @param id Identificador de la fila.
#' @param etiqueta Etiqueta visible en la salida.
#' @param condicion Condición lógica en texto.
#' @param notas Texto libre para documentación.
#'
#' @return Objeto de clase `pn_fila_condicion`.
#' @export
pn_fila_condicion <- function(id, etiqueta, condicion, notas = NULL) {
  .pn_assert_texto_scalar(id, "id")
  .pn_assert_texto_scalar(etiqueta, "etiqueta")
  .pn_assert_texto_scalar(condicion, "condicion")

  out <- list(
    clase = "fila_condicion",
    id = id,
    etiqueta = etiqueta,
    condicion = condicion,
    notas = notas %||% character(0)
  )
  class(out) <- c("pn_fila_condicion", "fila_condicion", "list")
  out
}

#' Crear una columna para tabla de condiciones
#'
#' @param id Identificador de la columna.
#' @param etiqueta Etiqueta visible en la salida.
#' @param tipo Tipo de cálculo de la columna.
#' @param variable Variable de referencia para columnas de media/suma.
#' @param referencia_total_fila Referencia de total para proporciones.
#' @param notas Texto libre para documentación.
#'
#' @return Objeto de clase `pn_columna_condicion`.
#' @export
pn_columna_condicion <- function(
    id,
    etiqueta,
    tipo = c("conteo_cond_fila", "proporcion_sobre_total", "media_cond_fila", "suma_cond_fila"),
    variable = NULL,
    referencia_total_fila = NULL,
    notas = NULL
) {
  .pn_assert_texto_scalar(id, "id")
  .pn_assert_texto_scalar(etiqueta, "etiqueta")
  tipo <- match.arg(tipo)

  out <- list(
    clase = "columna_condicion",
    id = id,
    etiqueta = etiqueta,
    tipo = tipo,
    variable = variable,
    referencia_total_fila = referencia_total_fila,
    notas = notas %||% character(0)
  )
  class(out) <- c("pn_columna_condicion", "columna_condicion", "list")
  out
}

#' Crear una tabla por condiciones
#'
#' @param id Identificador de la tabla.
#' @param titulo Título legible.
#' @param filas Lista de objetos `pn_fila_condicion()`.
#' @param columnas Lista de objetos `pn_columna_condicion()`.
#' @param filtro Expresión lógica en texto para filtrar filas.
#' @param notas Texto libre para documentación.
#'
#' @return Objeto de clase `pn_tabla_condiciones`.
#' @export
pn_tabla_condiciones <- function(
    id,
    titulo = NULL,
    filas,
    columnas,
    filtro = NULL,
    notas = NULL
) {
  .pn_assert_texto_scalar(id, "id")
  .pn_assert_lista(filas, "filas")
  .pn_assert_lista(columnas, "columnas")

  if (!length(filas)) stop("`filas` no puede estar vacío.", call. = FALSE)
  if (!length(columnas)) stop("`columnas` no puede estar vacío.", call. = FALSE)

  ok_filas <- vapply(filas, inherits, logical(1), what = "fila_condicion")
  ok_cols  <- vapply(columnas, inherits, logical(1), what = "columna_condicion")

  if (!all(ok_filas)) stop("Todos los elementos de `filas` deben crearse con `pn_fila_condicion()`.", call. = FALSE)
  if (!all(ok_cols)) stop("Todos los elementos de `columnas` deben crearse con `pn_columna_condicion()`.", call. = FALSE)

  out <- list(
    clase = "tabla_condiciones",
    id = id,
    titulo = titulo %||% id,
    filas = filas,
    columnas = columnas,
    filtro = filtro,
    notas = notas %||% character(0)
  )
  class(out) <- c("pn_tabla_condiciones", "tabla_condiciones", "list")
  out
}

#' Crear un grupo conceptual
#'
#' @param id Identificador del grupo.
#' @param variables Variables que componen el grupo.
#' @param notas Texto libre para documentación.
#'
#' @return Objeto de clase `pn_grupo_conceptual`.
#' @export
pn_grupo_conceptual <- function(id, variables, notas = NULL) {
  .pn_assert_texto_scalar(id, "id")
  variables <- .pn_compact_chr(variables)
  if (!length(variables)) stop("`variables` debe contener al menos una variable.", call. = FALSE)

  out <- list(
    clase = "grupo_conceptual",
    id = id,
    variables = variables,
    notas = notas %||% character(0)
  )
  class(out) <- c("pn_grupo_conceptual", "grupo_conceptual", "list")
  out
}

#' Crear una fila conceptual
#'
#' @param id Identificador de la fila.
#' @param etiqueta Etiqueta visible.
#' @param grupos Lista de objetos `pn_grupo_conceptual()`.
#' @param notas Texto libre para documentación.
#'
#' @return Objeto de clase `pn_fila_conceptual`.
#' @export
pn_fila_conceptual <- function(id, etiqueta, grupos, notas = NULL) {
  .pn_assert_texto_scalar(id, "id")
  .pn_assert_texto_scalar(etiqueta, "etiqueta")
  .pn_assert_lista(grupos, "grupos")

  if (!length(grupos)) stop("`grupos` no puede estar vacío.", call. = FALSE)
  ok <- vapply(grupos, inherits, logical(1), what = "grupo_conceptual")
  if (!all(ok)) stop("Todos los elementos de `grupos` deben crearse con `pn_grupo_conceptual()`.", call. = FALSE)

  out <- list(
    clase = "fila_conceptual",
    id = id,
    etiqueta = etiqueta,
    grupos = grupos,
    notas = notas %||% character(0)
  )
  class(out) <- c("pn_fila_conceptual", "fila_conceptual", "list")
  out
}

#' Crear una columna conceptual
#'
#' @param id Identificador de la columna.
#' @param etiqueta Etiqueta visible.
#' @param tipo Tipo de cálculo.
#' @param referencia Referencia adicional según `tipo`.
#' @param condicion Condición lógica (texto) para cálculo condicional.
#' @param numerador,denominador Referencias de numerador/denominador para proporciones.
#' @param notas Texto libre para documentación.
#'
#' @return Objeto de clase `pn_columna_conceptual`.
#' @export
pn_columna_conceptual <- function(
    id,
    etiqueta,
    tipo = c("suma", "conteo_cond", "proporcion_cond", "proporcion_rel",
             "media", "mediana", "minimo", "maximo"),
    referencia = NULL,
    condicion = NULL,
    numerador = NULL,
    denominador = NULL,
    notas = NULL
) {
  .pn_assert_texto_scalar(id, "id")
  .pn_assert_texto_scalar(etiqueta, "etiqueta")
  tipo <- match.arg(tipo)

  out <- list(
    clase = "columna_conceptual",
    id = id,
    etiqueta = etiqueta,
    tipo = tipo,
    referencia = referencia,
    condicion = condicion,
    numerador = numerador,
    denominador = denominador,
    notas = notas %||% character(0)
  )
  class(out) <- c("pn_columna_conceptual", "columna_conceptual", "list")
  out
}

#' Crear una tabla conceptual
#'
#' @param id Identificador de la tabla.
#' @param titulo Título legible.
#' @param filas Lista de objetos `pn_fila_conceptual()`.
#' @param columnas Lista de objetos `pn_columna_conceptual()`.
#' @param filtro Expresión lógica en texto para filtrar filas.
#' @param notas Texto libre para documentación.
#'
#' @return Objeto de clase `pn_tabla_conceptual`.
#' @export
pn_tabla_conceptual <- function(
    id,
    titulo = NULL,
    filas,
    columnas,
    filtro = NULL,
    notas = NULL
) {
  .pn_assert_texto_scalar(id, "id")
  .pn_assert_lista(filas, "filas")
  .pn_assert_lista(columnas, "columnas")

  if (!length(filas)) stop("`filas` no puede estar vacío.", call. = FALSE)
  if (!length(columnas)) stop("`columnas` no puede estar vacío.", call. = FALSE)

  ok_f <- vapply(filas, inherits, logical(1), what = "fila_conceptual")
  ok_c <- vapply(columnas, inherits, logical(1), what = "columna_conceptual")

  if (!all(ok_f)) stop("Todos los elementos de `filas` deben crearse con `pn_fila_conceptual()`.", call. = FALSE)
  if (!all(ok_c)) stop("Todos los elementos de `columnas` deben crearse con `pn_columna_conceptual()`.", call. = FALSE)

  out <- list(
    clase = "tabla_conceptual",
    id = id,
    titulo = titulo %||% id,
    filas = filas,
    columnas = columnas,
    filtro = filtro,
    notas = notas %||% character(0)
  )
  class(out) <- c("pn_tabla_conceptual", "tabla_conceptual", "list")
  out
}

#' Crear especificación de cruce de indicador
#'
#' @param id Identificador del cruce.
#' @param variable Variable principal.
#' @param cruces Variables de estratificación.
#' @param tipo_tabla Tipo de tabla a calcular.
#' @param titulo Título legible.
#' @param filtro Expresión lógica en texto para filtrar filas.
#' @param mostrar_significancia Si `TRUE`, calcula comparación entre columnas.
#' @param alpha Nivel de significancia para pruebas entre estratos.
#' @param opciones_excluir Códigos/etiquetas a excluir.
#' @param codigos_solo_si_presentes Códigos que solo se muestran si aparecen en datos.
#' @param notas Texto libre para documentación.
#'
#' @return Objeto de clase `pn_cruce_indicador`.
#' @export
pn_cruce_indicador <- function(
    id,
    variable,
    cruces,
    tipo_tabla = c("frecuencia", "resumen_numerico"),
    titulo = NULL,
    filtro = NULL,
    mostrar_significancia = TRUE,
    alpha = 0.05,
    opciones_excluir = NULL,
    codigos_solo_si_presentes = NULL,
    notas = NULL
) {
  .pn_assert_texto_scalar(id, "id")
  .pn_assert_texto_scalar(variable, "variable")
  cruces <- .pn_compact_chr(cruces)
  if (!length(cruces)) stop("`cruces` debe contener al menos una variable de cruce.", call. = FALSE)

  tipo_tabla <- match.arg(tipo_tabla)

  out <- list(
    clase = "cruce_indicador",
    id = id,
    variable = variable,
    cruces = cruces,
    tipo_tabla = tipo_tabla,
    titulo = titulo %||% id,
    filtro = filtro,
    mostrar_significancia = isTRUE(mostrar_significancia),
    alpha = as.numeric(alpha)[1],
    opciones_excluir = .pn_null_if_empty(as.character(opciones_excluir)),
    codigos_solo_si_presentes = .pn_null_if_empty(as.character(codigos_solo_si_presentes)),
    notas = notas %||% character(0)
  )
  class(out) <- c("pn_cruce_indicador", "cruce_indicador", "list")
  out
}

#' Construir un plan analítico
#'
#' @param nombre Nombre del plan.
#' @param peso Nombre de la variable de ponderación.
#' @param instrumento Lista con metadatos del instrumento (survey/choices/orders).
#' @param normalizacion Lista de reglas `pn_regla_normalizacion()`.
#' @param dimensiones Lista de dimensiones `pn_dimension_indicador()`.
#' @param tablas Lista de tablas (`pn_tabla_simple()`, `pn_tabla_condiciones()`, `pn_tabla_conceptual()`).
#' @param cruces Lista de cruces `pn_cruce_indicador()`.
#' @param etiquetas Lista opcional de etiquetas personalizadas.
#' @param notas Texto libre para documentación.
#'
#' @return Objeto de clase `pn_plan_analitico`.
#' @export
pn_plan_analitico <- function(
    nombre = "Plan analítico",
    peso = NULL,
    instrumento = NULL,
    normalizacion = NULL,
    dimensiones = NULL,
    tablas = NULL,
    cruces = NULL,
    etiquetas = NULL,
    notas = NULL
) {
  if (!.pn_es_texto_scalar(nombre)) nombre <- "Plan analítico"

  if (!is.null(normalizacion)) {
    if (inherits(normalizacion, "regla_normalizacion")) normalizacion <- list(normalizacion)
    .pn_assert_lista(normalizacion, "normalizacion")
  }

  if (!is.null(dimensiones)) {
    .pn_assert_lista(dimensiones, "dimensiones")
    ok <- vapply(dimensiones, inherits, logical(1), what = "dimension_indicador")
    if (!all(ok)) stop("Todos los elementos de `dimensiones` deben crearse con `pn_dimension_indicador()`.", call. = FALSE)
  }

  if (!is.null(tablas)) {
    .pn_assert_lista(tablas, "tablas")
    ok <- vapply(
      tablas,
      function(x) inherits(x, "tabla_simple") ||
        inherits(x, "tabla_condiciones") ||
        inherits(x, "tabla_conceptual"),
      logical(1)
    )
    if (!all(ok)) {
      stop("`tablas` solo puede contener objetos `pn_tabla_simple()`, `pn_tabla_condiciones()` o `pn_tabla_conceptual()`.", call. = FALSE)
    }
  }

  if (!is.null(cruces)) {
    .pn_assert_lista(cruces, "cruces")
    ok <- vapply(cruces, inherits, logical(1), what = "cruce_indicador")
    if (!all(ok)) stop("Todos los elementos de `cruces` deben crearse con `pn_cruce_indicador()`.", call. = FALSE)
  }

  out <- list(
    clase = "plan_analitico",
    nombre = nombre,
    peso = peso,
    instrumento = instrumento,
    normalizacion = normalizacion %||% list(),
    dimensiones = dimensiones %||% list(),
    tablas = tablas %||% list(),
    cruces = cruces %||% list(),
    etiquetas = etiquetas %||% list(),
    notas = notas %||% character(0)
  )
  class(out) <- c("pn_plan_analitico", "plan_analitico", "list")
  out
}

# =============================================================================
# Helpers de evaluación lógica
# =============================================================================

.pn_valido_helper <- function(..., .n) {
  args <- list(...)
  if (length(args) == 0L) return(rep(TRUE, .n))
  masks <- lapply(args, function(v) !is.na(v))
  Reduce("&", masks)
}

.pn_todo_prefijo_helper <- function(data, prefijo, valor = 1) {
  vars_pref <- grep(paste0("^", prefijo), names(data), value = TRUE)
  if (length(vars_pref) == 0L) return(rep(FALSE, nrow(data)))
  sub <- data[, vars_pref, drop = FALSE]
  apply(sub, 1L, function(row) {
    if (all(is.na(row))) FALSE else all(!is.na(row) & row == valor)
  })
}

.pn_valido_prefijo_helper <- function(data, prefijo) {
  vars_pref <- grep(paste0("^", prefijo), names(data), value = TRUE)
  if (length(vars_pref) == 0L) return(rep(FALSE, nrow(data)))
  sub <- data[, vars_pref, drop = FALSE]
  apply(!is.na(sub), 1L, any)
}

.pn_eval_condicion_fila <- function(data, condicion) {
  if (is.null(condicion) || !nzchar(condicion)) {
    return(rep(TRUE, nrow(data)))
  }

  env <- list2env(as.list(data), parent = parent.frame())
  env$valido <- function(...) .pn_valido_helper(..., .n = nrow(data))
  env$valid  <- env$valido
  env$todo_prefijo <- function(prefijo, valor = 1) {
    .pn_todo_prefijo_helper(data, prefijo = prefijo, valor = valor)
  }
  env$valido_prefijo <- function(prefijo) {
    .pn_valido_prefijo_helper(data, prefijo = prefijo)
  }

  res <- try(eval(parse(text = condicion), envir = env), silent = TRUE)

  if (inherits(res, "try-error")) {
    .pn_warn(
      "No se pudo evaluar la condición '", condicion,
      "'. Se devolverá NA.\nDetalle: ",
      conditionMessage(attr(res, "condition"))
    )
    return(rep(NA, nrow(data)))
  }

  if (!is.logical(res) || length(res) != nrow(data)) {
    stop("La condición '", condicion, "' no devolvió un vector lógico de longitud nrow(data).", call. = FALSE)
  }

  res
}

.pn_eval_condicion_vector <- function(x, condicion) {
  if (is.null(condicion) || !nzchar(condicion)) return(!is.na(x))
  if (identical(condicion, "no_es_na")) return(!is.na(x))

  expr <- if (grepl("\\bx\\b", condicion)) condicion else paste0("x ", condicion)
  res <- eval(parse(text = expr), envir = list(x = x))

  if (!is.logical(res) || length(res) != length(x)) {
    stop("La condición '", condicion, "' no devolvió un vector lógico de longitud length(x).", call. = FALSE)
  }

  res
}

# =============================================================================
# Normalización
# =============================================================================

.pn_variable_es_numericable <- function(x) {
  if (is.numeric(x)) return(TRUE)
  sx <- suppressWarnings(as.numeric(x))
  any(!is.na(sx))
}

.pn_obtener_variables_normalizacion <- function(regla, data) {
  if (isTRUE(regla$aplicar_a_todas)) {
    vars <- setdiff(names(data), .pn_compact_chr(regla$excluir))
    if (isTRUE(regla$solo_numericas)) {
      vars <- vars[vapply(data[vars], .pn_variable_es_numericable, logical(1))]
    }
    return(vars)
  }
  .pn_compact_chr(regla$variables)
}

.pn_normalizar_minmax <- function(x, a = 0, b = 100, minimo = NULL, maximo = NULL) {
  x <- suppressWarnings(as.numeric(x))
  if (all(is.na(x))) return(rep(NA_real_, length(x)))

  lo <- minimo %||% suppressWarnings(min(x, na.rm = TRUE))
  hi <- maximo %||% suppressWarnings(max(x, na.rm = TRUE))

  if (!is.finite(lo) || !is.finite(hi) || hi == lo) {
    return(rep(NA_real_, length(x)))
  }

  a + ((x - lo) / (hi - lo)) * (b - a)
}

.pn_normalizar_z <- function(x) {
  x <- suppressWarnings(as.numeric(x))
  if (all(is.na(x))) return(rep(NA_real_, length(x)))
  mu <- suppressWarnings(mean(x, na.rm = TRUE))
  sdv <- suppressWarnings(stats::sd(x, na.rm = TRUE))
  if (!is.finite(sdv) || sdv == 0) return(rep(NA_real_, length(x)))
  (x - mu) / sdv
}

.pn_normalizar_rango_teorico <- function(x, minimo, maximo, a = 0, b = 100) {
  x <- suppressWarnings(as.numeric(x))
  if (is.null(minimo) || is.null(maximo)) {
    stop("Para `metodo = 'rango_teorico'` deben indicarse `minimo` y `maximo`.", call. = FALSE)
  }
  if (!is.finite(minimo) || !is.finite(maximo) || maximo == minimo) {
    return(rep(NA_real_, length(x)))
  }
  a + ((x - minimo) / (maximo - minimo)) * (b - a)
}

.pn_aplicar_una_regla_normalizacion <- function(data, regla, prefijo = "norm_") {
  vars_obj <- .pn_obtener_variables_normalizacion(regla, data)
  vars_obj <- vars_obj[vars_obj %in% names(data)]

  res <- data
  meta <- list()

  if (!length(vars_obj)) {
    return(list(data = res, meta = meta))
  }

  for (v in vars_obj) {
    x <- data[[v]]
    out_name <- paste0(prefijo, v)

    x_num <- suppressWarnings(as.numeric(x))
    if (all(is.na(x_num))) {
      res[[out_name]] <- NA_real_
      meta[[v]] <- list(origen = v, normalizada = out_name, metodo = regla$metodo)
      next
    }

    y <- switch(
      regla$metodo,
      "minmax" = .pn_normalizar_minmax(
        x_num, a = regla$a, b = regla$b,
        minimo = regla$minimo, maximo = regla$maximo
      ),
      "z" = .pn_normalizar_z(x_num),
      "rango_teorico" = .pn_normalizar_rango_teorico(
        x_num, minimo = regla$minimo, maximo = regla$maximo,
        a = regla$a, b = regla$b
      ),
      "ninguna" = x_num
    )

    if (!is.null(regla$invertir) && v %in% .pn_compact_chr(regla$invertir)) {
      if (regla$metodo %in% c("minmax", "rango_teorico")) {
        y <- regla$a + regla$b - y
      } else {
        y <- -1 * y
      }
    }

    res[[out_name]] <- y
    attr(res[[out_name]], "label") <- paste0(pn_obtener_label_var(v, data = data), " (normalizada)")
    meta[[v]] <- list(
      origen = v,
      normalizada = out_name,
      metodo = regla$metodo,
      invertir = !is.null(regla$invertir) && v %in% .pn_compact_chr(regla$invertir)
    )
  }

  list(data = res, meta = meta)
}

pn_aplicar_normalizacion <- function(data, reglas, prefijo = "norm_") {
  .pn_assert_df(data, "data")

  if (is.null(reglas)) return(list(data = data, meta = list()))
  if (inherits(reglas, "regla_normalizacion")) reglas <- list(reglas)
  .pn_assert_lista(reglas, "reglas")

  res <- data
  meta_all <- list()

  for (rg in reglas) {
    if (!inherits(rg, "regla_normalizacion")) {
      stop("Cada elemento de `reglas` debe crearse con `pn_regla_normalizacion()`.", call. = FALSE)
    }
    tmp <- .pn_aplicar_una_regla_normalizacion(res, rg, prefijo = prefijo)
    res <- tmp$data
    meta_all <- c(meta_all, tmp$meta)
  }

  list(data = res, meta = meta_all)
}

# =============================================================================
# Dimensiones / índices
# =============================================================================

.pn_resolver_nombre_item <- function(item, data, prefijo_norm = "norm_") {
  v <- item$variable
  vn <- paste0(prefijo_norm, v)

  if (isTRUE(item$usar_normalizada) && vn %in% names(data)) return(vn)
  if (v %in% names(data)) return(v)
  vn
}

.pn_calcular_dimension_unica <- function(data, dim_obj, prefijo_norm = "norm_") {
  items <- dim_obj$items
  vars <- vapply(items, .pn_resolver_nombre_item, character(1), data = data, prefijo_norm = prefijo_norm)
  vars_exist <- vars[vars %in% names(data)]

  if (!length(vars_exist)) {
    y <- rep(NA_real_, nrow(data))
    return(list(data = data, variable = dim_obj$id, valores = y))
  }

  X <- as.data.frame(data[, vars_exist, drop = FALSE])
  X[] <- lapply(X, function(v) suppressWarnings(as.numeric(v)))

  for (i in seq_along(items)) {
    nm <- vars[i]
    if (!(nm %in% names(X))) next
    it <- items[[i]]

    if (isTRUE(it$invertir)) {
      rng <- suppressWarnings(range(X[[nm]], na.rm = TRUE))
      if (all(is.finite(rng)) && diff(rng) > 0) {
        X[[nm]] <- rng[1] + rng[2] - X[[nm]]
      } else {
        X[[nm]] <- -1 * X[[nm]]
      }
    }
  }

  pesos_items <- vapply(items, function(it) as.numeric(it$peso)[1], numeric(1))
  names(pesos_items) <- vars
  pesos_items <- pesos_items[names(X)]
  pesos_items[!is.finite(pesos_items) | is.na(pesos_items)] <- 1

  n_valid <- rowSums(!is.na(X))

  y <- switch(
    dim_obj$agregacion,
    "promedio" = rowMeans(X, na.rm = TRUE),
    "suma" = rowSums(X, na.rm = TRUE),
    "media_ponderada" = {
      num <- rowSums(sweep(X, 2, pesos_items, `*`), na.rm = TRUE)
      den <- rowSums(sweep(!is.na(X), 2, pesos_items, `*`), na.rm = TRUE)
      ifelse(den > 0, num / den, NA_real_)
    },
    "suma_ponderada" = rowSums(sweep(X, 2, pesos_items, `*`), na.rm = TRUE)
  )

  y[n_valid < dim_obj$minimo_items] <- NA_real_

  if (isTRUE(dim_obj$estandarizar_resultado)) {
    y <- .pn_normalizar_minmax(
      y,
      a = dim_obj$rango_resultado[1],
      b = dim_obj$rango_resultado[2]
    )
  }

  res <- data
  if (isTRUE(dim_obj$crear_variable)) {
    res[[dim_obj$id]] <- y
    attr(res[[dim_obj$id]], "label") <- dim_obj$titulo
  }

  list(data = res, variable = dim_obj$id, valores = y)
}

pn_calcular_dimensiones <- function(data, dimensiones, prefijo_norm = "norm_") {
  .pn_assert_df(data, "data")

  if (is.null(dimensiones) || !length(dimensiones)) {
    return(list(data = data, meta = list()))
  }

  res <- data
  meta <- list()

  for (dm in dimensiones) {
    if (!inherits(dm, "dimension_indicador")) {
      stop("Todos los elementos de `dimensiones` deben crearse con `pn_dimension_indicador()`.", call. = FALSE)
    }

    tmp <- .pn_calcular_dimension_unica(res, dm, prefijo_norm = prefijo_norm)
    res <- tmp$data
    meta[[dm$id]] <- list(
      id = dm$id,
      titulo = dm$titulo,
      agregacion = dm$agregacion,
      variable = dm$id
    )
  }

  list(data = res, meta = meta)
}

# =============================================================================
# Frecuencias
# =============================================================================

pn_freq_table <- function(data, var, survey = NULL, sm_vars_force = NULL,
                          orders_list = NULL, mostrar_todo = FALSE,
                          codigos_solo_si_presentes = NULL,
                          peso_var = NULL) {
  .pn_assert_df(data, "data")

  has_main <- var %in% names(data)
  var_escaped   <- gsub("([\\W])", "\\\\\\1", var)
  subvars_slash <- names(data)[grepl(paste0("^", var_escaped, "/"), names(data))]
  subvars_dot   <- names(data)[grepl(paste0("^", var_escaped, "\\.[^.]+$"), names(data))]
  subvars_all <- c(subvars_slash, subvars_dot)
  has_dummies <- length(subvars_all) > 0L

  if (!has_main && !has_dummies) {
    stop("`", var, "` no existe en `data` ni se detectaron dummies asociadas.", call. = FALSE)
  }

  tipo <- .pn_tipo_pregunta(var, survey, sm_vars_force, data = data)
  if (tipo != "sm" && has_dummies) tipo <- "sm"

  w <- .pn_get_pesos(data, peso_var = peso_var)

  if (tipo == "sm") {
    if (has_main && (is.character(data[[var]]) || is.factor(data[[var]]))) {
      vec <- as.character(data[[var]])
      df_long <- tibble::tibble(id = seq_len(nrow(data)), valor = vec) |>
        dplyr::filter(!is.na(valor) & nzchar(valor) & valor != "NA") |>
        dplyr::mutate(tokens = .pn_split_sm_tokens(valor)) |>
        dplyr::select(-valor) |>
        tidyr::unnest_longer(tokens, values_to = "op") |>
        dplyr::mutate(op = trimws(op)) |>
        dplyr::filter(nzchar(op)) |>
        dplyr::distinct(id, op)

      if (!nrow(df_long)) {
        return(tibble::tibble(Opciones = character(), n = numeric(), pct = numeric()))
      }

      ids_con_marca <- sort(unique(df_long$id))
      denom <- sum(w[ids_con_marca], na.rm = TRUE)

      tab <- df_long |>
        dplyr::left_join(tibble::tibble(id = seq_len(nrow(data)), peso = w), by = "id") |>
        dplyr::group_by(op) |>
        dplyr::summarise(n = sum(peso, na.rm = TRUE), .groups = "drop") |>
        dplyr::arrange(dplyr::desc(n)) |>
        dplyr::transmute(Opciones = op, n = as.numeric(n), pct = if (denom > 0) n / denom else NA_real_)

      tab <- .pn_map_from_attr_labels(tab, var, data)
      tab <- .pn_map_to_labels(tab, var, orders_list)
      tab <- .pn_completar_categorias(tab, var, orders_list, denom, mostrar_todo, codigos_solo_si_presentes)
      tab <- .pn_reordenar_por_instrumento(tab, var, orders_list)
      tab <- .pn_move_ns_pref_last(tab)

      total_row <- tibble::tibble(Opciones = "Total", n = as.numeric(denom), pct = 1)
      return(dplyr::bind_rows(tab, total_row))
    }

    if (!length(subvars_all)) {
      return(tibble::tibble(Opciones = character(), n = numeric(), pct = numeric()))
    }

    mat <- as.data.frame(data[, subvars_all, drop = FALSE])
    mat[] <- lapply(mat, function(v) suppressWarnings(as.numeric(as.character(v))))

    has_any <- rowSums(mat == 1, na.rm = TRUE) > 0
    denom   <- sum(w[has_any], na.rm = TRUE)

    n_w <- vapply(subvars_all, function(sv) {
      v <- suppressWarnings(as.numeric(as.character(mat[[sv]])))
      sum(w[v == 1 & !is.na(v)], na.rm = TRUE)
    }, numeric(1))

    tab <- tibble::tibble(subvar = subvars_all, n = as.numeric(n_w)) |>
      dplyr::mutate(Opciones = sub(paste0("^", var_escaped, "[/\\.]"), "", subvar)) |>
      dplyr::arrange(dplyr::desc(n)) |>
      dplyr::transmute(Opciones, n, pct = if (denom > 0) n / denom else NA_real_)

    tab <- .pn_map_to_labels(tab, var, orders_list)
    tab <- .pn_completar_categorias(tab, var, orders_list, denom, mostrar_todo, codigos_solo_si_presentes)
    tab <- .pn_reordenar_por_instrumento(tab, var, orders_list)
    tab <- .pn_move_ns_pref_last(tab)

    total_row <- tibble::tibble(Opciones = "Total", n = as.numeric(denom), pct = 1)
    return(dplyr::bind_rows(tab, total_row))
  }

  if (!has_main) {
    stop("`", var, "` no existe como columna en `data`.", call. = FALSE)
  }

  tib <- data |>
    dplyr::transmute(.op = as.character(.data[[var]]), peso = w) |>
    dplyr::filter(!is.na(.op) & nzchar(.op) & .op != "NA")

  if (!nrow(tib)) {
    return(tibble::tibble(Opciones = character(), n = numeric(), pct = numeric()))
  }

  denom <- sum(tib$peso, na.rm = TRUE)

  tab <- tib |>
    dplyr::group_by(.op) |>
    dplyr::summarise(n = sum(peso, na.rm = TRUE), .groups = "drop") |>
    dplyr::arrange(dplyr::desc(n)) |>
    dplyr::mutate(pct = if (denom > 0) n / denom else NA_real_) |>
    dplyr::rename(Opciones = .op)

  tab <- .pn_map_from_attr_labels(tab, var, data)
  tab <- .pn_map_to_labels(tab, var, orders_list)
  tab <- .pn_completar_categorias(tab, var, orders_list, denom, mostrar_todo, codigos_solo_si_presentes)
  tab <- .pn_reordenar_por_instrumento(tab, var, orders_list)
  tab <- .pn_move_ns_pref_last(tab)

  total_row <- tibble::tibble(Opciones = "Total", n = sum(tab$n, na.rm = TRUE), pct = 1)
  dplyr::bind_rows(tab, total_row)
}

# =============================================================================
# Tablas simples
# =============================================================================

.pn_get_categorias <- function(var, data, survey = NULL, orders_list = NULL, opciones_excluir = NULL) {
  x <- if (var %in% names(data)) data[[var]] else NULL
  lab_attr <- if (!is.null(x)) attr(x, "labels", exact = TRUE) else NULL

  ln <- .pn_get_list_name(survey, var)
  codes  <- character(0)
  labels <- character(0)

  obj <- NULL
  if (!is.null(orders_list)) {
    if (var %in% names(orders_list)) {
      obj <- orders_list[[var]]
    } else if (!is.na(ln) && ln %in% names(orders_list)) {
      obj <- orders_list[[ln]]
    }
  }

  if (!is.null(obj)) {
    codes  <- as.character(obj$names)
    labels <- as.character(obj$labels)
  } else if (!is.null(lab_attr) && length(lab_attr) > 0) {
    codes  <- names(lab_attr)
    labels <- as.character(unname(lab_attr))
  } else if (!is.null(x)) {
    codes  <- sort(unique(na.omit(as.character(x))))
    labels <- codes
  }

  ok <- !is.na(codes) & nzchar(codes)
  codes  <- codes[ok]
  labels <- labels[ok]

  if (!is.null(opciones_excluir) && length(opciones_excluir) > 0) {
    ok <- !(labels %in% opciones_excluir)
    codes  <- codes[ok]
    labels <- labels[ok]
  }

  list(codes = codes, labels = labels, list_name = ln)
}

pn_calcular_tabla_simple <- function(data, tabla_obj, instrumento = NULL, peso = NULL, sm_vars_force = NULL) {
  df <- data

  if (!is.null(tabla_obj$filtro) && nzchar(tabla_obj$filtro)) {
    mask <- .pn_eval_condicion_fila(df, tabla_obj$filtro)
    mask[is.na(mask)] <- FALSE
    df <- df[mask, , drop = FALSE]
  }

  survey <- .pn_get_survey(instrumento)
  orders_list <- .pn_get_orders_list(instrumento)
  var <- tabla_obj$variable
  tipo <- tabla_obj$tipo

  if (!nrow(df)) {
    return(list(id = tabla_obj$id, titulo = tabla_obj$titulo, tipo = tipo, tabla = data.frame()))
  }

  if (tipo == "frecuencia") {
    tab <- pn_freq_table(
      data = df,
      var = var,
      survey = survey,
      sm_vars_force = sm_vars_force,
      orders_list = orders_list,
      mostrar_todo = tabla_obj$mostrar_todo,
      codigos_solo_si_presentes = tabla_obj$codigos_solo_si_presentes,
      peso_var = peso
    )

    if (!is.null(tabla_obj$opciones_excluir) && nrow(tab)) {
      excl <- as.character(tabla_obj$opciones_excluir)
      is_total <- tab$Opciones == "Total"
      body <- tab[!is_total, , drop = FALSE]
      total <- tab[is_total, , drop = FALSE]
      body <- body[!(body$Opciones %in% excl), , drop = FALSE]
      tab <- dplyr::bind_rows(body, total)
    }

    if (!isTRUE(tabla_obj$incluir_total) && nrow(tab)) {
      tab <- tab[tab$Opciones != "Total", , drop = FALSE]
    }

    return(list(
      id = tabla_obj$id,
      titulo = tabla_obj$titulo,
      tipo = tipo,
      variable = var,
      tabla = tab
    ))
  }

  w <- .pn_get_pesos(df, peso_var = peso)

  if (tipo == "resumen_numerico") {
    tab <- .pn_resumen_numerico_w(df[[var]], w)
    return(list(
      id = tabla_obj$id,
      titulo = tabla_obj$titulo,
      tipo = tipo,
      variable = var,
      tabla = tab
    ))
  }

  x <- suppressWarnings(as.numeric(df[[var]]))
  idx_valid <- !is.na(x)

  if (tipo == "conteo") {
    valor <- sum(w[idx_valid], na.rm = TRUE)
    tab <- tibble::tibble(Indicador = tabla_obj$titulo, Valor = valor)
    return(list(id = tabla_obj$id, titulo = tabla_obj$titulo, tipo = tipo, variable = var, tabla = tab))
  }

  if (tipo == "media") {
    valor <- if (any(idx_valid)) stats::weighted.mean(x[idx_valid], w[idx_valid], na.rm = TRUE) else NA_real_
    tab <- tibble::tibble(Indicador = tabla_obj$titulo, Valor = valor)
    return(list(id = tabla_obj$id, titulo = tabla_obj$titulo, tipo = tipo, variable = var, tabla = tab))
  }

  if (tipo == "proporcion") {
    valor <- if (any(idx_valid)) sum(w[idx_valid & x == 1], na.rm = TRUE) / sum(w[idx_valid], na.rm = TRUE) else NA_real_
    tab <- tibble::tibble(Indicador = tabla_obj$titulo, Valor = valor)
    return(list(id = tabla_obj$id, titulo = tabla_obj$titulo, tipo = tipo, variable = var, tabla = tab))
  }

  if (tipo %in% c("media_por_grupo", "resumen_por_grupo")) {
    grupo <- tabla_obj$cruzar_por[1] %||% NA_character_
    if (is.na(grupo) || !(grupo %in% names(df))) {
      stop("`cruzar_por` debe contener una variable válida para `", tipo, "`.", call. = FALSE)
    }

    cats <- .pn_get_categorias(grupo, df, survey = survey, orders_list = orders_list, opciones_excluir = NULL)
    estr_codes  <- cats$codes
    estr_labels <- cats$labels
    if (!length(estr_codes)) {
      return(list(id = tabla_obj$id, titulo = tabla_obj$titulo, tipo = tipo, variable = var, tabla = data.frame()))
    }

    v_estr <- as.character(df[[grupo]])
    usa_codes  <- any(v_estr %in% estr_codes)
    usa_labels <- any(v_estr %in% estr_labels)
    keys_vec   <- if (usa_codes || !usa_labels) estr_codes else estr_labels

    if (tipo == "media_por_grupo") {
      filas <- lapply(seq_along(keys_vec), function(j) {
        mask_j <- !is.na(v_estr) & v_estr == keys_vec[j]
        xx <- x[mask_j]
        ww <- w[mask_j]
        idx <- !is.na(xx)
        tibble::tibble(
          Grupo = estr_labels[j],
          Valor = if (any(idx)) stats::weighted.mean(xx[idx], ww[idx], na.rm = TRUE) else NA_real_
        )
      })
      tab <- dplyr::bind_rows(filas)
    } else {
      filas <- lapply(seq_along(keys_vec), function(j) {
        mask_j <- !is.na(v_estr) & v_estr == keys_vec[j]
        rs <- .pn_resumen_numerico_w_mask(x, w, mask_j)
        tibble::tibble(
          Grupo = estr_labels[j],
          N = rs["N"],
          Media = rs["Media"],
          SD = rs["SD"],
          Min = rs["Min"],
          P25 = rs["P25"],
          Mediana = rs["Mediana"],
          P75 = rs["P75"],
          Max = rs["Max"]
        )
      })
      tab <- dplyr::bind_rows(filas)
    }

    return(list(
      id = tabla_obj$id,
      titulo = tabla_obj$titulo,
      tipo = tipo,
      variable = var,
      tabla = tab
    ))
  }

  stop("Tipo de tabla simple no soportado: ", tipo, call. = FALSE)
}

# =============================================================================
# Tablas por condiciones
# =============================================================================

pn_calcular_tabla_condiciones <- function(data, tabla_obj, peso = NULL) {
  df <- data

  if (!is.null(tabla_obj$filtro) && nzchar(tabla_obj$filtro)) {
    mask <- .pn_eval_condicion_fila(df, tabla_obj$filtro)
    mask[is.na(mask)] <- FALSE
    df <- df[mask, , drop = FALSE]
  }

  pesos <- .pn_get_pesos(df, peso_var = peso)
  filas_cfg <- tabla_obj$filas
  cols_cfg  <- tabla_obj$columnas

  n_por_fila <- numeric(length(filas_cfg))
  names(n_por_fila) <- vapply(filas_cfg, function(f) f$id, FUN.VALUE = character(1))

  for (i in seq_along(filas_cfg)) {
    f <- filas_cfg[[i]]
    cond <- .pn_eval_condicion_fila(df, f$condicion)
    cond_log <- cond
    cond_log[is.na(cond_log)] <- FALSE
    n_por_fila[i] <- sum(pesos[cond_log], na.rm = TRUE)
  }

  res_mat <- matrix(
    NA_real_,
    nrow = length(filas_cfg),
    ncol = length(cols_cfg),
    dimnames = list(
      vapply(filas_cfg, function(f) f$etiqueta, FUN.VALUE = character(1)),
      vapply(cols_cfg,  function(c) c$etiqueta, FUN.VALUE = character(1))
    )
  )

  ids_filas <- vapply(filas_cfg, function(f) f$id, FUN.VALUE = character(1))

  for (j in seq_along(cols_cfg)) {
    col_def <- cols_cfg[[j]]
    tipo    <- col_def$tipo

    if (tipo == "conteo_cond_fila") {
      res_mat[, j] <- n_por_fila

    } else if (tipo == "proporcion_sobre_total") {
      ref_id <- col_def$referencia_total_fila
      if (is.null(ref_id)) {
        stop("En `proporcion_sobre_total` se requiere `referencia_total_fila`.", call. = FALSE)
      }
      idx_total <- match(ref_id, ids_filas)
      if (is.na(idx_total)) {
        stop("No se encontró la fila de referencia `", ref_id, "`.", call. = FALSE)
      }
      denom <- n_por_fila[idx_total]
      res_mat[, j] <- ifelse(denom > 0, n_por_fila / denom, NA_real_)

    } else if (tipo %in% c("media_cond_fila", "suma_cond_fila")) {
      var_ref <- col_def$variable
      if (is.null(var_ref) || !(var_ref %in% names(df))) {
        stop("En `", tipo, "` se requiere una `variable` válida.", call. = FALSE)
      }
      x <- suppressWarnings(as.numeric(df[[var_ref]]))

      vals <- vapply(seq_along(filas_cfg), function(i) {
        cond <- .pn_eval_condicion_fila(df, filas_cfg[[i]]$condicion)
        cond[is.na(cond)] <- FALSE
        idx <- cond & !is.na(x)
        if (!any(idx)) return(NA_real_)

        if (tipo == "media_cond_fila") {
          stats::weighted.mean(x[idx], pesos[idx], na.rm = TRUE)
        } else {
          sum(x[idx] * pesos[idx], na.rm = TRUE)
        }
      }, numeric(1))

      res_mat[, j] <- vals

    } else {
      .pn_warn("Tipo de columna no soportado en tabla_condiciones: ", tipo)
    }
  }

  tabla <- as.data.frame(res_mat, check.names = FALSE)
  tabla <- cbind(
    Fila = vapply(filas_cfg, function(f) f$etiqueta, FUN.VALUE = character(1)),
    tabla
  )

  list(
    id = tabla_obj$id,
    titulo = tabla_obj$titulo,
    tipo = "tabla_condiciones",
    tabla = tabla
  )
}

# =============================================================================
# Tablas conceptuales
# =============================================================================

.pn_calc_vector_grupo_conceptual <- function(data, grupo_obj) {
  vars_g <- grupo_obj$variables
  faltan <- vars_g[!vars_g %in% names(data)]

  if (length(faltan)) {
    .pn_warn(
      "En el grupo `", grupo_obj$id,
      "` faltan variables: ", paste(faltan, collapse = ", "),
      ". Se usará NA."
    )
    return(rep(NA_real_, nrow(data)))
  }

  if (length(vars_g) == 1L) {
    return(data[[vars_g]])
  }

  sub <- data[, vars_g, drop = FALSE]
  sub[] <- lapply(sub, function(v) suppressWarnings(as.numeric(v)))
  rowSums(sub, na.rm = TRUE)
}

pn_calcular_tabla_conceptual <- function(data, tabla_obj, peso = NULL) {
  df <- data

  if (!is.null(tabla_obj$filtro) && nzchar(tabla_obj$filtro)) {
    mask <- .pn_eval_condicion_fila(df, tabla_obj$filtro)
    mask[is.na(mask)] <- FALSE
    df <- df[mask, , drop = FALSE]
  }

  pesos <- .pn_get_pesos(df, peso_var = peso)
  filas_cfg <- tabla_obj$filas
  cols_cfg  <- tabla_obj$columnas

  filas_labels <- vapply(filas_cfg, function(f) f$etiqueta, FUN.VALUE = character(1))

  res_mat <- matrix(
    NA_real_,
    nrow = length(filas_cfg),
    ncol = length(cols_cfg),
    dimnames = list(
      filas_labels,
      vapply(cols_cfg, function(c) c$etiqueta, FUN.VALUE = character(1))
    )
  )

  for (i in seq_along(filas_cfg)) {
    fila_cfg <- filas_cfg[[i]]

    grupos <- list()
    for (g in fila_cfg$grupos) {
      grupos[[g$id]] <- .pn_calc_vector_grupo_conceptual(df, g)
    }

    vals_fila <- list()

    for (j in seq_along(cols_cfg)) {
      col_def <- cols_cfg[[j]]
      tipo    <- col_def$tipo
      valor   <- NA_real_

      if (tipo %in% c("suma", "conteo_cond", "proporcion_cond", "media", "mediana", "minimo", "maximo")) {

        ref <- col_def$referencia
        if (is.null(ref) || !startsWith(ref, "@")) {
          stop("En tabla_conceptual, `referencia` debe apuntar a un grupo tipo '@nombre'.", call. = FALSE)
        }
        g_name <- substring(ref, 2L)

        if (is.null(grupos[[g_name]])) {
          .pn_warn("No se encontró el grupo `", g_name, "` en la fila ", fila_cfg$id, ".")
          x <- rep(NA_real_, nrow(df))
        } else {
          x <- grupos[[g_name]]
        }

        w <- pesos
        mask_base <- !is.na(x)

        condicion <- col_def$condicion
        if (!is.null(condicion) && nzchar(condicion)) {
          cond_vec <- .pn_eval_condicion_vector(x, condicion)
          mask_base <- mask_base & cond_vec
        }

        if (tipo == "suma") {
          valor <- sum(as.numeric(x[mask_base]) * w[mask_base], na.rm = TRUE)

        } else if (tipo == "conteo_cond") {
          valor <- sum(w[mask_base], na.rm = TRUE)

        } else if (tipo == "proporcion_cond") {
          num   <- sum(w[mask_base], na.rm = TRUE)
          denom <- sum(w[!is.na(x)],  na.rm = TRUE)
          valor <- if (denom > 0) num / denom else NA_real_

        } else if (tipo == "media") {
          xx <- suppressWarnings(as.numeric(x))
          num   <- sum(xx[mask_base] * w[mask_base], na.rm = TRUE)
          denom <- sum(w[mask_base], na.rm = TRUE)
          valor <- if (denom > 0) num / denom else NA_real_

        } else if (tipo == "mediana") {
          xx <- suppressWarnings(as.numeric(x))
          valor <- stats::median(xx[!is.na(xx)], na.rm = TRUE)

        } else if (tipo == "minimo") {
          xx <- suppressWarnings(as.numeric(x))
          valor <- suppressWarnings(min(xx, na.rm = TRUE))

        } else if (tipo == "maximo") {
          xx <- suppressWarnings(as.numeric(x))
          valor <- suppressWarnings(max(xx, na.rm = TRUE))
        }

      } else if (tipo == "proporcion_rel") {
        num_id <- col_def$numerador
        den_id <- col_def$denominador
        if (is.null(vals_fila[[num_id]]) || is.null(vals_fila[[den_id]])) {
          stop("En `proporcion_rel`, `numerador` y `denominador` deben haberse calculado antes.", call. = FALSE)
        }
        num <- vals_fila[[num_id]]
        den <- vals_fila[[den_id]]
        valor <- if (den > 0) num / den else NA_real_

      } else {
        .pn_warn("Tipo de columna no soportado en tabla_conceptual: ", tipo)
      }

      vals_fila[[col_def$id]] <- valor
      res_mat[i, j] <- valor
    }
  }

  tabla <- as.data.frame(res_mat, check.names = FALSE)
  tabla <- cbind(
    Fila = filas_labels,
    tabla
  )

  list(
    id = tabla_obj$id,
    titulo = tabla_obj$titulo,
    tipo = "tabla_conceptual",
    tabla = tabla
  )
}

# =============================================================================
# Cruces
# =============================================================================

.pn_contar_por_opcion <- function(data, var, codes, tp, mask, weight_col = NULL) {
  w <- .pn_get_pesos(data, peso_var = weight_col)

  if (tp == "so") {
    v_codes <- as.character(data[[var]])
    elig <- mask & !is.na(v_codes) & nzchar(v_codes) & v_codes != "NA"
    return(vapply(seq_along(codes), function(j) {
      sum(w[elig & v_codes == codes[j]], na.rm = TRUE)
    }, numeric(1)))
  }

  if (tp == "sm") {
    colc <- .pn_col_sm_compact(data, var)

    if (!is.na(colc)) {
      long <- .pn_sm_compact_to_long(data[[colc]], id = seq_len(nrow(data)), w = w)
      if (!nrow(long)) return(rep(0, length(codes)))
      ids_mask <- which(mask)
      long <- long[long$id %in% ids_mask & long$valor %in% codes, , drop = FALSE]
      return(vapply(seq_along(codes), function(j) {
        code_j <- codes[j]
        ids_j  <- unique(long$id[long$valor == code_j])
        sum(w[ids_j], na.rm = TRUE)
      }, numeric(1)))
    }

    subs <- grep(paste0("^", gsub("([\\W])", "\\\\\\1", var), "[/\\.]"), names(data), value = TRUE)
    if (!length(subs)) return(rep(0, length(codes)))
    codes_dummy <- sub(paste0("^", var, "[/\\.]"), "", subs)

    return(vapply(seq_along(codes), function(j) {
      code_j <- codes[j]
      cols_j <- subs[codes_dummy == code_j]
      if (!length(cols_j)) return(0)

      mat <- sapply(cols_j, function(col) {
        v <- suppressWarnings(as.numeric(as.character(data[[col]])))
        v == 1
      })
      if (!is.matrix(mat)) mat <- matrix(mat, ncol = 1)
      elig_ids <- which(mask & rowSums(mat, na.rm = TRUE) > 0)
      sum(w[elig_ids], na.rm = TRUE)
    }, numeric(1)))
  }

  rep(0, length(codes))
}

.pn_denominador_validos <- function(data, var, codes, tp, mask, weight_col = NULL) {
  w <- .pn_get_pesos(data, peso_var = weight_col)

  if (tp == "so") {
    v_codes <- as.character(data[[var]])
    elig <- mask & !is.na(v_codes) & nzchar(v_codes) & v_codes != "NA" & v_codes %in% codes
    return(sum(w[elig], na.rm = TRUE))
  }

  if (tp == "sm") {
    colc <- .pn_col_sm_compact(data, var)

    if (!is.na(colc)) {
      long <- .pn_sm_compact_to_long(data[[colc]], id = seq_len(nrow(data)), w = w)
      if (!nrow(long)) return(0)
      ids_mask <- which(mask)
      long <- long[long$id %in% ids_mask & long$valor %in% codes, , drop = FALSE]
      denom_ids <- unique(long$id)
      return(sum(w[denom_ids], na.rm = TRUE))
    }

    subs <- grep(paste0("^", gsub("([\\W])", "\\\\\\1", var), "[/\\.]"), names(data), value = TRUE)
    if (!length(subs)) return(0)
    codes_dummy <- sub(paste0("^", var, "[/\\.]"), "", subs)
    subs_keep <- subs[codes_dummy %in% codes]
    if (!length(subs_keep)) return(0)

    mat <- sapply(subs_keep, function(col) {
      v <- suppressWarnings(as.numeric(as.character(data[[col]])))
      v == 1
    })
    if (!is.matrix(mat)) mat <- matrix(mat, ncol = 1)
    elig_ids <- which(mask & rowSums(mat, na.rm = TRUE) > 0)
    return(sum(w[elig_ids], na.rm = TRUE))
  }

  0
}

.pn_comparar_columnas_sig <- function(n_mat, N_vec, alpha = 0.05) {
  K <- ncol(n_mat)
  R <- nrow(n_mat)

  letras <- matrix("", nrow = R, ncol = K, dimnames = dimnames(n_mat))
  sig    <- matrix(FALSE, nrow = R, ncol = K, dimnames = dimnames(n_mat))

  for (i in seq_len(R)) {
    n <- n_mat[i, ]
    N <- N_vec
    p <- ifelse(N > 0, n / N, NA_real_)
    lock <- is.na(p) | N == 0 | p <= 0 | p >= 1
    idx <- which(!lock)

    if (length(idx) >= 2) {
      pairs <- utils::combn(idx, 2, simplify = TRUE)
      pvals <- apply(pairs, 2, function(ab) {
        a <- ab[1]; b <- ab[2]
        pa <- p[a]; pb <- p[b]
        na <- N[a]; nb <- N[b]
        if (any(is.na(c(pa, pb, na, nb))) || any(c(na, nb) == 0)) return(NA_real_)
        ppool <- (n[a] + n[b]) / (na + nb)
        se <- sqrt(ppool * (1 - ppool) * (1 / na + 1 / nb))
        if (!is.finite(se) || se <= 0) return(NA_real_)
        z <- (pa - pb) / se
        2 * stats::pnorm(-abs(z))
      })
      padj <- stats::p.adjust(pvals, method = "bonferroni")
      for (k in seq_along(padj)) {
        if (is.na(padj[k]) || padj[k] >= alpha) next
        a <- pairs[1, k]; b <- pairs[2, k]
        if (p[a] > p[b]) {
          letras[i, a] <- paste(letras[i, a], LETTERS[b])
          sig[i, a]    <- TRUE
        } else if (p[b] > p[a]) {
          letras[i, b] <- paste(letras[i, b], LETTERS[a])
          sig[i, b]    <- TRUE
        }
      }
    }
    letras[i, lock] <- ifelse(nzchar(letras[i, lock]), letras[i, lock], ".a")
  }

  list(letras = letras, sig = sig)
}

.pn_nN_para_sig_simple <- function(data, var, opciones_labels, codes_row, estratos,
                                   var_estrato, tp, weight_col = NULL) {
  v_estrato <- as.character(data[[var_estrato]])

  n_mat <- matrix(
    0,
    nrow = length(opciones_labels),
    ncol = length(estratos),
    dimnames = list(opciones_labels, estratos)
  )
  N_vec <- numeric(length(estratos))
  names(N_vec) <- estratos

  for (j in seq_along(estratos)) {
    catj <- estratos[j]
    mask_j <- !is.na(v_estrato) & v_estrato == catj

    N_vec[j] <- .pn_denominador_validos(
      data = data, var = var, codes = codes_row, tp = tp,
      mask = mask_j, weight_col = weight_col
    )

    if (N_vec[j] == 0) next

    n_vec <- .pn_contar_por_opcion(
      data = data, var = var, codes = codes_row, tp = tp,
      mask = mask_j, weight_col = weight_col
    )

    n_mat[, j] <- n_vec
  }

  list(n_mat = n_mat, N_vec = N_vec)
}

pn_calcular_cruce_indicador <- function(data, cruce_obj, instrumento = NULL, peso = NULL, sm_vars_force = NULL) {
  survey <- .pn_get_survey(instrumento)
  orders_list <- .pn_get_orders_list(instrumento)

  df <- data
  if (!is.null(cruce_obj$filtro) && nzchar(cruce_obj$filtro)) {
    mask <- .pn_eval_condicion_fila(df, cruce_obj$filtro)
    mask[is.na(mask)] <- FALSE
    df <- df[mask, , drop = FALSE]
  }

  if (!nrow(df)) {
    return(list(
      id = cruce_obj$id,
      titulo = cruce_obj$titulo,
      tipo = "cruce_indicador",
      variable = cruce_obj$variable,
      tablas = list(),
      tabla = data.frame()
    ))
  }

  var <- cruce_obj$variable

  if (cruce_obj$tipo_tabla == "resumen_numerico") {
    x <- suppressWarnings(as.numeric(df[[var]]))
    w <- .pn_get_pesos(df, peso_var = peso)

    tabs <- lapply(cruce_obj$cruces, function(s) {
      if (!(s %in% names(df))) return(NULL)
      cats <- .pn_get_categorias(s, df, survey = survey, orders_list = orders_list)
      if (!length(cats$codes)) return(NULL)

      v_estr <- as.character(df[[s]])
      usa_codes <- any(v_estr %in% cats$codes)
      usa_labels <- any(v_estr %in% cats$labels)
      keys_vec <- if (usa_codes || !usa_labels) cats$codes else cats$labels

      filas <- lapply(seq_along(keys_vec), function(j) {
        mask_j <- !is.na(v_estr) & v_estr == keys_vec[j]
        rs <- .pn_resumen_numerico_w_mask(x, w, mask_j)
        tibble::tibble(
          Cruce = pn_obtener_label_var(s, instrumento = instrumento, data = df),
          Grupo = cats$labels[j],
          N = rs["N"],
          Media = rs["Media"],
          SD = rs["SD"],
          Min = rs["Min"],
          P25 = rs["P25"],
          Mediana = rs["Mediana"],
          P75 = rs["P75"],
          Max = rs["Max"]
        )
      })

      dplyr::bind_rows(filas)
    })

    tabs <- tabs[!vapply(tabs, is.null, logical(1))]
    tab <- if (length(tabs)) dplyr::bind_rows(tabs) else data.frame()

    return(list(
      id = cruce_obj$id,
      titulo = cruce_obj$titulo,
      tipo = "cruce_indicador",
      subtipo = "resumen_numerico",
      variable = var,
      tabla = tab,
      tablas = tabs,
      tabla_significancia = NULL
    ))
  }

  tp <- .pn_tipo_pregunta(var, survey = survey, sm_vars_force = sm_vars_force, data = df)
  cats_var <- .pn_get_categorias(
    var = var,
    data = df,
    survey = survey,
    orders_list = orders_list,
    opciones_excluir = cruce_obj$opciones_excluir
  )

  opciones <- cats_var$labels
  codes_row <- cats_var$codes

  op_chr <- trimws(tolower(as.character(opciones)))
  cd_chr <- trimws(tolower(as.character(codes_row)))
  drop_total <- (op_chr == "total") | (cd_chr == "total") | is.na(op_chr) | (op_chr == "")
  if (any(drop_total)) {
    opciones  <- opciones[!drop_total]
    codes_row <- codes_row[!drop_total]
  }

  keep <- !duplicated(trimws(tolower(as.character(opciones))))
  opciones  <- opciones[keep]
  codes_row <- codes_row[keep]

  if (!is.null(cruce_obj$codigos_solo_si_presentes) && length(codes_row)) {
    cod_cond <- as.character(cruce_obj$codigos_solo_si_presentes)
    n_total_all <- .pn_contar_por_opcion(
      data = df, var = var, codes = codes_row, tp = tp,
      mask = rep(TRUE, nrow(df)), weight_col = peso
    )
    to_drop <- codes_row %in% cod_cond & n_total_all == 0
    if (any(to_drop)) {
      codes_row <- codes_row[!to_drop]
      opciones  <- opciones[!to_drop]
    }
  }

  cuerpo <- tibble::tibble(Opciones = opciones)
  denom_map <- list()
  estratos_totales <- list()

  N_total <- .pn_denominador_validos(
    data = df, var = var, codes = codes_row, tp = tp,
    mask = rep(TRUE, nrow(df)), weight_col = peso
  )
  n_total <- .pn_contar_por_opcion(
    data = df, var = var, codes = codes_row, tp = tp,
    mask = rep(TRUE, nrow(df)), weight_col = peso
  )
  pct_total <- if (N_total > 0) n_total / N_total else rep(NA_real_, length(n_total))

  cuerpo <- dplyr::bind_cols(
    cuerpo,
    tibble::tibble(Total__n = as.numeric(n_total), Total__pct = as.numeric(pct_total))
  )
  denom_map[["Total__n"]] <- N_total

  for (s in cruce_obj$cruces) {
    if (!(s %in% names(df)) || identical(s, var)) next

    cats_s <- .pn_get_categorias(s, df, survey = survey, orders_list = orders_list, opciones_excluir = NULL)
    estr_codes  <- cats_s$codes
    estr_labels <- cats_s$labels
    if (!length(estr_codes)) next

    estratos_totales[[s]] <- list(codes = estr_codes, labels = estr_labels)

    v_estr <- as.character(df[[s]])
    usa_codes  <- any(v_estr %in% estr_codes)
    usa_labels <- any(v_estr %in% estr_labels)
    keys_vec <- if (usa_codes || !usa_labels) estr_codes else estr_labels

    bloques <- lapply(seq_along(keys_vec), function(j) {
      key_j <- keys_vec[j]
      mask_s <- !is.na(v_estr) & v_estr == key_j

      n_vec <- .pn_contar_por_opcion(
        data = df, var = var, codes = codes_row, tp = tp,
        mask = mask_s, weight_col = peso
      )

      N <- .pn_denominador_validos(
        data = df, var = var, codes = codes_row, tp = tp,
        mask = mask_s, weight_col = peso
      )

      pct <- if (N > 0) n_vec / N else rep(NA_real_, length(n_vec))
      nm_n   <- paste0(s, "__", make.names(estr_labels[j]), "__n")
      nm_pct <- paste0(s, "__", make.names(estr_labels[j]), "__pct")

      dfb <- tibble::tibble(
        !!nm_n   := as.numeric(n_vec),
        !!nm_pct := as.numeric(pct)
      )
      list(df = dfb, N = N)
    })

    cols_df <- dplyr::bind_cols(lapply(bloques, `[[`, "df"))
    idx_n_cols <- grep("__n$", names(cols_df))
    Ns <- vapply(bloques, `[[`, numeric(1), "N")

    if (length(idx_n_cols) == length(Ns) && length(Ns) > 0) {
      for (k in seq_along(idx_n_cols)) {
        denom_map[[names(cols_df)[idx_n_cols[k]]]] <- Ns[k]
      }
    }

    cuerpo <- dplyr::bind_cols(cuerpo, cols_df)
  }

  total_row <- as.list(rep(NA, ncol(cuerpo)))
  names(total_row) <- names(cuerpo)
  total_row[["Opciones"]] <- "Total"

  n_cols   <- grep("__n$", names(cuerpo))
  pct_cols <- grep("__pct$", names(cuerpo))

  for (j in n_cols) {
    nm <- names(cuerpo)[j]
    Nj <- denom_map[[nm]]
    total_row[[j]] <- if (is.null(Nj)) NA_real_ else round(as.numeric(Nj), 0)
  }
  for (j in pct_cols) {
    n_partner <- sub("__pct$", "__n", names(cuerpo)[j])
    Nj <- suppressWarnings(as.numeric(total_row[[n_partner]]))
    total_row[[j]] <- if (!is.na(Nj) && Nj > 0) 1.0 else NA_real_
  }

  cuerpo <- dplyr::bind_rows(cuerpo, tibble::as_tibble(total_row))

  tabla_sig <- NULL
  if (isTRUE(cruce_obj$mostrar_significancia) && length(estratos_totales)) {
    sig_out <- list()

    for (s in cruce_obj$cruces) {
      info_s <- estratos_totales[[s]]
      if (is.null(info_s)) next

      nn <- .pn_nN_para_sig_simple(
        data = df,
        var = var,
        opciones_labels = opciones,
        codes_row = codes_row,
        estratos = info_s$codes,
        var_estrato = s,
        tp = tp,
        weight_col = peso
      )

      cmp <- .pn_comparar_columnas_sig(nn$n_mat, nn$N_vec, alpha = cruce_obj$alpha)

      tmp <- as.data.frame(cmp$letras, stringsAsFactors = FALSE, check.names = FALSE)
      tmp <- cbind(Opciones = rownames(tmp), tmp, row.names = NULL)
      sig_out[[s]] <- tmp
    }

    tabla_sig <- sig_out
  }

  list(
    id = cruce_obj$id,
    titulo = cruce_obj$titulo,
    tipo = "cruce_indicador",
    subtipo = "frecuencia",
    variable = var,
    tabla = cuerpo,
    estratos = estratos_totales,
    tabla_significancia = tabla_sig
  )
}

# =============================================================================
# Ejecución integral
# =============================================================================

#' Ejecutar un plan analítico completo
#'
#' Aplica normalización, calcula dimensiones, tablas y cruces definidos en un
#' objeto `pn_plan_analitico`.
#'
#' @param data Data frame de entrada.
#' @param plan Objeto creado con `pn_plan_analitico()`.
#' @param prefijo_normalizadas Prefijo para variables normalizadas.
#' @param sm_vars_force Variables a forzar como `select_multiple`.
#'
#' @return Objeto de clase `pn_resultado_plan_analitico`.
#' @export
pn_ejecutar_plan_analitico <- function(data, plan, prefijo_normalizadas = "norm_",
                                       sm_vars_force = NULL) {
  .pn_assert_df(data, "data")
  if (!inherits(plan, "plan_analitico")) {
    stop("`plan` debe crearse con `pn_plan_analitico()`.", call. = FALSE)
  }

  instrumento <- plan$instrumento %||% NULL
  peso_var <- plan$peso %||% NULL

  norm_out <- pn_aplicar_normalizacion(
    data = data,
    reglas = plan$normalizacion,
    prefijo = prefijo_normalizadas
  )
  data1 <- norm_out$data

  dim_out <- pn_calcular_dimensiones(
    data = data1,
    dimensiones = plan$dimensiones,
    prefijo_norm = prefijo_normalizadas
  )
  data2 <- dim_out$data

  tablas_res <- lapply(plan$tablas, function(tb) {
    if (inherits(tb, "tabla_simple")) {
      pn_calcular_tabla_simple(data2, tb, instrumento = instrumento, peso = peso_var, sm_vars_force = sm_vars_force)
    } else if (inherits(tb, "tabla_condiciones")) {
      pn_calcular_tabla_condiciones(data2, tb, peso = peso_var)
    } else if (inherits(tb, "tabla_conceptual")) {
      pn_calcular_tabla_conceptual(data2, tb, peso = peso_var)
    } else {
      NULL
    }
  })
  if (length(tablas_res)) names(tablas_res) <- vapply(plan$tablas, function(x) x$id, character(1))

  cruces_res <- lapply(plan$cruces, function(cr) {
    pn_calcular_cruce_indicador(data2, cr, instrumento = instrumento, peso = peso_var, sm_vars_force = sm_vars_force)
  })
  if (length(cruces_res)) names(cruces_res) <- vapply(plan$cruces, function(x) x$id, character(1))

  out <- list(
    plan = plan,
    data_original = data,
    data_trabajo = data2,
    normalizacion = norm_out$meta,
    dimensiones = dim_out$meta,
    tablas = tablas_res,
    cruces = cruces_res
  )
  class(out) <- c("pn_resultado_plan_analitico", "resultado_plan_analitico", "list")
  out
}

# =============================================================================
# Estilos Excel
# =============================================================================

#' Crear estilos Excel para exportación de planes
#'
#' @return Lista de estilos `openxlsx::createStyle()` compatible con
#'   `pn_exportar_plan_analitico_excel()`.
#' @export
pn_mk_styles_plan <- function() {
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("Se requiere 'openxlsx' para exportar a Excel.", call. = FALSE)
  }

  list(
    titulo = openxlsx::createStyle(
      fontSize = 11,
      textDecoration = "italic",
      halign = "left",
      valign = "center",
      wrapText = TRUE
    ),
    subtitulo = openxlsx::createStyle(
      fontSize = 12,
      textDecoration = "bold",
      halign = "left",
      valign = "center",
      wrapText = TRUE
    ),
    notas = openxlsx::createStyle(
      fontSize = 9,
      textDecoration = "italic",
      halign = "left",
      valign = "top",
      wrapText = TRUE
    ),
    header = openxlsx::createStyle(
      fontSize = 10,
      textDecoration = "bold",
      border = c("top", "bottom"),
      borderStyle = "thin",
      halign = "center",
      valign = "center",
      wrapText = TRUE
    ),
    cuerpo = openxlsx::createStyle(
      fontSize = 10,
      halign = "center",
      valign = "center",
      wrapText = TRUE
    ),
    body_txt = openxlsx::createStyle(
      fontSize = 10,
      halign = "left",
      valign = "center",
      wrapText = TRUE
    ),
    body_int = openxlsx::createStyle(
      fontSize = 10,
      numFmt = "#,##0",
      halign = "right",
      valign = "center",
      wrapText = TRUE
    ),
    body_num = openxlsx::createStyle(
      fontSize = 10,
      numFmt = "#,##0.0",
      halign = "right",
      valign = "center",
      wrapText = TRUE
    ),
    body_pct = openxlsx::createStyle(
      fontSize = 10,
      numFmt = "0.0%",
      halign = "right",
      valign = "center",
      wrapText = TRUE
    ),
    fuente = openxlsx::createStyle(
      fontSize = 9,
      halign = "left",
      valign = "center",
      fontColour = "#808080",
      wrapText = TRUE
    ),
    table_end = openxlsx::createStyle(
      border = c("top"),
      borderStyle = "thin",
      borderColour = "#000000"
    )
  )
}

.pn_es_tipo_pct <- function(nm, tipo = NULL) {
  grepl("%", nm, fixed = TRUE) ||
    grepl("pct", nm, ignore.case = TRUE) ||
    identical(tipo, "proporcion")
}

# =============================================================================
# Escritura Excel especializada
# =============================================================================

.pn_escribir_titulo_bloque <- function(wb, sheet, titulo, row_start, col_start, ncol_block, estilos) {
  openxlsx::writeData(wb, sheet, titulo, startRow = row_start, startCol = col_start)
  if (ncol_block > 1) {
    openxlsx::mergeCells(wb, sheet, rows = row_start, cols = col_start:(col_start + ncol_block - 1L))
  }
  openxlsx::addStyle(wb, sheet, estilos$titulo, rows = row_start, cols = col_start, gridExpand = TRUE, stack = TRUE)
  invisible(row_start + 1L)
}

.pn_escribir_tabla_excel_simple <- function(
    wb, sheet, tabla, titulo, row_start = 1L, col_start = 1L,
    estilos, fuente = NULL, tipo = NULL
) {
  n_filas_tabla <- nrow(tabla)
  n_cols_tabla  <- ncol(tabla)
  r <- row_start

  last_col <- if (n_cols_tabla > 0) col_start + n_cols_tabla - 1L else col_start

  r <- .pn_escribir_titulo_bloque(wb, sheet, titulo, r, col_start, max(1L, n_cols_tabla), estilos)

  if (n_filas_tabla <= 0 || n_cols_tabla <= 0) {
    row_after_body <- r + 1L
  } else {
    header <- names(tabla)
    openxlsx::writeData(
      wb, sheet,
      x = matrix(header, nrow = 1),
      startRow = r,
      startCol = col_start,
      colNames = FALSE
    )
    openxlsx::addStyle(
      wb, sheet, estilos$header,
      rows = r, cols = seq(col_start, length.out = n_cols_tabla),
      gridExpand = TRUE, stack = TRUE
    )
    r <- r + 1L

    openxlsx::writeData(
      wb, sheet,
      x = tabla,
      startRow = r,
      startCol = col_start,
      colNames = FALSE
    )

    body_row_ini <- r
    body_row_fin <- r + n_filas_tabla - 1L

    openxlsx::addStyle(
      wb, sheet, estilos$body_txt,
      rows = body_row_ini:body_row_fin, cols = col_start,
      gridExpand = TRUE, stack = TRUE
    )

    if (n_cols_tabla > 1) {
      for (j in 2:n_cols_tabla) {
        col_abs  <- col_start + j - 1L
        nm       <- names(tabla)[j]
        col_data <- suppressWarnings(as.numeric(tabla[[j]]))

        es_pct <- .pn_es_tipo_pct(nm, tipo = tipo)
        if (es_pct) {
          estilo_col <- estilos$body_pct
        } else {
          is_entero <- all(is.na(col_data) | abs(col_data - round(col_data)) < 1e-8)
          estilo_col <- if (is_entero) estilos$body_int else estilos$body_num
        }

        openxlsx::addStyle(
          wb, sheet, estilo_col,
          rows = body_row_ini:body_row_fin,
          cols = col_abs,
          gridExpand = TRUE, stack = TRUE
        )
      }
    }

    row_after_body <- body_row_fin + 1L
  }

  if (!is.null(fuente)) {
    row_fuente <- row_after_body
    openxlsx::writeData(wb, sheet, fuente, startRow = row_fuente, startCol = col_start, colNames = FALSE)
    if (last_col > col_start) {
      openxlsx::mergeCells(wb, sheet, rows = row_fuente, cols = col_start:last_col)
    }
    openxlsx::addStyle(wb, sheet, estilos$fuente, rows = row_fuente, cols = col_start, gridExpand = TRUE, stack = TRUE)
    openxlsx::addStyle(wb, sheet, estilos$table_end, rows = row_fuente, cols = col_start:last_col, gridExpand = TRUE, stack = TRUE)
    next_row <- row_fuente + 2L
  } else {
    row_cierre <- row_after_body
    openxlsx::addStyle(wb, sheet, estilos$table_end, rows = row_cierre, cols = col_start:last_col, gridExpand = TRUE, stack = TRUE)
    next_row <- row_cierre + 2L
  }

  invisible(list(next_row = next_row))
}

.pn_escribir_tabla_excel_condiciones <- function(
    wb, sheet, tabla, titulo, row_start = 1L, col_start = 1L,
    estilos, fuente = NULL
) {
  .pn_escribir_tabla_excel_simple(
    wb = wb, sheet = sheet, tabla = tabla, titulo = titulo,
    row_start = row_start, col_start = col_start,
    estilos = estilos, fuente = fuente, tipo = NULL
  )
}

.pn_escribir_tabla_excel_conceptual <- function(
    wb, sheet, tabla, titulo, row_start = 1L, col_start = 1L,
    estilos, fuente = NULL
) {
  .pn_escribir_tabla_excel_simple(
    wb = wb, sheet = sheet, tabla = tabla, titulo = titulo,
    row_start = row_start, col_start = col_start,
    estilos = estilos, fuente = fuente, tipo = NULL
  )
}

.pn_escribir_tabla_excel_cruce <- function(
    wb, sheet, tabla, titulo, row_start = 1L, col_start = 1L,
    estilos, fuente = NULL
) {
  .pn_escribir_tabla_excel_simple(
    wb = wb, sheet = sheet, tabla = tabla, titulo = titulo,
    row_start = row_start, col_start = col_start,
    estilos = estilos, fuente = fuente, tipo = NULL
  )
}

.pn_escribir_tabla_excel_significancia <- function(
    wb, sheet, tablas_sig, titulo, row_start = 1L, col_start = 1L,
    estilos, fuente = NULL
) {
  r <- row_start
  r <- .pn_escribir_titulo_bloque(wb, sheet, titulo, r, col_start, 6L, estilos)

  for (nm in names(tablas_sig)) {
    tb <- tablas_sig[[nm]]
    openxlsx::writeData(wb, sheet, nm, startRow = r, startCol = col_start)
    openxlsx::addStyle(wb, sheet, estilos$subtitulo, rows = r, cols = col_start, gridExpand = TRUE, stack = TRUE)
    r <- r + 1L

    tmp <- .pn_escribir_tabla_excel_simple(
      wb = wb, sheet = sheet, tabla = tb, titulo = "",
      row_start = r, col_start = col_start,
      estilos = estilos, fuente = NULL
    )
    r <- tmp$next_row
  }

  if (!is.null(fuente)) {
    openxlsx::writeData(wb, sheet, fuente, startRow = r, startCol = col_start, colNames = FALSE)
    openxlsx::addStyle(wb, sheet, estilos$fuente, rows = r, cols = col_start, gridExpand = TRUE, stack = TRUE)
    r <- r + 2L
  }

  invisible(list(next_row = r))
}

# =============================================================================
# Exportación principal
# =============================================================================

#' Exportar resultados del plan analítico a Excel
#'
#' @param resultado Objeto resultado de `pn_ejecutar_plan_analitico()`.
#' @param path_xlsx Ruta del archivo de salida.
#' @param hoja_tablas Nombre de hoja para tablas.
#' @param hoja_cruces Nombre de hoja para cruces.
#' @param hoja_sig Nombre de hoja para significancia.
#' @param estilos Lista de estilos `openxlsx`; si `NULL`, usa `pn_mk_styles_plan()`.
#' @param fuente Texto de fuente al pie de cada bloque.
#'
#' @return Ruta normalizada del archivo generado (invisible).
#' @export
pn_exportar_plan_analitico_excel <- function(
    resultado,
    path_xlsx = "plan_analitico.xlsx",
    hoja_tablas = "Tablas",
    hoja_cruces = "Cruces",
    hoja_sig = "Significancia",
    estilos = NULL,
    fuente = NULL
) {
  if (!inherits(resultado, "resultado_plan_analitico")) {
    stop("`resultado` debe provenir de `pn_ejecutar_plan_analitico()`.", call. = FALSE)
  }

  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("Se requiere el paquete 'openxlsx'.", call. = FALSE)
  }

  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, hoja_tablas)
  openxlsx::addWorksheet(wb, hoja_cruces)
  openxlsx::addWorksheet(wb, hoja_sig)

  if (is.null(estilos)) estilos <- pn_mk_styles_plan()

  row_tab <- 1L
  if (length(resultado$tablas)) {
    for (nm in names(resultado$tablas)) {
      tb <- resultado$tablas[[nm]]
      if (inherits(tb$tabla, "data.frame")) {
        escrito <- if (identical(tb$tipo, "tabla_condiciones")) {
          .pn_escribir_tabla_excel_condiciones(
            wb, hoja_tablas, tb$tabla, tb$titulo %||% nm,
            row_start = row_tab, col_start = 1L, estilos = estilos, fuente = fuente
          )
        } else if (identical(tb$tipo, "tabla_conceptual")) {
          .pn_escribir_tabla_excel_conceptual(
            wb, hoja_tablas, tb$tabla, tb$titulo %||% nm,
            row_start = row_tab, col_start = 1L, estilos = estilos, fuente = fuente
          )
        } else {
          .pn_escribir_tabla_excel_simple(
            wb, hoja_tablas, tb$tabla, tb$titulo %||% nm,
            row_start = row_tab, col_start = 1L, estilos = estilos, fuente = fuente,
            tipo = tb$tipo %||% NULL
          )
        }
        row_tab <- escrito$next_row
      }
    }
  }

  row_cru <- 1L
  row_sig <- 1L
  if (length(resultado$cruces)) {
    for (nm in names(resultado$cruces)) {
      tb <- resultado$cruces[[nm]]

      if (inherits(tb$tabla, "data.frame")) {
        escrito <- .pn_escribir_tabla_excel_cruce(
          wb, hoja_cruces, tb$tabla, tb$titulo %||% nm,
          row_start = row_cru, col_start = 1L, estilos = estilos, fuente = fuente
        )
        row_cru <- escrito$next_row
      }

      if (!is.null(tb$tabla_significancia) && length(tb$tabla_significancia)) {
        escrito_sig <- .pn_escribir_tabla_excel_significancia(
          wb, hoja_sig, tb$tabla_significancia,
          titulo = paste0(tb$titulo %||% nm, " — Significancia"),
          row_start = row_sig, col_start = 1L,
          estilos = estilos, fuente = fuente
        )
        row_sig <- escrito_sig$next_row
      }
    }
  }

  openxlsx::saveWorkbook(wb, path_xlsx, overwrite = TRUE)
  invisible(normalizePath(path_xlsx, winslash = "/"))
}

# =============================================================================
# Resumen legible
# =============================================================================

#' Imprimir resumen de un plan analítico
#'
#' @param plan Objeto de clase `plan_analitico`.
#'
#' @return El objeto `plan` de entrada, invisiblemente.
#' @export
pn_imprimir_plan_analitico <- function(plan) {
  if (!inherits(plan, "plan_analitico")) {
    stop("`plan` debe ser un objeto `plan_analitico`.", call. = FALSE)
  }

  cat("\n")
  cat("============================================================\n")
  cat("PLAN ANALÍTICO\n")
  cat("============================================================\n")
  cat("Nombre: ", plan$nombre, "\n", sep = "")
  cat("Peso:   ", plan$peso %||% "(sin peso)", "\n", sep = "")
  cat("------------------------------------------------------------\n")

  cat("Normalización: ", length(plan$normalizacion), "\n", sep = "")
  if (length(plan$normalizacion)) {
    for (i in seq_along(plan$normalizacion)) {
      rg <- plan$normalizacion[[i]]
      cat("  - ", rg$id, " [", rg$metodo, "]", sep = "")
      if (isTRUE(rg$aplicar_a_todas)) {
        cat(" -> todas")
        if (isTRUE(rg$solo_numericas)) cat(" (solo numericables)")
      } else {
        cat(" -> ", paste(rg$variables, collapse = ", "), sep = "")
      }
      cat("\n")
    }
  }

  cat("Dimensiones/índices: ", length(plan$dimensiones), "\n", sep = "")
  if (length(plan$dimensiones)) {
    for (i in seq_along(plan$dimensiones)) {
      dm <- plan$dimensiones[[i]]
      vars <- vapply(dm$items, function(it) it$variable, character(1))
      cat("  - ", dm$id, " [", dm$agregacion, "] -> ", paste(vars, collapse = ", "), "\n", sep = "")
    }
  }

  cat("Tablas: ", length(plan$tablas), "\n", sep = "")
  if (length(plan$tablas)) {
    for (i in seq_along(plan$tablas)) {
      tb <- plan$tablas[[i]]
      cls <- class(tb)[1]
      cat("  - ", tb$id, " <", cls, ">\n", sep = "")
    }
  }

  cat("Cruces: ", length(plan$cruces), "\n", sep = "")
  if (length(plan$cruces)) {
    for (i in seq_along(plan$cruces)) {
      cr <- plan$cruces[[i]]
      cat("  - ", cr$id, " -> ", cr$variable, " x ", paste(cr$cruces, collapse = ", "), "\n", sep = "")
    }
  }

  cat("============================================================\n")
  invisible(plan)
}

# =============================================================================
# Helpers de construcción
# =============================================================================

#' Obtener todas las variables numéricas de un dataset
#'
#' @param data Data frame de entrada.
#' @param excluir Variables a excluir.
#'
#' @return Vector de nombres de variables numéricas.
#' @export
pn_usar_todas_las_numericas <- function(data, excluir = NULL) {
  .pn_assert_df(data, "data")
  excl <- .pn_compact_chr(excluir)
  vars <- names(data)[vapply(data, is.numeric, logical(1))]
  setdiff(vars, excl)
}

#' Crear items de indicador desde un vector de variables
#'
#' @param variables Nombres de variables.
#' @param peso Peso común para todos los items.
#' @param usar_normalizada Si `TRUE`, cada item usará su variable normalizada.
#'
#' @return Lista de objetos `pn_item_indicador`.
#' @export
pn_items_desde_variables <- function(variables, peso = 1, usar_normalizada = TRUE) {
  variables <- .pn_compact_chr(variables)
  lapply(variables, function(v) {
    pn_item_indicador(variable = v, peso = peso, usar_normalizada = usar_normalizada)
  })
}

#' Crear dimensión de forma rápida desde variables
#'
#' @param id Identificador de la dimensión.
#' @param variables Nombres de variables.
#' @param titulo Título legible.
#' @param agregacion Método de agregación.
#' @param usar_normalizada Si `TRUE`, usa variables normalizadas.
#' @param peso_item Peso común de cada item.
#' @param minimo_items Mínimo de items válidos por fila.
#' @param estandarizar_resultado Si `TRUE`, reescala al rango indicado.
#' @param rango_resultado Rango de salida cuando `estandarizar_resultado = TRUE`.
#'
#' @return Objeto de clase `pn_dimension_indicador`.
#' @export
pn_dimension_desde_variables <- function(
    id,
    variables,
    titulo = NULL,
    agregacion = "promedio",
    usar_normalizada = TRUE,
    peso_item = 1,
    minimo_items = 1,
    estandarizar_resultado = FALSE,
    rango_resultado = c(0, 100)
) {
  its <- pn_items_desde_variables(
    variables = variables,
    peso = peso_item,
    usar_normalizada = usar_normalizada
  )

  pn_dimension_indicador(
    id = id,
    titulo = titulo %||% id,
    items = its,
    agregacion = agregacion,
    minimo_items = minimo_items,
    estandarizar_resultado = estandarizar_resultado,
    rango_resultado = rango_resultado
  )
}

