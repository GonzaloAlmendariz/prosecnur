# =============================================================================
# Estilos para cruces
# =============================================================================

mk_styles_cruces <- function() {
  list(
    sec_title = openxlsx::createStyle(
      fontSize       = 18,
      textDecoration = "bold",
      halign         = "center",
      valign         = "center",
      wrapText       = TRUE,
      fontColour     = "#000000",
      fgFill         = "#FFFFFF",
      fontName       = "Arial"
    ),
    q_title = openxlsx::createStyle(
      fontSize       = 11,
      textDecoration = "italic",
      halign         = "left",
      valign         = "center",
      wrapText       = TRUE,
      fontColour     = "#000000",
      fontName       = "Arial"
    ),
    header = openxlsx::createStyle(
      fontSize       = 10,
      textDecoration = "bold",
      border         = c("top", "bottom"),
      borderStyle    = "thin",
      borderColour   = "#000000",
      halign         = "center",
      valign         = "center",
      wrapText       = TRUE,
      fgFill         = "#FFFFFF",
      fontName       = "Arial"
    ),
    header_A = openxlsx::createStyle(
      fontSize       = 10,
      textDecoration = "bold",
      halign         = "left",
      valign         = "center",
      wrapText       = TRUE,
      fgFill         = "#FFFFFF",
      fontName       = "Arial"
    ),
    body_txt = openxlsx::createStyle(
      fontSize = 10,
      halign   = "left",
      valign   = "center",
      wrapText = TRUE,
      fgFill   = "#FFFFFF",
      fontName = "Arial"
    ),
    body_int = openxlsx::createStyle(
      fontSize = 10,
      numFmt   = "#,##0",
      halign   = "right",
      valign   = "center",
      fontName = "Arial"
    ),
    body_pct = openxlsx::createStyle(
      fontSize = 10,
      numFmt   = "0.0%",
      halign   = "right",
      valign   = "center",
      fontName = "Arial"
    ),
    note = openxlsx::createStyle(
      fontSize       = 9,
      fontColour     = "#666666",
      halign         = "left",
      valign         = "center",
      wrapText       = TRUE,
      textDecoration = "italic",
      fontName       = "Arial"
    ),
    total_bold = openxlsx::createStyle(
      textDecoration = "bold",
      fontName       = "Arial"
    ),
    table_end = openxlsx::createStyle(
      border       = "bottom",
      borderStyle  = "thin",
      borderColour = "#000000"
    ),
    cell = openxlsx::createStyle(
      fontSize = 10,
      halign   = "center",
      valign   = "center",
      fontName = "Arial"
    ),
    fill_sig = openxlsx::createStyle(
      fgFill = "#f1d6b3"
    )
  )
}

# =============================================================================
# Helpers básicos
# =============================================================================

get_pesos <- function(data, weight_col = "peso") {
  if (!is.null(weight_col) && weight_col %in% names(data)) {
    w <- suppressWarnings(as.numeric(data[[weight_col]]))
    w[is.na(w) | !is.finite(w)] <- 0
    return(w)
  }
  rep(1, nrow(data))
}

tipo_pregunta <- function(var, survey = NULL, sm_vars_force = NULL) {
  if (!is.null(sm_vars_force) && var %in% sm_vars_force) return("sm")
  if (!is.null(survey) && any(survey$name == var)) {
    tipos <- unique(na.omit(survey$type[survey$name == var]))
    if (any(grepl("^select_multiple(\\s|$)", tipos))) return("sm")
    if (any(grepl("^select_one(\\s|$)", tipos)))      return("so")
  }
  "so"
}

col_sm_compact <- function(data, var) {
  v_orig <- paste0(var, "_ORIG")
  if (v_orig %in% names(data)) return(v_orig)
  if (var %in% names(data))    return(var)
  NA_character_
}

sm_compact_to_long <- function(x, id, w) {
  tibble::tibble(
    id    = id,
    valor = as.character(x),
    w     = as.numeric(w)
  ) |>
    tidyr::separate_rows(valor, sep = "\\s*;\\s*", convert = FALSE) |>
    dplyr::mutate(valor = trimws(valor)) |>
    dplyr::filter(!is.na(valor) & nzchar(valor) & valor != "NA")
}

label_variable <- function(var, dic_vars = NULL, labels_override = NULL, data = NULL) {
  if (!is.null(labels_override) && var %in% names(labels_override)) {
    return(as.character(labels_override[[var]]))
  }
  if (!is.null(data) && var %in% names(data)) {
    vlab <- attr(data[[var]], "label", exact = TRUE)
    if (!is.null(vlab) && nzchar(as.character(vlab))) {
      return(as.character(vlab))
    }
  }
  if (!is.null(dic_vars) && all(c("name", "label") %in% names(dic_vars))) {
    lab <- dic_vars$label[dic_vars$name == var]
    if (length(lab) && !all(is.na(lab))) return(as.character(lab[1]))
  }
  as.character(var)
}

# =============================================================================
# Mapeo códigos/labels usando instrumento (survey + orders_list)
# =============================================================================

get_list_name <- function(var, survey = NULL) {
  if (is.null(survey) ||
      !all(c("name", "list_name") %in% names(survey))) {
    return(NA_character_)
  }
  ln <- unique(na.omit(as.character(survey$list_name[survey$name == var])))
  if (!length(ln)) return(NA_character_)
  ln[1]
}

get_categorias <- function(var,
                           data,
                           survey          = NULL,
                           orders_list     = NULL,
                           opciones_excluir = NULL) {
  x <- data[[var]]
  lab_attr <- attr(x, "labels", exact = TRUE)

  ln <- get_list_name(var, survey)
  codes  <- character(0)
  labels <- character(0)

  # 1) orders_list: primero por variable, luego por list_name
  obj <- NULL
  if (!is.null(orders_list)) {
    if (var %in% names(orders_list)) {
      # mismo esquema que frecuencias: orders_list[[var]]
      obj <- orders_list[[var]]
    } else if (!is.na(ln) && ln %in% names(orders_list)) {
      # fallback: algunas implementaciones guardan por list_name
      obj <- orders_list[[ln]]
    }
  }

  if (!is.null(obj)) {
    codes  <- as.character(obj$names)
    labels <- as.character(obj$labels)

  } else if (!is.null(lab_attr) && length(lab_attr) > 0) {
    # 2) attr(labels) de la data (reporte_data)
    codes  <- names(lab_attr)
    labels <- as.character(unname(lab_attr))

  } else {
    # 3) fallback: categorías en los datos
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

# =============================================================================
# Conteo y denominador para SO y SM
# =============================================================================

contar_por_opcion <- function(data,
                              var,
                              codes,
                              tp,
                              mask,
                              weight_col = "peso") {
  w <- get_pesos(data, weight_col)

  if (tp == "so") {
    v_codes <- as.character(data[[var]])
    elig    <- mask & !is.na(v_codes) & nzchar(v_codes) & v_codes != "NA"
    vapply(seq_along(codes), function(j) {
      sum(w[elig & v_codes == codes[j]], na.rm = TRUE)
    }, numeric(1))
  } else if (tp == "sm") {

    colc <- col_sm_compact(data, var)
    if (!is.na(colc)) {
      long <- sm_compact_to_long(data[[colc]], id = seq_len(nrow(data)), w = w)
      if (!nrow(long)) return(rep(0, length(codes)))
      ids_mask <- which(mask)
      long <- long[long$id %in% ids_mask & long$valor %in% codes, , drop = FALSE]
      vapply(seq_along(codes), function(j) {
        code_j <- codes[j]
        ids_j  <- unique(long$id[long$valor == code_j])
        sum(w[ids_j], na.rm = TRUE)
      }, numeric(1))
    } else {
      subs <- grep(paste0("^", stringr::fixed(var), "/"), names(data), value = TRUE)
      if (!length(subs)) return(rep(0, length(codes)))
      codes_dummy <- sub(paste0("^", var, "/"), "", subs)

      vapply(seq_along(codes), function(j) {
        code_j   <- codes[j]
        cols_j   <- subs[codes_dummy == code_j]
        if (!length(cols_j)) return(0)
        mat <- sapply(cols_j, function(col) {
          v <- suppressWarnings(as.numeric(as.character(data[[col]])))
          v == 1
        })
        if (!is.matrix(mat)) mat <- matrix(mat, ncol = 1)
        elig_ids <- which(mask & rowSums(mat, na.rm = TRUE) > 0)
        sum(w[elig_ids], na.rm = TRUE)
      }, numeric(1))
    }
  } else {
    rep(0, length(codes))
  }
}

denominador_validos <- function(data,
                                var,
                                codes,
                                tp,
                                mask,
                                weight_col = "peso") {
  w <- get_pesos(data, weight_col)

  if (tp == "so") {
    v_codes <- as.character(data[[var]])
    elig <- mask &
      !is.na(v_codes) &
      nzchar(v_codes) &
      v_codes != "NA" &
      v_codes %in% codes
    return(sum(w[elig], na.rm = TRUE))
  }

  if (tp == "sm") {
    colc <- col_sm_compact(data, var)
    if (!is.na(colc)) {
      long <- sm_compact_to_long(data[[colc]], id = seq_len(nrow(data)), w = w)
      if (!nrow(long)) return(0)
      ids_mask <- which(mask)
      long <- long[long$id %in% ids_mask & long$valor %in% codes, , drop = FALSE]
      denom_ids <- unique(long$id)
      return(sum(w[denom_ids], na.rm = TRUE))
    } else {
      subs <- grep(paste0("^", stringr::fixed(var), "/"), names(data), value = TRUE)
      if (!length(subs)) return(0)
      codes_dummy <- sub(paste0("^", var, "/"), "", subs)
      subs_keep   <- subs[codes_dummy %in% codes]
      if (!length(subs_keep)) return(0)
      mat <- sapply(subs_keep, function(col) {
        v <- suppressWarnings(as.numeric(as.character(data[[col]])))
        v == 1
      })
      if (!is.matrix(mat)) mat <- matrix(mat, ncol = 1)
      elig_ids <- which(mask & rowSums(mat, na.rm = TRUE) > 0)
      return(sum(w[elig_ids], na.rm = TRUE))
    }
  }

  0
}

# =============================================================================
# Significancia (z + Bonferroni)
# =============================================================================

comparar_columnas_sig <- function(n_mat, N_vec, alpha = 0.05) {
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
        se <- sqrt(ppool * (1 - ppool) * (1/na + 1/nb))
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

nN_para_sig_simple <- function(data,
                               var,
                               opciones_labels,
                               codes_row,
                               estratos,
                               var_estrato,
                               tp,
                               weight_col = "peso") {

  w <- get_pesos(data, weight_col)
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

    N_vec[j] <- denominador_validos(
      data       = data,
      var        = var,
      codes      = codes_row,
      tp         = tp,
      mask       = mask_j,
      weight_col = weight_col
    )

    if (N_vec[j] == 0) next

    n_vec <- contar_por_opcion(
      data       = data,
      var        = var,
      codes      = codes_row,
      tp         = tp,
      mask       = mask_j,
      weight_col = weight_col
    )

    n_mat[, j] <- n_vec
  }

  list(n_mat = n_mat, N_vec = N_vec)
}

# =============================================================================
# Exportador principal: exportar_cruces_multi
# =============================================================================

exportar_cruces_multi <- function(data,
                                  dic_vars,
                                  SECCIONES,
                                  CRUZAR_CON,
                                  labels_override  = NULL,
                                  path_xlsx        = "cruces_multi.xlsx",
                                  hoja             = "Cruces",
                                  fuente           = "Pulso PUCP",
                                  survey           = NULL,
                                  sm_vars_force    = NULL,
                                  weight_col       = "peso",
                                  orders_list      = NULL,
                                  opciones_excluir = NULL,
                                  show_sig         = TRUE,
                                  alpha            = 0.05) {

  SECCIONES  <- lapply(SECCIONES, function(v) v[v %in% names(data)])
  CRUZAR_CON <- CRUZAR_CON[CRUZAR_CON %in% names(data)]
  stopifnot(length(CRUZAR_CON) > 0)

  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, hoja)
  st <- mk_styles_cruces()

  fila <- 1L

  # título general
  openxlsx::writeData(wb, hoja, "CRUCES", startRow = fila, startCol = 1)
  openxlsx::addStyle(wb, hoja, st$sec_title, rows = fila, cols = 1, gridExpand = TRUE)
  openxlsx::mergeCells(wb, hoja, rows = fila, cols = 1:6)
  fila <- fila + 2

  # helper para merges de encabezado
  escribir_encabezado <- function(h1, h2, h3, row0, col0 = 1) {
    ncols <- length(h3)
    if (ncols == 0) return(invisible(NULL))

    openxlsx::writeData(wb, hoja, t(h1), startRow = row0,     startCol = col0, colNames = FALSE)
    openxlsx::writeData(wb, hoja, t(h2), startRow = row0 + 1, startCol = col0, colNames = FALSE)
    openxlsx::writeData(wb, hoja, t(h3), startRow = row0 + 2, startCol = col0, colNames = FALSE)

    # columna A con header_A, resto con header
    if (ncols >= 1) {
      openxlsx::addStyle(wb, hoja, st$header_A,
                         rows = row0:(row0 + 2), cols = col0, gridExpand = TRUE)
    }
    if (ncols >= 2) {
      openxlsx::addStyle(wb, hoja, st$header,
                         rows = row0:(row0 + 2),
                         cols  = (col0 + 1):(col0 + ncols - 1),
                         gridExpand = TRUE)
    }

    merge_runs <- function(v) {
      if (!length(v)) return(list())
      res <- list(); s <- 1
      for (i in seq_along(v)) {
        if (i == length(v) || v[i] != v[i + 1]) {
          res[[length(res) + 1]] <- c(s, i); s <- i + 1
        }
      }
      res
    }

    runs1 <- merge_runs(h1)
    for (r in runs1) if ((r[2] - r[1] + 1) > 1) {
      openxlsx::mergeCells(wb, hoja,
                           rows = row0,
                           cols = (col0 + r[1] - 1):(col0 + r[2] - 1))
    }

    runs2 <- merge_runs(h2)
    for (r in runs2) if ((r[2] - r[1] + 1) > 1) {
      openxlsx::mergeCells(wb, hoja,
                           rows = row0 + 1,
                           cols = (col0 + r[1] - 1):(col0 + r[2] - 1))
    }

    invisible(NULL)
  }

  # ================== LOOP POR SECCIONES ==================
  for (sec in names(SECCIONES)) {
    vars_sec <- SECCIONES[[sec]]
    if (!length(vars_sec)) next

    # título sección
    openxlsx::writeData(wb, hoja, toupper(sec), startRow = fila, startCol = 1)
    openxlsx::addStyle(wb, hoja, st$sec_title, rows = fila, cols = 1, gridExpand = TRUE)
    openxlsx::mergeCells(wb, hoja, rows = fila, cols = 1:3)
    fila <- fila + 2

    # ---------- loop por variable de la sección ----------
    for (var in vars_sec) {

      tp <- tipo_pregunta(var, survey = survey, sm_vars_force = sm_vars_force)
      cats_var <- get_categorias(
        var              = var,
        data             = data,
        survey           = survey,
        orders_list      = orders_list,
        opciones_excluir = opciones_excluir
      )

      opciones  <- cats_var$labels
      codes_row <- cats_var$codes

      qlab <- label_variable(var, dic_vars, labels_override, data)

      if (!length(opciones)) {
        openxlsx::writeData(wb, hoja, qlab, startRow = fila, startCol = 1)
        openxlsx::addStyle(wb, hoja, st$q_title, rows = fila, cols = 1, gridExpand = TRUE)
        openxlsx::mergeCells(wb, hoja, rows = fila, cols = 1:6)
        fila <- fila + 1
        openxlsx::writeData(wb, hoja, "Sin datos válidos para cruzar.", startRow = fila, startCol = 1)
        openxlsx::addStyle(wb, hoja, st$body_txt, rows = fila, cols = 1, gridExpand = TRUE)
        fila <- fila + 2
        next
      }

      # cuerpo base
      cuerpo <- tibble::tibble(Opciones = opciones)
      denom_map        <- list()
      estratos_totales <- list()

      # ---------- TOTAL ----------
      mask_total <- rep(TRUE, nrow(data))
      N_total <- denominador_validos(
        data       = data,
        var        = var,
        codes      = codes_row,
        tp         = tp,
        mask       = mask_total,
        weight_col = weight_col
      )
      n_total <- contar_por_opcion(
        data       = data,
        var        = var,
        codes      = codes_row,
        tp         = tp,
        mask       = mask_total,
        weight_col = weight_col
      )
      pct_total <- if (N_total > 0) n_total / N_total else rep(NA_real_, length(n_total))

      cuerpo <- dplyr::bind_cols(
        cuerpo,
        tibble::tibble(
          Total__n   = as.numeric(n_total),
          Total__pct = as.numeric(pct_total)
        )
      )
      denom_map[["Total__n"]] <- N_total

      # ---------- Cruces con cada variable de CRUZAR_CON ----------
      for (s in CRUZAR_CON) {
        if (!(s %in% names(data)) || identical(s, var)) next

        cats_s <- get_categorias(
          var              = s,
          data             = data,
          survey           = survey,
          orders_list      = orders_list,
          opciones_excluir = NULL
        )
        estr_codes  <- cats_s$codes
        estr_labels <- cats_s$labels

        if (!length(estr_codes)) next

        estratos_totales[[s]] <- list(codes = estr_codes, labels = estr_labels)

        # Valores reales en la data para el estrato
        v_estr <- as.character(data[[s]])

        # Detectar si la data usa códigos o labels
        usa_codes  <- any(v_estr %in% estr_codes)
        usa_labels <- any(v_estr %in% estr_labels)

        # Clave con la que vamos a comparar en la data:
        #   - si la data contiene los códigos, usamos códigos
        #   - si no, pero contiene labels, usamos labels
        #   - si ambos, priorizamos códigos (es el caso clásico haven/labelled)
        keys_vec <- if (usa_codes || !usa_labels) estr_codes else estr_labels

        bloques <- lapply(seq_along(keys_vec), function(j) {
          key_j  <- keys_vec[j]
          mask_s <- !is.na(v_estr) & v_estr == key_j

          n_vec <- contar_por_opcion(
            data       = data,
            var        = var,
            codes      = codes_row,
            tp         = tp,
            mask       = mask_s,
            weight_col = weight_col
          )

          N <- denominador_validos(
            data       = data,
            var        = var,
            codes      = codes_row,
            tp         = tp,
            mask       = mask_s,
            weight_col = weight_col
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

        cols_df   <- dplyr::bind_cols(lapply(bloques, `[[`, "df"))
        idx_n_cols <- grep("__n$", names(cols_df))
        Ns         <- vapply(bloques, `[[`, numeric(1), "N")
        if (length(idx_n_cols) == length(Ns) && length(Ns) > 0) {
          for (k in seq_along(idx_n_cols)) {
            denom_map[[names(cols_df)[idx_n_cols[k]]]] <- Ns[k]
          }
        }

        cuerpo <- dplyr::bind_cols(cuerpo, cols_df)
      }

      # ---------- fila Total ----------
      total_row <- as.list(rep(NA, ncol(cuerpo)))
      names(total_row) <- names(cuerpo)
      total_row[["Opciones"]] <- "Total"

      n_cols   <- grep("__n$",   names(cuerpo))
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

      # ---------- título de la pregunta ----------
      ncols_tbl <- ncol(cuerpo)
      openxlsx::writeData(wb, hoja, qlab, startRow = fila, startCol = 1)
      openxlsx::addStyle(wb, hoja, st$q_title, rows = fila, cols = 1, gridExpand = TRUE)
      openxlsx::mergeCells(wb, hoja, rows = fila, cols = 1:ncols_tbl)
      openxlsx::addStyle(wb, hoja, st$table_end,
                         rows = fila, cols = 1:ncols_tbl,
                         gridExpand = TRUE, stack = TRUE)
      fila <- fila + 1

      # ---------- encabezados (3 niveles) ----------
      ncols_total <- ncol(cuerpo)
      hdr1_full <- rep("", ncols_total)
      hdr2_full <- rep("", ncols_total)
      hdr3_full <- rep("", ncols_total)

      # Columna A vacía en las tres filas
      hdr1_full[1] <- ""
      hdr2_full[1] <- ""
      hdr3_full[1] <- ""

      col_ptr <- 2L

      # Bloque Total: fila 2 "Total", fila 3 "n", "%"
      if (any(names(cuerpo) == "Total__n")) {
        hdr1_full[col_ptr:(col_ptr + 1)] <- ""
        hdr2_full[col_ptr:(col_ptr + 1)] <- "Total"
        hdr3_full[col_ptr:(col_ptr + 1)] <- c("n", "%")
        col_ptr <- col_ptr + 2L
      }

      # Bloques por variables de cruce
      if (length(estratos_totales)) {
        for (s in CRUZAR_CON) {
          info_s <- estratos_totales[[s]]
          if (is.null(info_s)) next
          estr_labels <- info_s$labels
          if (!length(estr_labels)) next

          s_lbl <- label_variable(s, dic_vars, labels_override, data)

          for (lab in estr_labels) {
            if (col_ptr > ncols_total) break
            hdr1_full[col_ptr:(col_ptr + 1)] <- s_lbl
            hdr2_full[col_ptr:(col_ptr + 1)] <- rep(as.character(lab), 2)
            hdr3_full[col_ptr:(col_ptr + 1)] <- c("n", "%")
            col_ptr <- col_ptr + 2L
          }
        }
      }

      if (col_ptr <= ncols_total) {
        remaining <- col_ptr:ncols_total
        hdr1_full[remaining] <- ""
        hdr2_full[remaining] <- ""
        hdr3_full[remaining] <- rep(c("n", "%"), length.out = length(remaining))
      }

      escribir_encabezado(hdr1_full, hdr2_full, hdr3_full, row0 = fila, col0 = 1)

      openxlsx::addStyle(wb, hoja, st$table_end,
                         rows = fila + 2, cols = 1:ncols_total,
                         gridExpand = TRUE, stack = TRUE)

      fila <- fila + 3

      # ---------- cuerpo en Excel ----------
      openxlsx::writeData(wb, hoja, cuerpo, startRow = fila, startCol = 1, colNames = FALSE)

      nfil     <- nrow(cuerpo)
      ncol_tbl <- ncol(cuerpo)

      openxlsx::addStyle(wb, hoja, st$body_txt,
                         rows = fila:(fila + nfil - 1), cols = 1,
                         gridExpand = TRUE)

      if (ncol_tbl > 1) {
        is_pct <- grepl("__pct$", names(cuerpo))
        pct_cols_w <- which(is_pct)
        int_cols   <- setdiff(2:ncol_tbl, pct_cols_w)

        if (length(int_cols)) {
          openxlsx::addStyle(wb, hoja, st$body_int,
                             rows = fila:(fila + nfil - 1),
                             cols  = int_cols,
                             gridExpand = TRUE)
        }
        if (length(pct_cols_w)) {
          openxlsx::addStyle(wb, hoja, st$body_pct,
                             rows = fila:(fila + nfil - 1),
                             cols  = pct_cols_w,
                             gridExpand = TRUE)
        }
      }

      fila_total <- fila + nfil - 1
      if (length(n_cols)) {
        openxlsx::addStyle(wb, hoja, st$body_int,
                           rows = fila_total, cols = n_cols,
                           gridExpand = TRUE, stack = TRUE)
        openxlsx::addStyle(wb, hoja, st$total_bold,
                           rows = fila_total, cols = 1,
                           gridExpand = TRUE, stack = TRUE)
      }
      if (length(pct_cols)) {
        openxlsx::addStyle(wb, hoja, st$body_pct,
                           rows = fila_total, cols = pct_cols,
                           gridExpand = TRUE, stack = TRUE)
      }
      openxlsx::addStyle(wb, hoja, st$table_end,
                         rows = fila_total, cols = 1:ncol_tbl,
                         gridExpand = TRUE, stack = TRUE)

      fila <- fila + nfil

      # ---------- pie de tabla ----------
      pie_txt <- sprintf("Fuente: %s", fuente)
      openxlsx::writeData(wb, hoja, pie_txt, startRow = fila, startCol = 1)
      openxlsx::addStyle(wb, hoja, st$note, rows = fila, cols = 1, gridExpand = TRUE)
      openxlsx::mergeCells(wb, hoja, rows = fila, cols = 1:ncol_tbl)
      fila <- fila + 1

      # ---------- tabla de significancia (letras) ----------
      if (isTRUE(show_sig) && length(estratos_totales)) {

        letras_map_text <- c()
        bloques_sig     <- list()
        sig_h1          <- c("")  # col A vacía
        sig_h2          <- c("")

        for (s in CRUZAR_CON) {
          info_s <- estratos_totales[[s]]
          if (is.null(info_s)) next
          estr_labels <- info_s$labels
          estr_codes  <- info_s$codes
          if (!length(estr_labels)) next

          s_lbl <- label_variable(s, dic_vars, labels_override, data)

          # encabezados
          sig_h1 <- c(sig_h1, rep(s_lbl, length(estr_labels)))
          col_letters <- LETTERS[seq_along(estr_labels)]
          sig_h2 <- c(sig_h2, paste0(estr_labels, " (", col_letters, ")"))

          letras_map_text <- c(
            letras_map_text,
            paste0(
              s_lbl, ": ",
              paste0("(", col_letters, ") ", estr_labels, collapse = " · ")
            )
          )

          # matriz n/N para esta variable de cruce
          nn <- nN_para_sig_simple(
            data            = data,
            var             = var,
            opciones_labels = opciones,
            codes_row       = codes_row,
            estratos        = estr_codes,  # columnas por CÓDIGOS
            var_estrato     = s,
            tp              = tp,
            weight_col      = weight_col
          )

          cmp <- comparar_columnas_sig(nn$n_mat, nn$N_vec, alpha = alpha)

          bloques_sig[[s]] <- list(
            opciones    = opciones,
            estr_codes  = estr_codes,
            estr_labels = estr_labels,
            letras      = cmp$letras,
            sig         = cmp$sig
          )
        }

        if (length(bloques_sig)) {

          # ---- título de la tabla de letras ----
          t_sig <- "Comparaciones de proporciones de columna"
          openxlsx::writeData(wb, hoja, t_sig, startRow = fila, startCol = 1)
          openxlsx::addStyle(wb, hoja, st$q_title, rows = fila, cols = 1, gridExpand = TRUE)

          # número real de columnas en la tabla de letras
          ncols_sig <- length(sig_h1)
          if (ncols_sig < 2) ncols_sig <- 2

          openxlsx::mergeCells(wb, hoja, rows = fila, cols = 1:ncols_sig)
          openxlsx::addStyle(wb, hoja, st$table_end,
                             rows = fila, cols = 1:ncols_sig,
                             gridExpand = TRUE, stack = TRUE)
          fila <- fila + 1

          # ---- encabezados (2 filas) ----
          openxlsx::writeData(wb, hoja, t(sig_h1),
                              startRow = fila, startCol = 1, colNames = FALSE)
          openxlsx::writeData(wb, hoja, t(sig_h2),
                              startRow = fila + 1, startCol = 1, colNames = FALSE)
          openxlsx::addStyle(wb, hoja, st$header,
                             rows = fila:(fila + 1),
                             cols  = 1:ncols_sig,
                             gridExpand = TRUE)

          # merges por bloques de variable (fila 1)
          merge_runs <- function(v) {
            if (!length(v)) return(list())
            res <- list(); s <- 1
            for (i in seq_along(v)) {
              if (i == length(v) || v[i] != v[i + 1]) {
                res[[length(res) + 1]] <- c(s, i); s <- i + 1
              }
            }
            res
          }
          runs1 <- merge_runs(sig_h1)
          for (r in runs1) if ((r[2] - r[1] + 1) > 1) {
            # OJO: aquí ya NO se suma +1, se usa el índice tal cual
            openxlsx::mergeCells(wb, hoja,
                                 rows = fila,
                                 cols = r[1]:r[2])
          }

          fila_datos <- fila + 2

          # ---- filas de opciones ----
          openxlsx::writeData(wb, hoja, opciones,
                              startRow = fila_datos, startCol = 1, colNames = FALSE)
          openxlsx::addStyle(wb, hoja, st$cell,
                             rows = fila_datos:(fila_datos + length(opciones) - 1),
                             cols  = 1, gridExpand = TRUE)

          # ---- letras por columna ----
          col_cursor <- 2
          for (s in CRUZAR_CON) {
            bl <- bloques_sig[[s]]
            if (is.null(bl)) next

            for (j in seq_along(bl$estr_labels)) {
              col_let <- character(length(opciones))
              col_sig <- logical(length(opciones))

              rr <- match(opciones, bl$opciones)
              cc <- j  # columna j en n_mat/letras

              ok <- which(!is.na(rr))
              if (length(ok)) {
                col_let[ok] <- bl$letras[cbind(rr[ok], cc)]
                col_sig[ok] <- bl$sig[cbind(rr[ok], cc)]
              }

              openxlsx::writeData(
                wb, hoja, col_let,
                startRow = fila_datos, startCol = col_cursor,
                colNames = FALSE
              )
              openxlsx::addStyle(
                wb, hoja, st$cell,
                rows = fila_datos:(fila_datos + length(opciones) - 1),
                cols  = col_cursor,
                gridExpand = TRUE
              )
              if (any(col_sig)) {
                openxlsx::addStyle(
                  wb, hoja, st$fill_sig,
                  rows = fila_datos - 1 + which(col_sig),
                  cols  = col_cursor,
                  gridExpand = TRUE, stack = TRUE
                )
              }

              col_cursor <- col_cursor + 1
            }
          }

          # ---- pie de la tabla de letras ----
          pie_sig <- paste0(
            "Las letras indican columnas cuya proporción es significativamente mayor ",
            "que la proporción de la columna marcada por esa letra, según pruebas z ",
            "de diferencia de proporciones con corrección de Bonferroni para ",
            "comparaciones múltiples (α = ", alpha, "). ",
            "'.a' indica categoría excluida del contraste (proporciones 0 o 1).\n",
            "Letras por estrato: ",
            paste(letras_map_text, collapse = "  |  "),
            "\nFuente: ", fuente
          )

          fila_note <- fila_datos + length(opciones)
          openxlsx::writeData(wb, hoja, pie_sig,
                              startRow = fila_note, startCol = 1)
          openxlsx::addStyle(wb, hoja, st$note,
                             rows = fila_note, cols = 1, gridExpand = TRUE)
          openxlsx::mergeCells(wb, hoja, rows = fila_note, cols = 1:(col_cursor - 1))
          openxlsx::setRowHeights(wb, hoja, rows = fila_note, heights = 80)

          fila <- fila_note + 2
        } else {
          fila <- fila + 1
        }
      }

      fila <- fila + 1
    }

    fila <- fila + 1
  }

  openxlsx::setColWidths(wb, hoja, cols = 1, widths = 52)
  openxlsx::setColWidths(wb, hoja, cols = 2:200, widths = 12)

  openxlsx::saveWorkbook(wb, path_xlsx, overwrite = TRUE)
  message("✅ Cruces exportados a: ", normalizePath(path_xlsx))
  invisible(path_xlsx)
}

# =============================================================================
# Wrapper de alto nivel: reporte_cruces
# =============================================================================

#' Generar reporte de cruces en Excel
#'
#' Envuelve a \code{exportar_cruces_multi()} usando como insumo el objeto de
#' instrumento devuelto por \code{reporte_instrumento()}, de manera coherente
#' con \code{reporte_instrumento()}, \code{reporte_data()},
#' \code{reporte_frecuencias()} y \code{reporte_codebook()}.
#'
#' @param data Data frame ya adaptado por \code{reporte_data()} (por lo general
#'   la salida de esa función), que contiene las variables a cruzar y, de ser
#'   el caso, la variable de pesos.
#' @param instrumento Objeto devuelto por \code{reporte_instrumento()}, que
#'   debe contener al menos \code{$survey} y, cuando sea posible,
#'   \code{$orders_list}.
#' @param SECCIONES Lista nombrada que agrupa las variables a cruzar. Cada
#'   elemento es un vector de nombres de variables que se mostrarán en filas
#'   bajo el título de la sección.
#' @param cruces Vector de nombres de variables con las que se cruzarán todas
#'   las variables de \code{SECCIONES}. Una variable no se cruza consigo misma.
#' @param path_xlsx Ruta del archivo .xlsx de salida.
#' @param hoja Nombre de la hoja del libro de Excel.
#' @param fuente Texto para el pie de las tablas.
#' @param labels_override Lista nombrada opcional para sobrescribir etiquetas
#'   de preguntas.
#' @param sm_vars_force Vector opcional de nombres de variables a tratar como
#'   \emph{select multiple}.
#' @param weight_col Nombre de la variable de pesos.
#' @param opciones_excluir Vector de labels de opciones a excluir de las
#'   tablas (por ejemplo, categorías de no respuesta).
#' @param show_sig Lógico; si TRUE, muestra tablas de letras de significancia.
#' @param alpha Nivel de significancia (para pruebas z con corrección de
#'   Bonferroni).
#'
#' @return Invisiblemente, la ruta normalizada al archivo de salida.
#' @export
reporte_cruces <- function(
    data,
    instrumento,
    SECCIONES,
    cruces,
    path_xlsx        = "cruces_multi.xlsx",
    hoja             = "Cruces",
    fuente           = "Pulso PUCP",
    labels_override  = NULL,
    sm_vars_force    = NULL,
    weight_col       = "peso",
    opciones_excluir = NULL,
    show_sig         = TRUE,
    alpha            = 0.05
) {

  survey <- NULL
  if (!is.null(instrumento) && "survey" %in% names(instrumento)) {
    survey <- instrumento$survey
  }

  orders_list <- NULL
  if (!is.null(instrumento) && "orders_list" %in% names(instrumento)) {
    orders_list <- instrumento$orders_list
  }

  dic_vars <- NULL
  if (!is.null(survey) && all(c("name", "label") %in% names(survey))) {
    dic_vars <- dplyr::select(survey, name, label)
    dic_vars <- dplyr::filter(dic_vars, !is.na(name) & name != "")
  }

  exportar_cruces_multi(
    data            = data,
    dic_vars        = dic_vars,
    SECCIONES       = SECCIONES,
    CRUZAR_CON      = cruces,
    labels_override = labels_override,
    path_xlsx       = path_xlsx,
    hoja            = hoja,
    fuente          = fuente,
    survey          = survey,
    sm_vars_force   = sm_vars_force,
    weight_col      = weight_col,
    orders_list     = orders_list,
    opciones_excluir = opciones_excluir,
    show_sig        = show_sig,
    alpha           = alpha
  )
}
