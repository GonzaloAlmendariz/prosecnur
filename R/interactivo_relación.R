# =============================================================================
# Pestaña: Relación (Cruces) — versión minimalista/cliente
# - Selector: Sección + Variable (principal) y Sección + Variable (cruce)
# - Cruce: internamente solo SO (no se comunica en UI)
# - Sin "Opciones" (todo interno)
# - SM con demasiadas opciones: solo tabla (no gráfico)
# - Estratos sin datos: se omiten en el gráfico (no barras vacías)
# - Fix: DT::withTags -> htmltools::withTags
# =============================================================================

# -----------------------------------------------------------------------------
# UI del módulo (MINIMAL)
# -----------------------------------------------------------------------------
relacion_tab_ui <- function(id) {
  ns <- shiny::NS(id)

  shiny::tabPanel(
    title = "Relación",
    shiny::sidebarLayout(
      shiny::sidebarPanel(
        width = 3,

        shiny::h3("Relación"),

        shiny::selectInput(
          inputId = ns("main_seccion"),
          label   = "Sección (variable)",
          choices = NULL
        ),
        shiny::selectInput(
          inputId = ns("main_var"),
          label   = "Variable",
          choices = NULL
        ),

        shiny::hr(),

        shiny::selectInput(
          inputId = ns("cruce_seccion"),
          label   = "Sección (cruce)",
          choices = NULL
        ),
        shiny::selectInput(
          inputId = ns("cruce_var"),
          label   = "Cruce",
          choices = NULL
        )
      ),

      shiny::mainPanel(
        width = 9,

        shiny::fluidRow(
          shiny::column(
            width = 12,
            shiny::div(
              class = "cardbox",
              shiny::div(class = "cardbox-header", shiny::uiOutput(ns("rel_plot_header"))),
              plotly::plotlyOutput(ns("rel_plot"), height = "520px")
            )
          )
        ),

        shiny::br(),

        shiny::fluidRow(
          shiny::column(
            width = 12,
            shiny::div(
              class = "cardbox",
              shiny::div(
                class = "cardbox-header",
                shiny::div(class = "cardbox-title", "Tabla de cruces")
              ),
              DT::dataTableOutput(ns("rel_tabla"))
            )
          )
        ),

        shiny::div(style = "height: 48px;")
      )
    )
  )
}

# -----------------------------------------------------------------------------
# Server del módulo
# -----------------------------------------------------------------------------
relacion_tab_server <- function(
    id,
    data,
    instrumento,
    secciones,                   # lista nombrada: sección -> vector vars
    vars_so,                     # variables SO disponibles
    vars_sm_madres,              # variables madres SM disponibles
    colores_apiladas_por_listname = NULL,
    codigos_perdidos = NULL,
    weight_col = "peso",
    orders_list = NULL,          # opcional
    labels_override = NULL       # opcional
) {
  shiny::moduleServer(id, function(input, output, session) {

    # =========================================================================
    # Parámetros internos (no UI)
    # =========================================================================
    MAX_SM_PLOT <- 12L  # si SM tiene más opciones que esto, se omite gráfico

    `%||%` <- get0("%||%", ifnotfound = function(x, y) if (!is.null(x)) x else y)

    .wrap_titulo_html <- get0(
      ".wrap_titulo_html",
      ifnotfound = function(txt, width = 110) {
        if (!requireNamespace("stringr", quietly = TRUE)) return(as.character(txt))
        if (is.null(txt)) return("")
        lineas <- stringr::str_wrap(as.character(txt), width = width)
        paste(lineas, collapse = "<br>")
      }
    )

    .obtener_label_var <- get0(
      ".obtener_label_var",
      ifnotfound = function(var, instrumento, data = NULL) {
        surv <- instrumento$survey
        if (!is.null(surv) && all(c("name","label") %in% names(surv)) && var %in% surv$name) {
          lab <- surv$label[surv$name == var][1]
          if (!is.na(lab) && nzchar(as.character(lab))) return(as.character(lab))
        }
        if (!is.null(data) && var %in% names(data)) {
          vl <- attr(data[[var]], "label", exact = TRUE)
          if (!is.null(vl) && nzchar(as.character(vl))) return(as.character(vl))
        }
        as.character(var)
      }
    )

    # =========================================================================
    # Helpers base
    # =========================================================================
    get_pesos <- function(data, weight_col = "peso") {
      if (!is.null(weight_col) && weight_col %in% names(data)) {
        w <- suppressWarnings(as.numeric(data[[weight_col]]))
        w[is.na(w) | !is.finite(w)] <- 0
        return(w)
      }
      rep(1, nrow(data))
    }

    .has_var_or_dummies <- function(data, var) {
      if (!is.data.frame(data)) return(FALSE)
      if (var %in% names(data)) return(TRUE)
      var_esc <- gsub("([\\W])", "\\\\\\1", var)
      any(grepl(paste0("^", var_esc, "[/\\.]"), names(data)))
    }

    tipo_pregunta <- function(var, survey = NULL, sm_vars_force = NULL, data = NULL) {
      if (!is.null(sm_vars_force) && var %in% sm_vars_force) return("sm")
      if (!is.null(survey) && any(survey$name == var)) {
        tipos <- unique(na.omit(survey$type[survey$name == var]))
        if (any(grepl("^select_multiple(\\s|$)", tipos))) return("sm")
        if (any(grepl("^select_one(\\s|$)", tipos)))      return("so")
      }
      if (!is.null(data) && .has_var_or_dummies(data, var) && !(var %in% names(data))) {
        return("sm")
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
        if (!is.null(vlab) && nzchar(as.character(vlab))) return(as.character(vlab))
      }
      if (!is.null(dic_vars) && all(c("name", "label") %in% names(dic_vars))) {
        lab <- dic_vars$label[dic_vars$name == var]
        if (length(lab) && !all(is.na(lab))) return(as.character(lab[1]))
      }
      as.character(var)
    }

    get_list_name <- function(var, survey = NULL) {
      if (is.null(survey) || !all(c("name","list_name") %in% names(survey))) return(NA_character_)
      ln <- unique(na.omit(as.character(survey$list_name[survey$name == var])))
      if (!length(ln)) return(NA_character_)
      ln[1]
    }

    get_categorias <- function(var,
                               data,
                               survey          = NULL,
                               orders_list     = NULL,
                               opciones_excluir = NULL) {

      x <- if (var %in% names(data)) data[[var]] else NULL
      lab_attr <- if (!is.null(x)) attr(x, "labels", exact = TRUE) else NULL

      ln <- get_list_name(var, survey)
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
        ok2 <- !(labels %in% opciones_excluir)
        codes  <- codes[ok2]
        labels <- labels[ok2]
      }

      list(codes = codes, labels = labels, list_name = ln)
    }

    contar_por_opcion <- function(data, var, codes, tp, mask, weight_col = "peso") {
      w <- get_pesos(data, weight_col)

      if (tp == "so") {
        v_codes <- as.character(data[[var]])
        elig    <- mask & !is.na(v_codes) & nzchar(v_codes) & v_codes != "NA"
        return(vapply(seq_along(codes), function(j) sum(w[elig & v_codes == codes[j]], na.rm = TRUE), numeric(1)))
      }

      if (tp == "sm") {
        colc <- col_sm_compact(data, var)

        if (!is.na(colc)) {
          long <- sm_compact_to_long(data[[colc]], id = seq_len(nrow(data)), w = w)
          if (!nrow(long)) return(rep(0, length(codes)))
          ids_mask <- which(mask)
          long <- long[long$id %in% ids_mask & long$valor %in% codes, , drop = FALSE]
          return(vapply(seq_along(codes), function(j) {
            code_j <- codes[j]
            ids_j  <- unique(long$id[long$valor == code_j])
            sum(w[ids_j], na.rm = TRUE)
          }, numeric(1)))
        }

        # dummies
        if (!requireNamespace("stringr", quietly = TRUE)) return(rep(0, length(codes)))
        subs <- grep(paste0("^", stringr::fixed(var), "[/\\.]"), names(data), value = TRUE)
        if (!length(subs)) return(rep(0, length(codes)))
        codes_dummy <- sub(paste0("^", var, "[/\\.]"), "", subs)

        return(vapply(seq_along(codes), function(j) {
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
        }, numeric(1)))
      }

      rep(0, length(codes))
    }

    denominador_validos <- function(data, var, codes, tp, mask, weight_col = "peso") {
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
        }

        if (!requireNamespace("stringr", quietly = TRUE)) return(0)
        subs <- grep(paste0("^", stringr::fixed(var), "[/\\.]"), names(data), value = TRUE)
        if (!length(subs)) return(0)
        codes_dummy <- sub(paste0("^", var, "[/\\.]"), "", subs)
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

      0
    }

    .resolver_paleta_var <- function(var, instrumento, colores_apiladas_por_listname, opcion_levels) {
      surv <- instrumento$survey
      pal  <- NULL

      if (!is.null(colores_apiladas_por_listname) &&
          !is.null(surv) &&
          all(c("name", "list_name") %in% names(surv))) {
        ln <- surv$list_name[surv$name == var][1]
        if (!is.na(ln) && ln %in% names(colores_apiladas_por_listname)) {
          pal <- colores_apiladas_por_listname[[ln]]
        }
      }

      if (is.null(pal) || !length(pal)) {
        out <- grDevices::hcl.colors(max(3L, length(opcion_levels)), "Blues")
        out <- out[seq_len(length(opcion_levels))]
        names(out) <- opcion_levels
        return(out)
      }

      if (!is.null(names(pal)) && all(opcion_levels %in% names(pal))) {
        pal2 <- pal[opcion_levels]
        names(pal2) <- opcion_levels
        return(pal2)
      }

      fila <- instrumento$survey[instrumento$survey$name == var, , drop = FALSE]
      list_var <- if (nrow(fila)) fila$list_name[1] else NA_character_

      if (!is.null(instrumento$choices) &&
          all(c("list_name", "name", "label") %in% names(instrumento$choices)) &&
          !is.na(list_var) && nzchar(list_var) &&
          !is.null(names(pal))) {

        ch <- instrumento$choices[instrumento$choices$list_name == list_var, , drop = FALSE]
        map_code_to_label <- stats::setNames(as.character(ch$label), as.character(ch$name))

        idx <- names(pal) %in% names(map_code_to_label)
        if (any(idx)) {
          pal_lab <- stats::setNames(pal[idx], map_code_to_label[names(pal)[idx]])
          if (!all(opcion_levels %in% names(pal_lab))) {
            falt <- setdiff(opcion_levels, names(pal_lab))
            extra <- grDevices::hcl.colors(max(3L, length(falt)), "Blues")
            extra <- extra[seq_len(length(falt))]
            pal_lab <- c(pal_lab, stats::setNames(extra, falt))
          }
          pal_lab <- pal_lab[opcion_levels]
          names(pal_lab) <- opcion_levels
          return(pal_lab)
        }
      }

      pal <- rep(pal, length.out = length(opcion_levels))
      names(pal) <- opcion_levels
      pal
    }

    # =========================================================================
    # Tabla (cuerpo) estilo reporte_cruces (semántica)
    # =========================================================================
    .build_cuerpo <- function(df, var_main, var_cruce) {
      survey <- instrumento$survey %||% NULL
      dic_vars <- NULL
      if (!is.null(survey) && all(c("name","label") %in% names(survey))) {
        dic_vars <- dplyr::select(survey, name, label)
      }

      tp_main <- tipo_pregunta(var_main, survey = survey, data = df)

      cats_main <- get_categorias(
        var          = var_main,
        data         = df,
        survey       = survey,
        orders_list  = orders_list %||% instrumento$orders_list %||% NULL,
        opciones_excluir = NULL
      )
      codes_row <- as.character(cats_main$codes)
      opciones  <- as.character(cats_main$labels)

      if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0 && length(codes_row)) {
        codp <- as.character(codigos_perdidos)
        keep <- !(codes_row %in% codp)
        codes_row <- codes_row[keep]
        opciones  <- opciones[keep]
      }

      cats_cruce <- get_categorias(
        var          = var_cruce,
        data         = df,
        survey       = survey,
        orders_list  = orders_list %||% instrumento$orders_list %||% NULL,
        opciones_excluir = NULL
      )
      estr_codes  <- as.character(cats_cruce$codes)
      estr_labels <- as.character(cats_cruce$labels)

      cuerpo <- tibble::tibble(Opciones = opciones)
      denom_map <- list()

      # Total
      mask_total <- rep(TRUE, nrow(df))
      N_total <- denominador_validos(df, var_main, codes_row, tp_main, mask_total, weight_col = weight_col)
      n_total <- contar_por_opcion(df, var_main, codes_row, tp_main, mask_total, weight_col = weight_col)
      pct_total <- if (N_total > 0) n_total / N_total else rep(0, length(n_total))

      cuerpo <- dplyr::bind_cols(
        cuerpo,
        tibble::tibble(
          Total__n   = as.numeric(n_total),
          Total__pct = as.numeric(pct_total)
        )
      )
      denom_map[["Total__n"]] <- N_total

      # Por estrato
      v_estr <- as.character(df[[var_cruce]])
      for (j in seq_along(estr_labels)) {
        key_j <- estr_codes[j]
        mask_j <- !is.na(v_estr) & v_estr == key_j

        n_vec <- contar_por_opcion(df, var_main, codes_row, tp_main, mask_j, weight_col = weight_col)
        N_j   <- denominador_validos(df, var_main, codes_row, tp_main, mask_j, weight_col = weight_col)
        pct <- if (N_j > 0) n_vec / N_j else rep(0, length(n_vec))

        nm_n   <- paste0(var_cruce, "__", make.names(estr_labels[j]), "__n")
        nm_pct <- paste0(var_cruce, "__", make.names(estr_labels[j]), "__pct")

        cuerpo <- dplyr::bind_cols(
          cuerpo,
          tibble::tibble(!!nm_n := as.numeric(n_vec), !!nm_pct := as.numeric(pct))
        )
        denom_map[[nm_n]] <- N_j
      }

      # Fila Total
      total_row <- as.list(rep(NA, ncol(cuerpo)))
      names(total_row) <- names(cuerpo)
      total_row[["Opciones"]] <- "Total"

      n_cols   <- grep("__n$",   names(cuerpo))
      pct_cols <- grep("__pct$", names(cuerpo))

      for (k in n_cols) {
        nm <- names(cuerpo)[k]
        Nj <- denom_map[[nm]]
        total_row[[k]] <- if (is.null(Nj)) NA_real_ else round(as.numeric(Nj), 0)
      }
      for (k in pct_cols) {
        n_partner <- sub("__pct$", "__n", names(cuerpo)[k])
        Nj <- suppressWarnings(as.numeric(total_row[[n_partner]]))
        total_row[[k]] <- if (!is.na(Nj) && Nj > 0) 1.0 else 0.0
      }

      cuerpo <- dplyr::bind_rows(cuerpo, tibble::as_tibble(total_row))

      cruce_lbl <- label_variable(var_cruce, dic_vars = dic_vars, labels_override = labels_override, data = df)

      list(
        cuerpo       = cuerpo,
        tipo_main    = tp_main,
        estr_labels  = estr_labels,
        cruce_lbl    = cruce_lbl,
        codes_main   = codes_row,
        labels_main  = opciones,
        estr_codes   = estr_codes
      )
    }

    # =========================================================================
    # Encabezado DT multi-nivel (fix withTags)
    # =========================================================================
    .dt_container_multihdr <- function(cuerpo, cruce_lbl, estr_labels) {

      n_blocks <- 1L + length(estr_labels)     # Total + estratos
      ncols    <- ncol(cuerpo)
      exp_cols <- 1L + 2L * n_blocks

      # Fallback seguro
      if (is.na(ncols) || ncols != exp_cols) {
        return(htmltools::withTags(
          table(
            class = "display nowrap compact",
            thead(
              tr(lapply(names(cuerpo), function(x) htmltools::tags$th(x)))
            )
          )
        ))
      }

      # ---------- fila 2: Total + estratos ----------
      fila2 <- c(
        list(htmltools::tags$th(colspan = 2, "Total")),
        lapply(
          estr_labels,
          function(lab) htmltools::tags$th(colspan = 2, as.character(lab))
        )
      )

      # ---------- fila 3: n / % ----------
      fila3 <- unlist(
        replicate(n_blocks, list(
          htmltools::tags$th("n"),
          htmltools::tags$th("%")
        ), simplify = FALSE),
        recursive = FALSE
      )

      htmltools::withTags(
        table(
          class = "display nowrap compact",
          thead(
            # Fila 1: encabezado superior
            tr(
              htmltools::tags$th(rowspan = 3, ""),
              htmltools::tags$th(colspan = ncols - 1, cruce_lbl)
            ),
            # Fila 2
            tr(fila2),
            # Fila 3
            tr(fila3)
          )
        )
      )
    }

    # =========================================================================
    # Gráficos
    # =========================================================================
    .plot_so_so <- function(df, var_main, var_cruce) {
      survey <- instrumento$survey %||% NULL

      cats_main <- get_categorias(var_main, df, survey, orders_list %||% instrumento$orders_list %||% NULL, NULL)
      codes_row <- as.character(cats_main$codes)
      opciones  <- as.character(cats_main$labels)

      if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0) {
        codp <- as.character(codigos_perdidos)
        keep <- !(codes_row %in% codp)
        codes_row <- codes_row[keep]
        opciones  <- opciones[keep]
      }

      cats_cruce <- get_categorias(var_cruce, df, survey, orders_list %||% instrumento$orders_list %||% NULL, NULL)
      estr_codes  <- as.character(cats_cruce$codes)
      estr_labels <- as.character(cats_cruce$labels)

      v_main  <- as.character(df[[var_main]])
      v_cruce <- as.character(df[[var_cruce]])
      w <- get_pesos(df, weight_col)

      rows <- list()

      for (j in seq_along(estr_codes)) {
        key_j <- estr_codes[j]
        mask_j <- !is.na(v_cruce) & v_cruce == key_j

        elig <- mask_j & !is.na(v_main) & nzchar(v_main) & v_main != "NA" & (v_main %in% codes_row)
        N_j  <- sum(w[elig], na.rm = TRUE)

        # ✅ estrato sin datos -> NO SE MUESTRA en gráfico
        if (is.na(N_j) || N_j <= 0) next

        for (i in seq_along(codes_row)) {
          code_i <- codes_row[i]
          n_ij <- sum(w[elig & v_main == code_i], na.rm = TRUE)
          rows[[length(rows) + 1]] <- data.frame(
            estrato_label = estr_labels[j],
            opcion_label  = opciones[i],
            pct = n_ij / N_j,
            n   = n_ij,
            stringsAsFactors = FALSE
          )
        }
      }

      df_tab <- dplyr::bind_rows(rows)
      if (!nrow(df_tab)) {
        return(plotly::plot_ly() |>
                 plotly::layout(annotations = list(list(text = "Sin datos para graficar.", showarrow = FALSE))))
      }

      pal <- .resolver_paleta_var(
        var = var_main,
        instrumento = instrumento,
        colores_apiladas_por_listname = colores_apiladas_por_listname,
        opcion_levels = unique(opciones)
      )

      df_tab$opcion_label  <- factor(df_tab$opcion_label, levels = opciones)
      df_tab$estrato_label <- factor(df_tab$estrato_label, levels = unique(df_tab$estrato_label))

      p <- plotly::plot_ly()
      for (opt in opciones) {
        dfo <- df_tab[df_tab$opcion_label == opt, , drop = FALSE]
        if (!nrow(dfo)) next

        dfo$hover <- sprintf(
          "%s<br>%s: %s%%<br>n: %s",
          as.character(dfo$estrato_label),
          opt,
          round(100 * dfo$pct, 1),
          format(round(dfo$n, 0), big.mark = ",")
        )

        p <- p |>
          plotly::add_bars(
            data          = dfo,
            x             = ~pct,
            y             = ~estrato_label,
            name          = opt,
            orientation   = "h",
            text          = ~paste0(round(100 * pct, 1), "%"),
            textposition  = "inside",
            insidetextanchor = "middle",
            textfont      = list(color = "white", size = 11),
            customdata    = ~hover,
            hovertemplate = "%{customdata}<extra></extra>",
            marker        = list(color = unname(pal[opt]), line = list(width = 0))
          )
      }

      p |>
        plotly::layout(
          barmode = "stack",
          bargap  = 0.25,
          xaxis   = list(title = "", range = c(0, 1), showgrid = FALSE, zeroline = FALSE,
                         showticklabels = FALSE, ticks = ""),
          yaxis   = list(title = "", automargin = TRUE, showgrid = FALSE, zeroline = FALSE, ticks = ""),
          legend  = list(orientation = "h", x = 0.5, xanchor = "center", y = -0.12),
          margin  = list(l = 170, r = 25, t = 10, b = 55),
          hovermode  = "closest",
          transition = list(duration = 450, easing = "cubic-in-out")
        ) |>
        plotly::config(displayModeBar = FALSE, responsive = TRUE)
    }

    .plot_sm_so <- function(df, var_main_sm, var_cruce_so) {
      survey <- instrumento$survey %||% NULL

      cats_sm <- get_categorias(var_main_sm, df, survey, orders_list %||% instrumento$orders_list %||% NULL, NULL)
      codes_all <- as.character(cats_sm$codes)
      labels_all <- as.character(cats_sm$labels)

      if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0) {
        codp <- as.character(codigos_perdidos)
        keep <- !(codes_all %in% codp)
        codes_all <- codes_all[keep]
        labels_all <- labels_all[keep]
      }

      # ✅ demasiadas opciones -> NO GRAFICAR (solo tabla)
      if (length(labels_all) > MAX_SM_PLOT) {
        return(plotly::plot_ly() |>
                 plotly::layout(annotations = list(list(
                   text = "Gráfico no disponible para esta variable.",
                   showarrow = FALSE
                 ))))
      }

      cats_cruce <- get_categorias(var_cruce_so, df, survey, orders_list %||% instrumento$orders_list %||% NULL, NULL)
      estr_codes  <- as.character(cats_cruce$codes)
      estr_labels <- as.character(cats_cruce$labels)

      tp_sm <- "sm"
      pal <- .resolver_paleta_var(
        var = var_main_sm,
        instrumento = instrumento,
        colores_apiladas_por_listname = colores_apiladas_por_listname,
        opcion_levels = unique(labels_all)
      )

      plots <- list()

      for (i in seq_along(codes_all)) {
        code_i <- codes_all[i]
        lab_i  <- labels_all[i]

        rows <- list()
        for (j in seq_along(estr_codes)) {
          key_j <- estr_codes[j]
          mask_j <- !is.na(df[[var_cruce_so]]) & as.character(df[[var_cruce_so]]) == key_j

          N_j <- denominador_validos(df, var_main_sm, codes_all, tp_sm, mask_j, weight_col)
          # ✅ estrato sin datos -> NO SE MUESTRA en gráfico
          if (is.na(N_j) || N_j <= 0) next

          n_vec <- contar_por_opcion(df, var_main_sm, codes = code_i, tp = tp_sm, mask = mask_j, weight_col = weight_col)
          n_ij <- as.numeric(n_vec[1])
          pct  <- n_ij / N_j

          rows[[length(rows) + 1]] <- data.frame(
            estrato_label = estr_labels[j],
            pct = pct,
            n   = n_ij,
            N   = N_j,
            stringsAsFactors = FALSE
          )
        }

        dfi <- dplyr::bind_rows(rows)
        if (!nrow(dfi)) next

        dfi$estrato_label <- factor(dfi$estrato_label, levels = unique(dfi$estrato_label))
        dfi$hover <- sprintf(
          "%s<br>%s: %s%%<br>n: %s",
          as.character(dfi$estrato_label),
          lab_i,
          round(100 * dfi$pct, 1),
          format(round(dfi$n, 0), big.mark = ",")
        )

        p_i <- plotly::plot_ly(
          data = dfi,
          x = ~pct,
          y = ~estrato_label,
          type = "bar",
          orientation = "h",
          text = ~paste0(round(100 * pct, 1), "%"),
          textposition = "inside",
          insidetextanchor = "middle",
          textfont = list(color = "white", size = 11),
          marker = list(color = unname(pal[lab_i]), line = list(width = 0)),
          customdata = ~hover,
          hovertemplate = "%{customdata}<extra></extra>"
        ) |>
          plotly::layout(
            title = list(text = .wrap_titulo_html(lab_i, width = 35), x = 0.02, xanchor = "left", font = list(size = 12)),
            xaxis = list(range = c(0, 1), showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE, ticks = ""),
            yaxis = list(title = "", automargin = TRUE, showgrid = FALSE, zeroline = FALSE),
            margin = list(l = 150, r = 10, t = 38, b = 10),
            showlegend = FALSE
          ) |>
          plotly::config(displayModeBar = FALSE, responsive = TRUE)

        plots[[length(plots) + 1]] <- p_i
      }

      if (!length(plots)) {
        return(plotly::plot_ly() |>
                 plotly::layout(annotations = list(list(text="Sin datos para graficar.", showarrow = FALSE))))
      }

      # Subplot en 2 columnas
      n_pl <- length(plots)
      ncol <- 2
      nrow <- ceiling(n_pl / ncol)

      plotly::subplot(
        plots,
        nrows   = nrow,
        margin  = 0.03,
        shareX  = TRUE,
        titleX  = FALSE,
        titleY  = TRUE
      ) |>
        plotly::layout(margin = list(l = 10, r = 10, t = 10, b = 10))
    }

    # =========================================================================
    # Wiring UI: secciones y variables (independientes)
    # =========================================================================
    secciones_limpias <- lapply(secciones, function(vs) {
      vs[vapply(vs, function(v) .has_var_or_dummies(data, v), logical(1))]
    })
    secciones_limpias <- secciones_limpias[vapply(secciones_limpias, length, integer(1)) > 0]

    shiny::observe({
      secs <- names(secciones_limpias)
      if (!length(secs)) {
        shiny::updateSelectInput(session, "main_seccion", choices = c())
        shiny::updateSelectInput(session, "cruce_seccion", choices = c())
      } else {
        shiny::updateSelectInput(session, "main_seccion",  choices = stats::setNames(secs, secs), selected = secs[1])
        shiny::updateSelectInput(session, "cruce_seccion", choices = stats::setNames(secs, secs), selected = secs[1])
      }
    })

    shiny::observeEvent(input$main_seccion, {
      sec <- input$main_seccion
      if (is.null(sec) || !nzchar(sec) || is.null(secciones_limpias[[sec]])) return()

      vars_sec <- secciones_limpias[[sec]]

      # Variable principal: SO + SM madres
      main_choices <- sort(unique(intersect(vars_sec, unique(c(vars_so, vars_sm_madres)))))
      if (!length(main_choices)) main_choices <- sort(unique(c(vars_so, vars_sm_madres)))

      main_lab <- stats::setNames(
        main_choices,
        vapply(main_choices, function(v) .obtener_label_var(v, instrumento, data), character(1))
      )

      shiny::updateSelectInput(session, "main_var", choices = main_lab, selected = main_choices[1] %||% "")
    }, ignoreInit = TRUE)

    shiny::observeEvent(input$cruce_seccion, {
      sec <- input$cruce_seccion
      if (is.null(sec) || !nzchar(sec) || is.null(secciones_limpias[[sec]])) return()

      vars_sec <- secciones_limpias[[sec]]

      # Cruce: internamente SOLO SO (sin decirlo)
      cruce_choices <- sort(unique(intersect(vars_sec, vars_so)))
      if (!length(cruce_choices)) cruce_choices <- sort(unique(vars_so))

      cruce_lab <- stats::setNames(
        cruce_choices,
        vapply(cruce_choices, function(v) .obtener_label_var(v, instrumento, data), character(1))
      )

      shiny::updateSelectInput(session, "cruce_var", choices = cruce_lab, selected = cruce_choices[1] %||% "")
    }, ignoreInit = TRUE)

    # Primera carga: forzar update de ambos selects
    shiny::observeEvent(names(secciones_limpias), {
      if (!is.null(input$main_seccion))  shiny::isolate(shiny::updateSelectInput(session, "main_seccion",  selected = input$main_seccion))
      if (!is.null(input$cruce_seccion)) shiny::isolate(shiny::updateSelectInput(session, "cruce_seccion", selected = input$cruce_seccion))
    }, once = TRUE)

    # =========================================================================
    # Header gráfico (minimal)
    # =========================================================================
    output$rel_plot_header <- shiny::renderUI({
      shiny::req(input$main_var, input$cruce_var)
      t_main  <- .wrap_titulo_html(.obtener_label_var(input$main_var, instrumento, data), width = 110)
      t_cruce <- .obtener_label_var(input$cruce_var, instrumento, data)

      shiny::tagList(
        shiny::div(class = "cardbox-title", shiny::HTML(t_main)),
        shiny::div(class = "cardbox-subtitle", paste0("Cruce: ", t_cruce))
      )
    })

    # =========================================================================
    # Reactives: objeto central
    # =========================================================================
    rel_obj <- shiny::reactive({
      shiny::req(input$main_var, input$cruce_var)

      var_main  <- input$main_var
      var_cruce <- input$cruce_var

      # seguridad: cruce debe ser SO (pero no se muestra en UI)
      if (!(var_cruce %in% vars_so)) {
        return(list(error = "No es posible cruzar con la selección actual."))
      }

      df <- data
      if (var_cruce %in% names(df)) df <- df[!is.na(df[[var_cruce]]), , drop = FALSE]
      if (!nrow(df)) return(list(error = "Sin datos disponibles."))

      out <- .build_cuerpo(df, var_main, var_cruce)
      out$df <- df
      out$var_main <- var_main
      out$var_cruce <- var_cruce
      out
    })

    # =========================================================================
    # Tabla
    # =========================================================================
    output$rel_tabla <- DT::renderDataTable({
      obj <- rel_obj()
      if (!is.null(obj$error)) {
        return(DT::datatable(
          data.frame(Mensaje = obj$error),
          rownames = FALSE,
          options = list(
            paging    = FALSE,
            searching = FALSE,
            info      = FALSE,
            ordering  = FALSE,
            orderCellsTop = TRUE,
            scrollX   = TRUE,
            language  = list(url = "//cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json"),
            columnDefs = list(
              list(className = "dt-center", targets = "_all")
            )
          )
        ))
      }

      cuerpo <- obj$cuerpo
      container <- .dt_container_multihdr(
        cuerpo = cuerpo,
        cruce_lbl = obj$cruce_lbl,
        estr_labels = obj$estr_labels
      )

      is_pct <- grepl("__pct$", names(cuerpo))
      is_n   <- grepl("__n$",   names(cuerpo))

      DT::datatable(
        cuerpo,
        rownames  = FALSE,
        container = container,
        options = list(
          paging    = FALSE,
          searching = FALSE,
          info      = FALSE,
          ordering  = FALSE,
          scrollX   = TRUE,
          language  = list(url = "//cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json"),
          columnDefs = list(
            list(className = "dt-left",  targets = 0),
            list(className = "dt-right", targets = which(is_n) - 1),
            list(className = "dt-right", targets = which(is_pct) - 1)
          )
        )
      ) |>
        DT::formatRound(columns = which(is_n), digits = 0) |>
        DT::formatPercentage(columns = which(is_pct), digits = 1)
    })

    # =========================================================================
    # Gráfico
    # =========================================================================
    output$rel_plot <- plotly::renderPlotly({
      obj <- rel_obj()
      if (!is.null(obj$error)) {
        return(plotly::plot_ly() |>
                 plotly::layout(annotations = list(list(text = obj$error, showarrow = FALSE))))
      }

      df        <- obj$df
      var_main  <- obj$var_main
      var_cruce <- obj$var_cruce

      survey <- instrumento$survey %||% NULL
      tp_main <- tipo_pregunta(var_main, survey = survey, data = df)

      if (identical(tp_main, "so")) {
        .plot_so_so(df, var_main, var_cruce)
      } else {
        .plot_sm_so(df, var_main, var_cruce)
      }
    })

    invisible(NULL)
  })
}
