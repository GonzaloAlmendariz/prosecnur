# =============================================================================
# Pestaña: Relación (Cruces) — versión theme-aware corregida
# -----------------------------------------------------------------------------
# - Usa theme_app recibido como argumento (sin get("theme_app", ...)).
# - Respeta paleta personalizada en títulos, subtítulos, cards, barras SM y leyenda SO.
# - SO: leyenda externa HTML, sin tocar orden del apilado.
# - SM: chips por opción, barras fill-only por estrato.
# - Estratos sin datos válidos: se omiten en cada chip.
# =============================================================================

# -----------------------------------------------------------------------------
# UI del módulo
# -----------------------------------------------------------------------------
relacion_tab_ui <- function(id) {
  ns <- shiny::NS(id)

  shiny::tabPanel(
    title = "Relación",
    shiny::sidebarLayout(
      shiny::sidebarPanel(
        width = 3,

        shiny::h3("Relación"),

        shiny::div(
          class = "cardbox",
          style = "padding:12px 12px 10px 12px;",

          shiny::div(
            style = "font-weight:600; color:#111827; margin-bottom:8px;",
            "Variable"
          ),

          shiny::selectInput(
            inputId = ns("main_seccion"),
            label   = "Sección",
            choices = NULL
          ),

          shiny::selectInput(
            inputId = ns("main_var"),
            label   = "Variable",
            choices = NULL
          ),

          shiny::div(
            style = "margin-top:-6px; margin-bottom:12px; font-size:12px; color:#6b7280;",
            "Se muestra la distribución de la variable elegida."
          ),

          shiny::hr(style = "margin:10px 0 12px 0; border-color:#eef2f7;"),

          shiny::div(
            style = "font-weight:600; color:#111827; margin-bottom:8px;",
            "Cruce"
          ),

          shiny::selectInput(
            inputId = ns("cruce_seccion"),
            label   = "Sección",
            choices = NULL
          ),

          shiny::selectInput(
            inputId = ns("cruce_var"),
            label   = "Segmento",
            choices = NULL
          ),

          shiny::div(
            style = "margin-top:-6px; font-size:12px; color:#6b7280;",
            "Se calcula dentro de cada grupo del cruce."
          )
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
              shiny::uiOutput(ns("rel_plot_ui")),
              shiny::uiOutput(ns("rel_so_legend"))
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
    secciones,
    vars_so,
    vars_sm_madres,
    colores_apiladas_por_listname = NULL,
    codigos_perdidos = NULL,
    weight_col = "peso",
    orders_list = NULL,
    labels_override = NULL,
    theme_app = NULL
) {

  shiny::moduleServer(id, function(input, output, session) {

    MAX_SM_CHIPS <- 14L
    BAR_HEIGHT   <- 52
    PCT_FSIZE    <- 12

    `%||%` <- get0("%||%", ifnotfound = function(x, y) if (!is.null(x)) x else y)

    # -------------------------------------------------------------------------
    # Tema visual
    # -------------------------------------------------------------------------
    theme_default <- if (exists("reporte_interactivo_theme_default", mode = "function")) {
      reporte_interactivo_theme_default()
    } else {
      list(
        color_primario      = "#002457",
        color_fondo_app     = "#f5f6fa",
        color_borde         = "#e6e9f2",
        color_texto         = "#1f2933",
        color_texto_suave   = "#5f6b7a",
        color_superficie    = "#ffffff",
        color_superficie_2  = "#fafbff",
        color_header_tabla  = "#f1f3f9"
      )
    }

    theme_rel <- theme_default
    if (!is.null(theme_app) && is.list(theme_app)) {
      nm <- intersect(names(theme_app), names(theme_rel))
      if (length(nm)) theme_rel[nm] <- theme_app[nm]
    }

    COLOR_PRIMARIO    <- theme_rel$color_primario
    COLOR_FONDO_APP   <- theme_rel$color_fondo_app
    COLOR_BORDE       <- theme_rel$color_borde
    COLOR_TEXTO       <- theme_rel$color_texto
    COLOR_TEXTO_SUAVE <- theme_rel$color_texto_suave
    COLOR_SUPERFICIE  <- theme_rel$color_superficie
    COLOR_SUPERFICIE2 <- theme_rel$color_superficie_2

    SM_COLOR_YES <- COLOR_PRIMARIO
    SM_COLOR_BG  <- COLOR_SUPERFICIE2

    # -------------------------------------------------------------------------
    # Helpers texto
    # -------------------------------------------------------------------------
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

        if (!is.null(surv) && "name" %in% names(surv)) {
          label_col <- if ("label" %in% names(surv)) {
            "label"
          } else {
            cand <- grep("^label(::|$)", names(surv), value = TRUE)
            if (length(cand)) cand[1] else NULL
          }

          if (!is.null(label_col) && label_col %in% names(surv)) {
            i <- which(!is.na(surv$name) & surv$name == var)[1]
            if (!is.na(i)) {
              lab <- surv[[label_col]][i]
              if (!is.na(lab) && nzchar(as.character(lab))) return(as.character(lab))
            }
          }
        }

        if (!is.null(data) && var %in% names(data)) {
          vl <- attr(data[[var]], "label", exact = TRUE)
          if (!is.null(vl) && nzchar(as.character(vl))) return(as.character(vl))
        }

        as.character(var)
      }
    )

    # -------------------------------------------------------------------------
    # Helpers base
    # -------------------------------------------------------------------------
    get_pesos <- function(df, weight_col = "peso") {
      if (!is.null(weight_col) && weight_col %in% names(df)) {
        w <- suppressWarnings(as.numeric(df[[weight_col]]))
        w[is.na(w) | !is.finite(w)] <- 0
        return(w)
      }
      rep(1, nrow(df))
    }

    .has_var_or_dummies <- function(df, var) {
      if (!is.data.frame(df)) return(FALSE)
      if (var %in% names(df)) return(TRUE)
      var_esc <- gsub("([\\W])", "\\\\\\1", var)
      any(grepl(paste0("^", var_esc, "[/\\.]"), names(df)))
    }

    tipo_pregunta <- function(var, survey = NULL, sm_vars_force = NULL, df = NULL) {
      if (!is.null(sm_vars_force) && var %in% sm_vars_force) return("sm")
      if (!is.null(survey) && any(survey$name == var)) {
        tipos <- unique(na.omit(survey$type[survey$name == var]))
        tipos <- tolower(as.character(tipos))
        if (any(grepl("^select_multiple(\\s|$)", tipos))) return("sm")
        if (any(grepl("^select_one(\\s|$)", tipos))) return("so")
      }
      if (!is.null(df) && .has_var_or_dummies(df, var) && !(var %in% names(df))) return("sm")
      "so"
    }

    get_list_name <- function(var, survey = NULL) {
      if (is.null(survey) || !all(c("name", "list_name") %in% names(survey))) return(NA_character_)
      ln <- unique(na.omit(as.character(survey$list_name[survey$name == var])))
      if (!length(ln)) return(NA_character_)
      ln[1]
    }

    .get_list_name_safe <- function(survey, var) {
      if (is.null(survey) || !all(c("name","list_name") %in% names(survey))) return(NA_character_)
      i <- which(!is.na(survey$name) & survey$name == var)[1]
      if (is.na(i)) return(NA_character_)
      ln <- as.character(survey$list_name[i])
      if (is.na(ln) || !nzchar(ln)) return(NA_character_)
      ln
    }

    # -------------------------------------------------------------------------
    # Resolver SM seguro
    # -------------------------------------------------------------------------
    .resolver_var_spec_safe <- function(var_madre, df) {

      f <- get0("resolver_var_spec", mode = "function", ifnotfound = NULL)
      if (!is.null(f)) {
        out <- tryCatch(
          f(var_madre = var_madre, ctx = list(data = df, instrumento = instrumento), df = df),
          error = function(e) NULL
        )
        if (is.list(out) && length(out$cols)) return(out)
      }

      var_esc <- gsub("([\\W])", "\\\\\\1", var_madre)
      cols <- grep(paste0("^", var_esc, "\\."), names(df), value = TRUE)

      surv <- instrumento$survey %||% NULL
      ch   <- instrumento$choices %||% NULL

      ln <- NA_character_
      if (!is.null(surv) && all(c("name","list_name") %in% names(surv)) && var_madre %in% surv$name) {
        ln <- as.character(surv$list_name[surv$name == var_madre][1])
      }

      map_code_to_label <- list()

      if (!is.null(ch) && all(c("list_name","name") %in% names(ch)) && !is.na(ln) && nzchar(ln)) {
        label_col <- if ("label" %in% names(ch)) {
          "label"
        } else {
          cand <- grep("^label(::|$)", names(ch), value = TRUE)
          if (length(cand)) cand[1] else NULL
        }

        if (!is.null(label_col) && label_col %in% names(ch)) {
          ch_v <- ch[ch$list_name == ln, , drop = FALSE]
          if (nrow(ch_v)) {
            m <- stats::setNames(as.character(ch_v[[label_col]]), as.character(ch_v$name))
            map_code_to_label <- as.list(m)
          }
        }
      }

      list(
        var_madre = var_madre,
        cols = cols,
        map_code_to_label = map_code_to_label,
        list_name = ln,
        col_compact = NA_character_
      )
    }

    # -------------------------------------------------------------------------
    # Categorías SO
    # -------------------------------------------------------------------------
    get_categorias_so <- function(var, df, survey = NULL, orders_list = NULL) {

      x <- df[[var]]
      lab_attr <- attr(x, "labels", exact = TRUE)

      ln <- get_list_name(var, survey)

      obj <- NULL
      if (!is.null(orders_list)) {
        if (var %in% names(orders_list)) obj <- orders_list[[var]]
        else if (!is.na(ln) && ln %in% names(orders_list)) obj <- orders_list[[ln]]
      }

      if (!is.null(obj)) {
        codes  <- as.character(obj$names)
        labels <- as.character(obj$labels)
      } else if (!is.null(lab_attr) && length(lab_attr) > 0) {
        codes  <- names(lab_attr)
        labels <- as.character(unname(lab_attr))
      } else {
        codes  <- sort(unique(na.omit(as.character(x))))
        labels <- codes
      }

      ok <- !is.na(codes) & nzchar(codes)
      codes  <- codes[ok]
      labels <- labels[ok]

      list(codes = codes, labels = labels, list_name = ln)
    }

    # -------------------------------------------------------------------------
    # SM: validez + numerador
    # -------------------------------------------------------------------------
    .sm_valid_ids <- function(df, var_madre, cols_dummies = NULL, col_compact = NULL) {

      if (!is.null(col_compact) && !is.na(col_compact) && col_compact %in% names(df)) {
        x <- as.character(df[[col_compact]])
        ok <- !is.na(x) & nzchar(x) & x != "NA"
        return(which(ok))
      }

      if (!is.null(cols_dummies) && length(cols_dummies)) {
        cols_dummies <- cols_dummies[cols_dummies %in% names(df)]
        if (!length(cols_dummies)) return(integer(0))

        mat <- sapply(cols_dummies, function(cc) {
          v <- suppressWarnings(as.numeric(as.character(df[[cc]])))
          v %in% c(0, 1)
        })
        if (!is.matrix(mat)) mat <- matrix(mat, ncol = 1)

        ok <- rowSums(mat, na.rm = TRUE) > 0
        return(which(ok))
      }

      integer(0)
    }

    .sm_numerador_option <- function(df, var_madre, code, cols_dummies = NULL, col_compact = NULL) {

      if (!is.null(col_compact) && !is.na(col_compact) && col_compact %in% names(df)) {
        x <- as.character(df[[col_compact]])
        ok <- !is.na(x) & nzchar(x) & x != "NA"
        if (!any(ok)) return(integer(0))
        vals <- strsplit(x[ok], "\\s*;\\s*")
        ids_ok <- which(ok)
        hit <- vapply(vals, function(v) any(trimws(v) == code), logical(1))
        return(ids_ok[hit])
      }

      if (!is.null(cols_dummies) && length(cols_dummies)) {
        col <- paste0(var_madre, ".", code)
        if (!col %in% names(df)) return(integer(0))
        v <- suppressWarnings(as.numeric(as.character(df[[col]])))
        return(which(!is.na(v) & v == 1))
      }

      integer(0)
    }

    # -------------------------------------------------------------------------
    # Paleta SO
    # -------------------------------------------------------------------------
    .resolver_paleta_var <- function(var, instrumento, colores_apiladas_por_listname, opcion_levels) {

      surv <- instrumento$survey %||% NULL
      pal  <- NULL

      if (!is.null(colores_apiladas_por_listname) &&
          !is.null(surv) &&
          all(c("name", "list_name") %in% names(surv))) {

        ln <- .get_list_name_safe(surv, var)
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

      if (!is.null(instrumento$choices) &&
          all(c("list_name", "name") %in% names(instrumento$choices)) &&
          !is.null(names(pal))) {

        fila <- surv[surv$name == var, , drop = FALSE]
        list_var <- if (nrow(fila)) fila$list_name[1] else NA_character_

        label_col <- if ("label" %in% names(instrumento$choices)) {
          "label"
        } else {
          cand <- grep("^label(::|$)", names(instrumento$choices), value = TRUE)
          if (length(cand)) cand[1] else NULL
        }

        if (!is.na(list_var) && nzchar(list_var) &&
            !is.null(label_col) && label_col %in% names(instrumento$choices)) {

          ch <- instrumento$choices[instrumento$choices$list_name == list_var, , drop = FALSE]
          map_code_to_label <- stats::setNames(
            as.character(ch[[label_col]]),
            as.character(ch$name)
          )

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
      }

      pal2 <- rep(pal, length.out = length(opcion_levels))
      names(pal2) <- opcion_levels
      pal2
    }

    # -------------------------------------------------------------------------
    # Plot SO x SO
    # -------------------------------------------------------------------------
    .plot_so_so <- function(df, var_main, var_cruce) {

      survey <- instrumento$survey %||% NULL

      cats_main <- get_categorias_so(var_main, df, survey, orders_list %||% instrumento$orders_list %||% NULL)
      codes_row <- as.character(cats_main$codes)
      opciones  <- as.character(cats_main$labels)

      if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0 && length(codes_row)) {
        codp <- as.character(codigos_perdidos)
        keep <- !(codes_row %in% codp)
        codes_row <- codes_row[keep]
        opciones  <- opciones[keep]
      }

      cats_cruce <- get_categorias_so(var_cruce, df, survey, orders_list %||% instrumento$orders_list %||% NULL)
      estr_codes  <- as.character(cats_cruce$codes)
      estr_labels <- as.character(cats_cruce$labels)

      v_main  <- as.character(df[[var_main]])
      v_cruce <- as.character(df[[var_cruce]])
      w <- get_pesos(df, weight_col)

      rows <- list()

      for (j in seq_along(estr_codes)) {
        key_j  <- estr_codes[j]
        mask_j <- !is.na(v_cruce) & v_cruce == key_j

        elig <- mask_j & !is.na(v_main) & nzchar(v_main) & v_main != "NA" & (v_main %in% codes_row)
        N_j  <- sum(w[elig], na.rm = TRUE)

        if (is.na(N_j) || N_j <= 0) next

        for (i in seq_along(codes_row)) {
          code_i <- codes_row[i]
          n_ij <- sum(w[elig & v_main == code_i], na.rm = TRUE)
          rows[[length(rows) + 1]] <- data.frame(
            estrato_label = .wrap_titulo_html(estr_labels[j], width = 50),
            opcion_label  = opciones[i],
            pct = n_ij / N_j,
            n   = n_ij,
            stringsAsFactors = FALSE
          )
        }
      }

      df_tab <- dplyr::bind_rows(rows)
      if (!nrow(df_tab)) {
        return(
          plotly::plot_ly() |>
            plotly::layout(
              annotations = list(list(
                text = "Sin datos para graficar.",
                showarrow = FALSE,
                font = list(color = COLOR_TEXTO_SUAVE)
              )),
              paper_bgcolor = COLOR_SUPERFICIE,
              plot_bgcolor  = COLOR_SUPERFICIE
            ) |>
            plotly::config(displayModeBar = FALSE, responsive = TRUE)
        )
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

        col_opt <- pal[[opt]] %||% unname(pal[opt]) %||% "#9aa4b2"

        p <- p |>
          plotly::add_bars(
            data             = dfo,
            x                = ~pct,
            y                = ~estrato_label,
            name             = opt,
            orientation      = "h",
            text             = ~paste0("<b>", round(100 * pct, 0), "%</b>"),
            textposition     = "inside",
            insidetextanchor = "middle",
            textfont         = list(color = "white", size = 11),
            customdata       = ~hover,
            hovertemplate    = "%{customdata}<extra></extra>",
            marker           = list(color = col_opt, line = list(width = 0)),
            showlegend       = FALSE
          )
      }

      p |>
        plotly::layout(
          barmode = "stack",
          bargap  = 0.25,
          xaxis   = list(
            title = "",
            range = c(0, 1),
            showgrid = FALSE,
            zeroline = FALSE,
            showticklabels = FALSE,
            ticks = ""
          ),
          yaxis   = list(
            title = "",
            automargin = TRUE,
            showgrid = FALSE,
            zeroline = FALSE,
            ticks = "",
            tickpadding = 10,
            tickfont = list(color = COLOR_TEXTO)
          ),
          margin  = list(l = 50, r = 25, t = 10, b = 25),
          hovermode = "closest",
          transition = list(duration = 450, easing = "cubic-in-out"),
          showlegend = FALSE,
          paper_bgcolor = COLOR_SUPERFICIE,
          plot_bgcolor  = COLOR_SUPERFICIE
        ) |>
        plotly::config(displayModeBar = FALSE, responsive = TRUE)
    }

    # -------------------------------------------------------------------------
    # Plot SM por opción x estratos
    # -------------------------------------------------------------------------
    .plot_sm_option_chip <- function(df, var_madre, code, opt_label, var_cruce,
                                     cols_dummies = NULL, col_compact = NULL) {

      survey <- instrumento$survey %||% NULL
      cats_cruce <- get_categorias_so(var_cruce, df, survey, orders_list %||% instrumento$orders_list %||% NULL)
      estr_codes  <- as.character(cats_cruce$codes)
      estr_labels <- as.character(cats_cruce$labels)

      v_cruce <- as.character(df[[var_cruce]])
      w <- get_pesos(df, weight_col)

      rows <- list()

      for (j in seq_along(estr_codes)) {
        key_j  <- estr_codes[j]
        mask_j <- !is.na(v_cruce) & v_cruce == key_j

        ids_valid_sm <- .sm_valid_ids(df, var_madre, cols_dummies = cols_dummies, col_compact = col_compact)
        if (!length(ids_valid_sm)) next

        ids_mask <- which(mask_j)
        ids_denom <- intersect(ids_mask, ids_valid_sm)
        if (!length(ids_denom)) next

        N_j <- sum(w[ids_denom], na.rm = TRUE)
        if (is.na(N_j) || N_j <= 0) next

        ids_yes <- .sm_numerador_option(df, var_madre, code, cols_dummies = cols_dummies, col_compact = col_compact)
        ids_yes <- intersect(ids_yes, ids_denom)

        n_yes <- sum(w[ids_yes], na.rm = TRUE)
        pct_y <- if (N_j > 0) n_yes / N_j else 0
        pct_y <- max(0, min(1, pct_y))

        rows[[length(rows) + 1]] <- data.frame(
          estrato_label = .wrap_titulo_html(estr_labels[j], width = 50),
          pct_yes = pct_y,
          n_yes = n_yes,
          N = N_j,
          stringsAsFactors = FALSE
        )
      }

      dfi <- dplyr::bind_rows(rows)

      if (!nrow(dfi)) {
        return(
          plotly::plot_ly(height = BAR_HEIGHT) |>
            plotly::layout(
              annotations = list(list(
                text = "Sin datos válidos.",
                showarrow = FALSE,
                font = list(color = COLOR_TEXTO_SUAVE)
              )),
              xaxis = list(visible = FALSE),
              yaxis = list(visible = FALSE),
              margin = list(l = 10, r = 10, t = 0, b = 0),
              paper_bgcolor = COLOR_SUPERFICIE,
              plot_bgcolor  = COLOR_SUPERFICIE
            ) |>
            plotly::config(displayModeBar = FALSE, responsive = TRUE)
        )
      }

      dfi$estrato_label <- factor(dfi$estrato_label, levels = unique(dfi$estrato_label))
      dfi$pct_bg <- 1 - dfi$pct_yes

      dfi$txt <- paste0("<b>", round(100 * dfi$pct_yes, 0), "%</b>")
      dfi$hover <- sprintf(
        "%s<br>%s: %s%%<br>n: %s<br>N: %s",
        as.character(dfi$estrato_label),
        opt_label,
        round(100 * dfi$pct_yes, 1),
        format(round(dfi$n_yes, 0), big.mark = ","),
        format(round(dfi$N, 0), big.mark = ",")
      )

      p <- plotly::plot_ly(height = max(220, 70 + 32 * nrow(dfi)))

      p <- p |>
        plotly::add_bars(
          data             = dfi,
          x                = ~pct_yes,
          y                = ~estrato_label,
          name             = "Sí",
          orientation      = "h",
          text             = ~txt,
          textposition     = "inside",
          insidetextanchor = "middle",
          textfont         = list(color = "white", size = PCT_FSIZE),
          customdata       = ~hover,
          hovertemplate    = "%{customdata}<extra></extra>",
          marker           = list(color = SM_COLOR_YES, line = list(width = 0))
        )

      p <- p |>
        plotly::add_bars(
          data        = dfi,
          x           = ~pct_bg,
          y           = ~estrato_label,
          name        = " ",
          orientation = "h",
          hoverinfo   = "skip",
          marker      = list(color = SM_COLOR_BG, line = list(width = 0)),
          showlegend  = FALSE
        )

      p |>
        plotly::layout(
          barmode = "stack",
          xaxis = list(
            range = c(0, 1),
            showgrid = FALSE,
            zeroline = FALSE,
            showticklabels = FALSE,
            ticks = "",
            title = ""
          ),
          yaxis = list(
            title = "",
            automargin = TRUE,
            showgrid = FALSE,
            zeroline = FALSE,
            tickfont = list(color = COLOR_TEXTO)
          ),
          margin = list(l = 120, r = 15, t = 8, b = 10),
          showlegend = FALSE,
          uniformtext = list(minsize = 10, mode = "hide"),
          paper_bgcolor = COLOR_SUPERFICIE,
          plot_bgcolor  = COLOR_SUPERFICIE
        ) |>
        plotly::config(displayModeBar = FALSE, responsive = TRUE)
    }

    # -------------------------------------------------------------------------
    # Tabla
    # -------------------------------------------------------------------------
    .build_cuerpo <- function(df, var_main, var_cruce) {

      survey <- instrumento$survey %||% NULL
      tp_main <- tipo_pregunta(var_main, survey = survey, sm_vars_force = vars_sm_madres, df = df)

      cats_cruce <- get_categorias_so(var_cruce, df, survey, orders_list %||% instrumento$orders_list %||% NULL)
      estr_codes  <- as.character(cats_cruce$codes)
      estr_labels <- as.character(cats_cruce$labels)

      v_cruce <- as.character(df[[var_cruce]])
      w <- get_pesos(df, weight_col)

      if (tp_main == "so") {

        cats_main <- get_categorias_so(var_main, df, survey, orders_list %||% instrumento$orders_list %||% NULL)
        codes_row <- as.character(cats_main$codes)
        labels_row <- as.character(cats_main$labels)

        if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0 && length(codes_row)) {
          codp <- as.character(codigos_perdidos)
          keep <- !(codes_row %in% codp)
          codes_row <- codes_row[keep]
          labels_row <- labels_row[keep]
        }

        cuerpo <- tibble::tibble(Opciones = labels_row)
        denom_map <- list()

        v_main <- as.character(df[[var_main]])
        elig_total <- !is.na(v_main) & nzchar(v_main) & v_main != "NA" & (v_main %in% codes_row)
        N_total <- sum(w[elig_total], na.rm = TRUE)

        n_total <- vapply(seq_along(codes_row), function(i) sum(w[elig_total & v_main == codes_row[i]], na.rm = TRUE), numeric(1))
        pct_total <- if (N_total > 0) n_total / N_total else rep(0, length(n_total))

        cuerpo <- dplyr::bind_cols(cuerpo, tibble::tibble(Total__n = n_total, Total__pct = pct_total))
        denom_map[["Total__n"]] <- N_total

        for (j in seq_along(estr_codes)) {
          key_j <- estr_codes[j]
          mask_j <- !is.na(v_cruce) & v_cruce == key_j

          elig <- mask_j & elig_total
          N_j <- sum(w[elig], na.rm = TRUE)

          if (is.na(N_j) || N_j <= 0) {
            n_vec <- rep(0, length(codes_row))
            pct   <- rep(0, length(codes_row))
          } else {
            n_vec <- vapply(seq_along(codes_row), function(i) sum(w[elig & v_main == codes_row[i]], na.rm = TRUE), numeric(1))
            pct   <- n_vec / N_j
          }

          nm_n   <- paste0(var_cruce, "__", make.names(estr_labels[j]), "__n")
          nm_pct <- paste0(var_cruce, "__", make.names(estr_labels[j]), "__pct")

          cuerpo <- dplyr::bind_cols(cuerpo, tibble::tibble(!!nm_n := n_vec, !!nm_pct := pct))
          denom_map[[nm_n]] <- N_j
        }

      } else {

        spec <- .resolver_var_spec_safe(var_main, df)
        cols <- spec$cols %||% character(0)

        codes_row <- sub(paste0("^", var_main, "\\."), "", cols)
        codes_row <- codes_row[nzchar(codes_row)]

        map <- spec$map_code_to_label %||% list()
        labels_row <- vapply(codes_row, function(cd) as.character(map[[cd]] %||% cd), character(1))

        if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0 && length(codes_row)) {
          codp <- as.character(codigos_perdidos)
          keep <- !(codes_row %in% codp)
          codes_row  <- codes_row[keep]
          labels_row <- labels_row[keep]
          cols       <- cols[sub(paste0("^", var_main, "\\."), "", cols) %in% codes_row]
        }

        cuerpo <- tibble::tibble(Opciones = labels_row)
        denom_map <- list()

        ids_valid <- .sm_valid_ids(df, var_main, cols_dummies = cols, col_compact = NA_character_)
        N_total <- if (length(ids_valid)) sum(w[ids_valid], na.rm = TRUE) else 0

        n_total <- vapply(seq_along(codes_row), function(i) {
          ids_yes <- .sm_numerador_option(df, var_main, codes_row[i], cols_dummies = cols, col_compact = NA_character_)
          ids_yes <- intersect(ids_yes, ids_valid)
          sum(w[ids_yes], na.rm = TRUE)
        }, numeric(1))

        pct_total <- if (N_total > 0) n_total / N_total else rep(0, length(n_total))

        cuerpo <- dplyr::bind_cols(cuerpo, tibble::tibble(Total__n = n_total, Total__pct = pct_total))
        denom_map[["Total__n"]] <- N_total

        for (j in seq_along(estr_codes)) {
          key_j <- estr_codes[j]
          mask_ids <- which(!is.na(v_cruce) & v_cruce == key_j)

          ids_denom <- intersect(mask_ids, ids_valid)
          N_j <- if (length(ids_denom)) sum(w[ids_denom], na.rm = TRUE) else 0

          if (is.na(N_j) || N_j <= 0) {
            n_vec <- rep(0, length(codes_row))
            pct   <- rep(0, length(codes_row))
          } else {
            n_vec <- vapply(seq_along(codes_row), function(i) {
              ids_yes <- .sm_numerador_option(df, var_main, codes_row[i], cols_dummies = cols, col_compact = NA_character_)
              ids_yes <- intersect(ids_yes, ids_denom)
              sum(w[ids_yes], na.rm = TRUE)
            }, numeric(1))
            pct <- n_vec / N_j
          }

          nm_n   <- paste0(var_cruce, "__", make.names(estr_labels[j]), "__n")
          nm_pct <- paste0(var_cruce, "__", make.names(estr_labels[j]), "__pct")

          cuerpo <- dplyr::bind_cols(cuerpo, tibble::tibble(!!nm_n := n_vec, !!nm_pct := pct))
          denom_map[[nm_n]] <- N_j
        }
      }

      total_row <- as.list(rep(NA, ncol(cuerpo)))
      names(total_row) <- names(cuerpo)
      total_row[["Opciones"]] <- "Total"

      n_cols   <- grep("__n$",   names(cuerpo))
      pct_cols <- grep("__pct$", names(cuerpo))

      for (k in n_cols) {
        nm <- names(cuerpo)[k]
        Nj <- denom_map[[nm]]
        total_row[[k]] <- if (is.null(Nj) || is.na(Nj)) NA_real_ else round(as.numeric(Nj), 0)
      }
      for (k in pct_cols) {
        nm_pct <- names(cuerpo)[k]
        nm_n   <- sub("__pct$", "__n", nm_pct)
        Nj <- denom_map[[nm_n]]
        total_row[[k]] <- if (is.null(Nj) || is.na(Nj) || Nj <= 0) 0 else 1
      }

      cuerpo <- dplyr::bind_rows(cuerpo, tibble::as_tibble(total_row))

      dic_vars <- NULL
      surv <- instrumento$survey %||% NULL
      if (!is.null(surv) && "name" %in% names(surv)) {
        label_col <- if ("label" %in% names(surv)) {
          "label"
        } else {
          cand <- grep("^label(::|$)", names(surv), value = TRUE)
          if (length(cand)) cand[1] else NULL
        }
        if (!is.null(label_col) && label_col %in% names(surv)) {
          dic_vars <- dplyr::transmute(surv, name = .data$name, label = .data[[label_col]])
        }
      }

      label_variable <- function(var, dic_vars = NULL, labels_override = NULL, df = NULL) {
        if (!is.null(labels_override) && var %in% names(labels_override)) return(as.character(labels_override[[var]]))
        if (!is.null(df) && var %in% names(df)) {
          vlab <- attr(df[[var]], "label", exact = TRUE)
          if (!is.null(vlab) && nzchar(as.character(vlab))) return(as.character(vlab))
        }
        if (!is.null(dic_vars) && all(c("name","label") %in% names(dic_vars))) {
          lab <- dic_vars$label[dic_vars$name == var]
          if (length(lab) && !all(is.na(lab))) return(as.character(lab[1]))
        }
        as.character(var)
      }

      cruce_lbl <- label_variable(var_cruce, dic_vars = dic_vars, labels_override = labels_override, df = df)

      list(
        cuerpo       = cuerpo,
        tipo_main    = tp_main,
        estr_labels  = estr_labels,
        cruce_lbl    = cruce_lbl
      )
    }

    # -------------------------------------------------------------------------
    # Encabezado DT
    # -------------------------------------------------------------------------
    .dt_container_multihdr <- function(cuerpo, cruce_lbl, estr_labels) {

      n_blocks <- 1L + length(estr_labels)
      ncols    <- ncol(cuerpo)
      exp_cols <- 1L + 2L * n_blocks

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

      fila2 <- c(
        list(htmltools::tags$th(colspan = 2, "Total")),
        lapply(estr_labels, function(lab) htmltools::tags$th(colspan = 2, as.character(lab)))
      )

      fila3 <- unlist(
        replicate(n_blocks, list(htmltools::tags$th("n"), htmltools::tags$th("%")), simplify = FALSE),
        recursive = FALSE
      )

      htmltools::withTags(
        table(
          class = "display nowrap compact",
          thead(
            tr(
              htmltools::tags$th(rowspan = 3, ""),
              htmltools::tags$th(colspan = ncols - 1, cruce_lbl)
            ),
            tr(fila2),
            tr(fila3)
          )
        )
      )
    }

    # -------------------------------------------------------------------------
    # Wiring UI
    # -------------------------------------------------------------------------
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
        shiny::updateSelectInput(session, "main_seccion", choices = stats::setNames(secs, secs), selected = secs[1])
        shiny::updateSelectInput(session, "cruce_seccion", choices = stats::setNames(secs, secs), selected = secs[1])
      }
    })

    shiny::observeEvent(input$main_seccion, {
      sec <- input$main_seccion
      if (is.null(sec) || !nzchar(sec) || is.null(secciones_limpias[[sec]])) return()

      vars_sec <- secciones_limpias[[sec]]
      pool_main <- unique(c(vars_so, vars_sm_madres))

      main_choices <- unique(vars_sec[vars_sec %in% pool_main])
      if (!length(main_choices)) main_choices <- pool_main

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
      cruce_choices <- unique(vars_sec[vars_sec %in% vars_so])
      if (!length(cruce_choices)) cruce_choices <- vars_so

      cruce_lab <- stats::setNames(
        cruce_choices,
        vapply(cruce_choices, function(v) .obtener_label_var(v, instrumento, data), character(1))
      )

      shiny::updateSelectInput(session, "cruce_var", choices = cruce_lab, selected = cruce_choices[1] %||% "")
    }, ignoreInit = TRUE)

    # -------------------------------------------------------------------------
    # Header gráfico
    # -------------------------------------------------------------------------
    output$rel_plot_header <- shiny::renderUI({
      shiny::req(input$main_var, input$cruce_var)
      t_main  <- .wrap_titulo_html(.obtener_label_var(input$main_var, instrumento, data), width = 110)
      t_cruce <- .obtener_label_var(input$cruce_var, instrumento, data)

      shiny::tagList(
        shiny::div(class = "cardbox-title", shiny::HTML(t_main)),
        shiny::div(
          class = "cardbox-subtitle",
          style = paste0("color:", COLOR_TEXTO_SUAVE, ";"),
          paste0("Cruce: ", t_cruce)
        )
      )
    })

    # -------------------------------------------------------------------------
    # Reactivo central
    # -------------------------------------------------------------------------
    rel_obj <- shiny::reactive({
      shiny::req(input$main_var, input$cruce_var)

      var_main  <- input$main_var
      var_cruce <- input$cruce_var

      if (!(var_cruce %in% vars_so)) {
        return(list(error = "No es posible cruzar con la selección actual."))
      }

      df <- data
      if (var_cruce %in% names(df)) df <- df[!is.na(df[[var_cruce]]), , drop = FALSE]
      if (!nrow(df)) return(list(error = "Sin datos disponibles."))

      survey <- instrumento$survey %||% NULL
      tp_main <- tipo_pregunta(var_main, survey = survey, sm_vars_force = vars_sm_madres, df = df)

      out_tab <- .build_cuerpo(df, var_main, var_cruce)

      list(
        df          = df,
        var_main    = var_main,
        var_cruce   = var_cruce,
        tipo_main   = tp_main,
        cuerpo      = out_tab$cuerpo,
        cruce_lbl   = out_tab$cruce_lbl,
        estr_labels = out_tab$estr_labels,
        error       = NULL
      )
    })

    # -------------------------------------------------------------------------
    # UI dinámica del gráfico
    # -------------------------------------------------------------------------
    output$rel_plot_ui <- shiny::renderUI({
      obj <- rel_obj()
      if (!is.null(obj$error)) {
        return(shiny::div(style = paste0("padding:12px;color:", COLOR_TEXTO_SUAVE, ";"), obj$error))
      }

      if (identical(obj$tipo_main, "so")) {
        return(plotly::plotlyOutput(session$ns("rel_plot"), height = "520px"))
      }

      shiny::div(
        style = "display:flex; flex-direction:column; gap:12px;",
        shiny::uiOutput(session$ns("rel_sm_chips_ui"))
      )
    })

    # -------------------------------------------------------------------------
    # Leyenda SO externa
    # -------------------------------------------------------------------------
    output$rel_so_legend <- shiny::renderUI({

      obj <- rel_obj()
      if (!is.null(obj$error)) return(NULL)
      if (!identical(obj$tipo_main, "so")) return(NULL)

      df <- obj$df
      var_main <- obj$var_main
      survey <- instrumento$survey %||% NULL

      cats_main <- get_categorias_so(
        var = var_main,
        df = df,
        survey = survey,
        orders_list = orders_list %||% instrumento$orders_list %||% NULL
      )

      codes_row <- as.character(cats_main$codes)
      opciones  <- as.character(cats_main$labels)

      if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0 && length(codes_row)) {
        codp <- as.character(codigos_perdidos)
        keep <- !(codes_row %in% codp)
        opciones <- opciones[keep]
      }

      pal <- .resolver_paleta_var(
        var = var_main,
        instrumento = instrumento,
        colores_apiladas_por_listname = colores_apiladas_por_listname,
        opcion_levels = unique(opciones)
      )

      legend_levels <- opciones

      shiny::div(
        class = "rel-legend",
        style = paste0(
          "margin-top:14px;",
          "padding:12px 14px;",
          "border:1px solid ", COLOR_BORDE, ";",
          "border-radius:14px;",
          "background:", COLOR_SUPERFICIE, ";"
        ),
        shiny::div(
          style = "display:flex; justify-content:center; width:100%;",
          shiny::div(
            style = "display:flex; flex-wrap:wrap; justify-content:center; gap:12px 18px; align-items:center; max-width:980px; width:100%;",
            lapply(legend_levels, function(lab) {
              col <- pal[[lab]] %||% unname(pal[lab]) %||% "#9aa4b2"

              shiny::div(
                style = "display:flex; align-items:center; gap:10px; max-width:360px;",
                shiny::span(style = paste0(
                  "display:inline-block;",
                  "width:16px; height:16px;",
                  "border-radius:4px;",
                  "background:", col, ";",
                  "box-shadow:0 0 0 1px rgba(0,0,0,0.06) inset;"
                )),
                shiny::span(
                  style = paste0(
                    "font-size:14px;",
                    "font-weight:500;",
                    "color:", COLOR_TEXTO, ";",
                    "line-height:1.15;",
                    "white-space:normal;",
                    "word-break:break-word;"
                  ),
                  lab
                )
              )
            })
          )
        )
      )
    })

    # -------------------------------------------------------------------------
    # UI chips SM
    # -------------------------------------------------------------------------
    output$rel_sm_chips_ui <- shiny::renderUI({
      obj <- rel_obj()
      if (!is.null(obj$error)) return(NULL)
      if (!identical(obj$tipo_main, "sm")) return(NULL)

      df <- obj$df
      var_main <- obj$var_main

      spec <- .resolver_var_spec_safe(var_main, df)
      cols <- spec$cols %||% character(0)
      if (!length(cols)) {
        return(shiny::div(style = paste0("padding:12px;color:", COLOR_TEXTO_SUAVE, ";"), "SM sin dummies disponibles."))
      }

      codes <- sub(paste0("^", var_main, "\\."), "", cols)
      codes <- codes[nzchar(codes)]

      if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0) {
        codp <- as.character(codigos_perdidos)
        keep <- !(codes %in% codp)
        codes <- codes[keep]
      }

      if (!length(codes)) {
        return(shiny::div(style = paste0("padding:12px;color:", COLOR_TEXTO_SUAVE, ";"), "SM sin opciones graficables."))
      }

      if (length(codes) > MAX_SM_CHIPS) {
        return(shiny::div(
          style = paste0("padding:12px;color:", COLOR_TEXTO_SUAVE, ";"),
          "Variable con demasiadas opciones para graficar en chips. (Ver tabla)"
        ))
      }

      map <- spec$map_code_to_label %||% list()

      shiny::div(
        style = "display:flex; flex-direction:column; gap:12px;",
        lapply(seq_along(codes), function(i) {
          code_i <- codes[i]
          lab_i  <- as.character(map[[code_i]] %||% code_i)
          out_id <- paste0("rel_sm_plot_", i)

          shiny::div(
            style = paste0(
              "border:1px solid ", COLOR_BORDE, ";",
              "border-radius:14px;",
              "padding:10px 12px;",
              "background:", COLOR_SUPERFICIE, ";"
            ),
            shiny::div(
              style = paste0(
                "font-size:12px;",
                "font-weight:400;",
                "color:", COLOR_TEXTO, ";",
                "margin:0 0 8px 0;"
              ),
              lab_i
            ),
            plotly::plotlyOutput(session$ns(out_id), height = "260px")
          )
        })
      )
    })

    # -------------------------------------------------------------------------
    # Render SO
    # -------------------------------------------------------------------------
    output$rel_plot <- plotly::renderPlotly({
      obj <- rel_obj()
      if (!is.null(obj$error)) {
        return(
          plotly::plot_ly() |>
            plotly::layout(
              annotations = list(list(
                text = obj$error,
                showarrow = FALSE,
                font = list(color = COLOR_TEXTO_SUAVE)
              )),
              paper_bgcolor = COLOR_SUPERFICIE,
              plot_bgcolor  = COLOR_SUPERFICIE
            ) |>
            plotly::config(displayModeBar = FALSE, responsive = TRUE)
        )
      }

      if (!identical(obj$tipo_main, "so")) return(NULL)
      .plot_so_so(obj$df, obj$var_main, obj$var_cruce)
    })

    # -------------------------------------------------------------------------
    # Render SM chips
    # -------------------------------------------------------------------------
    shiny::observe({
      obj <- rel_obj()
      if (!is.null(obj$error)) return()
      if (!identical(obj$tipo_main, "sm")) return()

      df <- obj$df
      var_main  <- obj$var_main
      var_cruce <- obj$var_cruce

      spec <- .resolver_var_spec_safe(var_main, df)
      cols <- spec$cols %||% character(0)
      if (!length(cols)) return()

      codes <- sub(paste0("^", var_main, "\\."), "", cols)
      codes <- codes[nzchar(codes)]

      if (!is.null(codigos_perdidos) && length(codigos_perdidos) > 0) {
        codp <- as.character(codigos_perdidos)
        keep <- !(codes %in% codp)
        codes <- codes[keep]
      }

      if (!length(codes)) return()
      if (length(codes) > MAX_SM_CHIPS) return()

      map <- spec$map_code_to_label %||% list()

      for (i in seq_along(codes)) {
        local({
          ii <- i
          code_i <- codes[ii]
          lab_i  <- as.character(map[[code_i]] %||% code_i)
          out_id <- paste0("rel_sm_plot_", ii)

          output[[out_id]] <- plotly::renderPlotly({
            .plot_sm_option_chip(
              df = df,
              var_madre = var_main,
              code = code_i,
              opt_label = lab_i,
              var_cruce = var_cruce,
              cols_dummies = cols,
              col_compact  = NA_character_
            )
          })
        })
      }
    })

    # -------------------------------------------------------------------------
    # Tabla DT
    # -------------------------------------------------------------------------
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
            columnDefs = list(list(className = "dt-center", targets = "_all"))
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
      is_n   <- grepl("__n$", names(cuerpo))

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

    invisible(NULL)
  })
}
