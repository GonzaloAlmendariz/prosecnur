# =============================================================================
# Tab 4: Dimensiones (UI + server)
# - Vista General / Indicadores
# - Motor visual automático (Radar -> Barras cuando no aplica)
# - Heatmap semafórico + cruce principal + filtros categóricos
# =============================================================================
#' @keywords internal
#' @noRd

.ui_tab_dimensiones <- function(ctx) {
  `%||%` <- get0("%||%", ifnotfound = function(x, y) if (!is.null(x)) x else y)

  cfg <- ctx$dimensiones$config %||% list()
  vis <- cfg$visual %||% list()
  show_total_default <- isTRUE(vis$incluir_total_default)

  shiny::sidebarLayout(
    shiny::sidebarPanel(
      width = 3,
      class = "sidebar-panel-base",
      shiny::div(
        class = "sidebar-stack",
        shiny::div(
          class = "sidebar-module sidebar-module-rel",
          shiny::h3(class = "sidebar-module-title", "Dimensiones"),
          shiny::p(
            class = "sidebar-module-help",
            "Explora resultados de forma simple: elige una vista, compara por grupos y aplica filtros."
          ),

          shiny::div(
            class = "sidebar-module-card",
            shiny::div(class = "sidebar-subtitle", "Vista"),
            shiny::p(
              class = "rel-sidebar-hint",
              shiny::HTML(
                "<strong>General:</strong> resume el indicador por sus componentes. &nbsp;&nbsp; <strong>Indicadores:</strong> abre el detalle por preguntas."
              )
            ),
            shiny::div(
              class = "toggle-row dim-vista-switch-row",
              shiny::span(class = "toggle-label dim-vista-label", "General"),
              shiny::tags$label(
                class = "switch",
                shiny::tags$input(id = "dim_vista_indicadores", type = "checkbox"),
                shiny::tags$span(class = "slider")
              ),
              shiny::span(class = "toggle-label dim-vista-label", "Indicadores")
            ),
            shiny::selectizeInput(
              inputId = "dim_objetivo",
              label = "Selecciona un indicador",
              choices = c(),
              options = list(dropdownParent = "body")
            ),
            shiny::uiOutput("dim_objetivo_help_ui")
          ),
          shiny::div(
            class = "sidebar-module-card rel-sidebar-card-gap",
            shiny::div(class = "sidebar-subtitle", "Comparación"),
            shiny::selectizeInput(
              inputId = "dim_principal_seccion",
              label = "Sección",
              choices = c(),
              selected = "",
              options = list(dropdownParent = "body")
            ),
            shiny::selectizeInput(
              inputId = "dim_principal_var",
              label = "Comparar por",
              choices = c("Sin cruce" = ""),
              selected = "",
              options = list(dropdownParent = "body")
            ),
            shiny::div(
              class = "toggle-row",
              shiny::span(class = "toggle-label", "Incluir total"),
              shiny::tags$label(
                class = "switch",
                if (isTRUE(show_total_default)) {
                  shiny::tags$input(id = "dim_show_total", type = "checkbox", checked = "checked")
                } else {
                  shiny::tags$input(id = "dim_show_total", type = "checkbox")
                },
                shiny::tags$span(class = "slider")
              )
            ),
            shiny::p(
              class = "rel-sidebar-hint",
              "Usa el cruce para comparar el indicador entre categorías (ejemplo: servicio o distrito)."
            )
          ),

          shiny::div(
            class = "sidebar-module-card rel-sidebar-card-gap",
            shiny::div(class = "sidebar-subtitle", "Filtrar datos"),
            shiny::selectizeInput(
              inputId = "dim_filtro_seccion",
              label = "Sección",
              choices = c(),
              selected = "",
              options = list(dropdownParent = "body")
            ),
            shiny::selectizeInput(
              inputId = "dim_filtro_var",
              label = "Variable",
              choices = c(),
              selected = "",
              options = list(dropdownParent = "body")
            ),
            shiny::uiOutput("dim_filtro_categorias_ui"),
            shiny::actionButton(
              inputId = "dim_limpiar_filtros",
              label = "Limpiar filtros",
              class = "sidebar-quick-btn"
            )
          )
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
            shiny::div(
              class = "cardbox-header",
              shiny::div(class = "cardbox-title", "Mapa de calor"),
              shiny::uiOutput("dim_heatmap_subtitle_ui")
            ),
            shiny::div(class = "dim-plot-wrap", shiny::uiOutput("dim_heatmap_ui")),
            shiny::uiOutput("dim_heatmap_legend_ui")
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
              class = "cardbox-header rel-plot-header",
              shiny::div(
                class = "rel-plot-header-main",
                shiny::uiOutput("dim_main_title_ui"),
                shiny::uiOutput("dim_main_subtitle_ui")
              ),
              shiny::div(
                class = "rel-plot-header-actions",
                shiny::uiOutput("dim_focus_controls_ui")
              )
            ),
            shiny::div(class = "dim-plot-wrap", shiny::uiOutput("dim_main_plot_ui"))
          )
        )
      ),
      shiny::div(style = "height: 48px;")
    )
  )
}

#' @keywords internal
#' @noRd
.server_tab_dimensiones <- function(ctx, input, output, session) {
  `%||%` <- get0("%||%", ifnotfound = function(x, y) if (!is.null(x)) x else y)

  dim_ctx <- ctx$dimensiones %||% NULL
  if (is.null(dim_ctx) || !isTRUE(dim_ctx$habilitado)) return(invisible(NULL))

  data_dim <- dim_ctx$data
  instrumento <- ctx$instrumento
  cfg <- dim_ctx$config %||% list()

  if (!is.data.frame(data_dim) || !nrow(data_dim)) return(invisible(NULL))

  idx_meta <- attr(data_dim, "indices_meta", exact = TRUE)
  rec_meta <- attr(data_dim, "recodificacion_items_meta", exact = TRUE)

  meta_indices <- if (is.list(idx_meta) && is.list(idx_meta$indices)) idx_meta$indices else list()
  meta_bloques <- if (is.list(idx_meta) && is.list(idx_meta$bloques)) idx_meta$bloques else list()

  idx_key_to_var <- stats::setNames(
    vapply(meta_indices, function(x) as.character(x$salida %||% NA_character_)[1], character(1)),
    names(meta_indices)
  )
  idx_key_to_var <- idx_key_to_var[!is.na(idx_key_to_var) & nzchar(idx_key_to_var)]
  idx_var_to_key <- stats::setNames(names(idx_key_to_var), as.character(idx_key_to_var))

  bloq_key_to_var <- stats::setNames(
    vapply(meta_bloques, function(x) as.character(x$salida %||% NA_character_)[1], character(1)),
    names(meta_bloques)
  )
  bloq_key_to_var <- bloq_key_to_var[!is.na(bloq_key_to_var) & nzchar(bloq_key_to_var)]
  bloq_var_to_key <- stats::setNames(names(bloq_key_to_var), as.character(bloq_key_to_var))

  rec_out_to_src <- stats::setNames(character(0), character(0))
  if (is.list(rec_meta) && length(rec_meta)) {
    rec_df <- data.frame(
      src = names(rec_meta),
      out = vapply(rec_meta, function(x) as.character(x$variable_salida %||% NA_character_)[1], character(1)),
      stringsAsFactors = FALSE
    )
    rec_df <- rec_df[!is.na(rec_df$out) & nzchar(rec_df$out), , drop = FALSE]
    if (nrow(rec_df)) {
      rec_out_to_src <- stats::setNames(as.character(rec_df$src), as.character(rec_df$out))
    }
  }

  .as_named_chr <- function(x) {
    if (is.null(x)) return(stats::setNames(character(0), character(0)))
    v <- as.character(unlist(x, use.names = TRUE))
    n <- names(v)
    if (is.null(n)) return(stats::setNames(character(0), character(0)))
    ok <- !is.na(n) & nzchar(trimws(n)) & !is.na(v) & nzchar(trimws(v))
    stats::setNames(v[ok], n[ok])
  }

  .nm_get <- function(x, key) {
    key <- as.character(key %||% "")[1]
    if (!nzchar(key)) return(NULL)
    nms <- names(x)
    if (is.null(nms)) return(NULL)
    i <- match(key, nms)
    if (is.na(i)) return(NULL)
    as.character(x[i])[1]
  }

  .round_half_up <- function(x, digits = 0L) {
    s <- 10^as.integer(digits)
    out <- ifelse(
      is.na(x),
      NA_real_,
      ifelse(x >= 0, floor(x * s + 0.5), ceiling(x * s - 0.5)) / s
    )
    as.numeric(out)
  }

  .fmt_int <- function(x) {
    x <- .round_half_up(x, 0)
    ifelse(is.na(x), "", format(as.integer(x), trim = TRUE, scientific = FALSE))
  }

  .clamp <- function(x, lo, hi) max(lo, min(hi, x))

  .pretty_label <- function(x) {
    x <- as.character(x %||% "")
    x <- gsub("^idx_", "", x)
    x <- gsub("^bloq_", "", x)
    x <- gsub("^r100_", "", x)
    x <- gsub("[_\\.]+", " ", x)
    x <- trimws(x)
    if (!nzchar(x)) return("Variable")
    paste0(toupper(substring(x, 1, 1)), substring(x, 2))
  }

  .first_nonempty <- function(...) {
    vals <- list(...)
    for (vv in vals) {
      v <- as.character(vv %||% "")[1]
      if (!is.na(v) && nzchar(trimws(v))) return(trimws(v))
    }
    ""
  }

  .label_var <- function(v) {
    f <- get0(".obtener_label_var", mode = "function", ifnotfound = NULL)
    if (is.function(f)) return(as.character(f(v, instrumento, data_dim)))
    as.character(v)
  }

  .label_data <- function(v) {
    if (!(v %in% names(data_dim))) return(.pretty_label(v))
    lb <- attr(data_dim[[v]], "label", exact = TRUE)
    lb <- as.character(lb %||% "")
    lb <- gsub("\\s*\\[0-100\\]$", "", lb)
    if (nzchar(trimws(lb))) trimws(lb) else .pretty_label(v)
  }

  lbl_idx <- .as_named_chr(cfg$labels_indices)
  lbl_bloq <- .as_named_chr(cfg$labels_bloques)
  lbl_ind <- .as_named_chr(cfg$labels_indicadores)

  .label_idx <- function(v, key = NULL) {
    kk <- as.character(key %||% .nm_get(idx_var_to_key, v) %||% "")
    .first_nonempty(
      .nm_get(lbl_idx, kk),
      .nm_get(lbl_idx, v),
      if (nzchar(kk)) .pretty_label(kk) else "",
      .label_data(v),
      .label_var(v),
      .pretty_label(v)
    )
  }

  .label_bloq <- function(v, key = NULL) {
    kk <- as.character(key %||% .nm_get(bloq_var_to_key, v) %||% "")
    .first_nonempty(
      .nm_get(lbl_bloq, kk),
      .nm_get(lbl_bloq, v),
      if (nzchar(kk)) .pretty_label(kk) else "",
      .label_data(v),
      .label_var(v),
      .pretty_label(v)
    )
  }

  .label_ind <- function(v) {
    src <- as.character(.nm_get(rec_out_to_src, v) %||% "")
    .first_nonempty(
      .nm_get(lbl_ind, v),
      if (nzchar(src)) .nm_get(lbl_ind, src) else "",
      .label_data(v),
      .label_var(v),
      if (nzchar(src)) .pretty_label(src) else "",
      .pretty_label(v)
    )
  }

  .wrap_axis_label <- function(x, width = 16L) {
    x <- as.character(x %||% "")
    if (!length(x)) return(x)
    if (requireNamespace("stringr", quietly = TRUE)) {
      out <- stringr::str_wrap(x, width = width)
    } else {
      out <- vapply(x, function(xx) paste(strwrap(xx, width = width), collapse = "\n"), character(1))
    }
    gsub("\n", "<br>", out, fixed = TRUE)
  }

  .add_alpha <- function(col, alpha = 0.22) {
    grDevices::adjustcolor(as.character(col %||% "#1F4E85"), alpha.f = alpha)
  }

  .palette_ipe <- function(n) {
    base_cols <- c(
      "#355C7D", "#6C5B7B", "#C06C84", "#F67280", "#F8B195",
      "#4575B4", "#74ADD1", "#ABD9E9", "#E0F3F8", "#FEE090",
      "#FDAE61", "#F46D43", "#D73027", "#66BD63", "#1A9850",
      "#006837", "#8C510A", "#BF812D", "#DFC27D", "#80CDC1",
      "#018571", "#35978F", "#A6CEE3", "#1F78B4", "#B2DF8A", "#33A02C"
    )
    if (n <= length(base_cols)) base_cols[seq_len(n)] else grDevices::colorRampPalette(base_cols)(n)
  }

  .palette_okabe <- function(n) {
    cols <- c("#0072B2", "#E69F00", "#009E73", "#D55E00", "#CC79A7", "#56B4E9", "#F0E442", "#000000")
    if (n <= length(cols)) cols[seq_len(n)] else grDevices::colorRampPalette(cols)(n)
  }

  vis_cfg <- cfg$visual %||% list()
  radar_min_ejes <- suppressWarnings(as.integer(vis_cfg$radar_min_ejes %||% 3L)[1])
  if (!is.finite(radar_min_ejes) || is.na(radar_min_ejes) || radar_min_ejes < 1L) radar_min_ejes <- 3L

  max_categorias_principal <- suppressWarnings(as.integer(vis_cfg$max_categorias_principal %||% 8L)[1])
  if (!is.finite(max_categorias_principal) || is.na(max_categorias_principal) || max_categorias_principal < 1L) {
    max_categorias_principal <- 8L
  }

  paleta_radar <- as.character(vis_cfg$paleta_radar %||% "okabe_ito")[1]
  if (!paleta_radar %in% c("okabe_ito", "ipe")) paleta_radar <- "okabe_ito"

  sem_cfg <- cfg$semaforo %||% list()
  sem_cortes <- suppressWarnings(as.numeric(sem_cfg$cortes %||% c(50, 75)))
  sem_cortes <- sem_cortes[is.finite(sem_cortes)]
  if (length(sem_cortes) < 2L) sem_cortes <- c(50, 75)
  sem_cortes <- sort(unique(sem_cortes))[1:2]
  sem_cortes <- pmax(0, pmin(100, sem_cortes))
  if (length(sem_cortes) < 2L || sem_cortes[1] >= sem_cortes[2]) sem_cortes <- c(50, 75)

  sem_cols <- as.character(sem_cfg$colores %||% character(0))
  nms_sem <- names(sem_cols %||% character(0))
  if (is.null(nms_sem)) nms_sem <- character(0)
  sem_col_rojo <- if ("rojo" %in% nms_sem) sem_cols[["rojo"]] else "#D84B55"
  sem_col_amb <- if ("ambar" %in% nms_sem) sem_cols[["ambar"]] else "#E0B44C"
  sem_col_ver <- if ("verde" %in% nms_sem) sem_cols[["verde"]] else "#3A9A5B"
  sem_col_na <- "#DFE5EE"

  .range_labels <- function(c1, c2) {
    c(
      paste0("Menor a ", .fmt_int(c1)),
      paste0(.fmt_int(c1), " - ", .fmt_int(c2 - 1)),
      paste0("Mayor a ", .fmt_int(c2 - 1))
    )
  }

  .group_colors <- function(groups) {
    groups <- unique(as.character(groups))
    if (!length(groups)) return(stats::setNames(character(0), character(0)))

    total_color <- as.character(ctx$theme_app$color_primario %||% "#0E3B74")
    others <- setdiff(groups, "Total")
    pal <- if (identical(paleta_radar, "ipe")) .palette_ipe(length(others)) else .palette_okabe(length(others))
    names(pal) <- others

    out <- stats::setNames(rep("#4B6E99", length(groups)), groups)
    if ("Total" %in% groups) out[["Total"]] <- total_color
    if (length(others)) out[others] <- pal[others]
    out
  }

  weight_col <- as.character(dim_ctx$weight_col %||% "")[1]
  if (!nzchar(weight_col) || !(weight_col %in% names(data_dim))) {
    weight_col <- if ("peso" %in% names(data_dim)) "peso" else ""
  }

  .safe_weights <- function(df) {
    if (!nzchar(weight_col) || !(weight_col %in% names(df))) return(rep(1, nrow(df)))
    w <- suppressWarnings(as.numeric(df[[weight_col]]))
    w[!is.finite(w) | is.na(w)] <- 0
    w
  }

  .weighted_mean <- function(x, w) {
    x <- suppressWarnings(as.numeric(x))
    w <- suppressWarnings(as.numeric(w))
    ok <- is.finite(x) & !is.na(x) & is.finite(w) & !is.na(w) & w > 0
    if (!any(ok)) return(NA_real_)
    sum(x[ok] * w[ok], na.rm = TRUE) / sum(w[ok], na.rm = TRUE)
  }

  .choices_label_col <- function(ch) {
    if (is.null(ch)) return(NULL)
    if ("label" %in% names(ch)) return("label")
    cand <- grep("^label(::|$)", names(ch), value = TRUE)
    if (length(cand)) cand[1] else NULL
  }

  .choice_map <- function(var) {
    surv <- instrumento$survey %||% NULL
    ch <- instrumento$choices %||% NULL
    if (is.null(surv) || is.null(ch) ||
        !all(c("name", "list_name") %in% names(surv)) ||
        !all(c("list_name", "name") %in% names(ch))) {
      return(stats::setNames(character(0), character(0)))
    }

    ln <- as.character(surv$list_name[surv$name == var][1])
    if (is.na(ln) || !nzchar(ln)) return(stats::setNames(character(0), character(0)))

    col_lab <- .choices_label_col(ch)
    if (is.null(col_lab) || !(col_lab %in% names(ch))) return(stats::setNames(character(0), character(0)))

    chv <- ch[ch$list_name == ln, , drop = FALSE]
    if (!nrow(chv)) return(stats::setNames(character(0), character(0)))
    stats::setNames(as.character(chv[[col_lab]]), as.character(chv$name))
  }

  .level_label_map <- function(v) {
    if (!(v %in% names(data_dim))) return(stats::setNames(character(0), character(0)))

    out <- stats::setNames(character(0), character(0))
    labs <- attr(data_dim[[v]], "labels", exact = TRUE)
    if (!is.null(labs) && length(labs)) {
      out <- stats::setNames(as.character(unname(labs)), as.character(names(labs)))
    }

    map_choice <- .choice_map(v)
    if (length(map_choice)) out[names(map_choice)] <- map_choice
    out
  }

  .categorias_var <- function(df, var, w, max_levels = 12L) {
    out_empty <- list(
      rows = data.frame(value = character(0), label = character(0), base = numeric(0), stringsAsFactors = FALSE),
      total_levels = 0L,
      hidden_levels = 0L
    )

    if (!(var %in% names(df)) || !nrow(df)) return(out_empty)

    x <- trimws(as.character(df[[var]]))
    ok <- !is.na(x) & nzchar(x) & x != "NA"
    if (!any(ok)) return(out_empty)

    ww <- as.numeric(w)
    if (length(ww) != nrow(df)) ww <- rep(1, nrow(df))

    tab <- stats::aggregate(
      ww[ok],
      by = list(value = x[ok]),
      FUN = sum,
      na.rm = TRUE
    )
    names(tab) <- c("value", "base")
    tab <- tab[order(-tab$base, tab$value), , drop = FALSE]

    map <- .level_label_map(var)
    labs <- unname(map[tab$value])
    labs[is.na(labs) | !nzchar(labs)] <- tab$value[is.na(labs) | !nzchar(labs)]
    tab$label <- as.character(labs)

    n_tot <- nrow(tab)
    if (is.finite(max_levels) && max_levels > 0L && n_tot > max_levels) {
      tab <- tab[seq_len(max_levels), , drop = FALSE]
    }

    list(
      rows = tab[, c("value", "label", "base"), drop = FALSE],
      total_levels = n_tot,
      hidden_levels = max(0L, n_tot - nrow(tab))
    )
  }

  .build_catalog <- function(cat_in, mode = c("general", "indicadores")) {
    mode <- match.arg(mode)
    out <- list()

    if (is.list(cat_in) && length(cat_in)) {
      for (nm in names(cat_in)) {
        it <- cat_in[[nm]]
        if (!is.list(it)) next

        if (identical(mode, "general")) {
          id_var <- as.character(it$id %||% nm)[1]
          key <- as.character(it$key %||% .nm_get(idx_var_to_key, id_var) %||% nm)[1]
          axis_vars <- as.character(it$axis_vars %||% character(0))
          axis_vars <- axis_vars[axis_vars %in% names(data_dim)]
          if (!length(axis_vars) || !(id_var %in% names(data_dim))) next

          out[[id_var]] <- list(
            id = id_var,
            key = key,
            mode = "general",
            label = .label_idx(id_var, key),
            axis_vars = axis_vars,
            axis_labels = vapply(axis_vars, .label_bloq, character(1))
          )
        } else {
          key <- as.character(it$key %||% it$id %||% nm)[1]
          bvar <- as.character(it$block_var %||% .nm_get(bloq_key_to_var, key) %||% NA_character_)[1]
          axis_vars <- as.character(it$axis_vars %||% character(0))
          axis_vars <- axis_vars[axis_vars %in% names(data_dim)]
          if (!length(axis_vars)) next

          out[[key]] <- list(
            id = key,
            key = key,
            mode = "indicadores",
            label = .label_bloq(bvar, key),
            block_var = bvar,
            axis_vars = axis_vars,
            axis_labels = vapply(axis_vars, .label_ind, character(1))
          )
        }
      }
    }

    out
  }

  catalog_general <- .build_catalog(cfg$catalog_general, mode = "general")
  catalog_indicadores <- .build_catalog(cfg$catalog_indicadores, mode = "indicadores")

  if (!length(catalog_general)) {
    for (nm in names(meta_indices)) {
      it <- meta_indices[[nm]]
      idx_var <- as.character(it$salida %||% NA_character_)[1]
      if (is.na(idx_var) || !nzchar(idx_var) || !(idx_var %in% names(data_dim))) next

      refs <- unique(c(
        as.character(it$refs_resueltas %||% character(0)),
        as.character(it$refs %||% character(0))
      ))
      axis_vars <- character(0)
      for (r in refs) {
        rv <- if (r %in% names(data_dim)) {
          r
        } else if (r %in% names(bloq_key_to_var)) {
          as.character(bloq_key_to_var[[r]])
        } else {
          NA_character_
        }
        if (!is.na(rv) && nzchar(rv) && rv %in% names(data_dim) && !(rv %in% axis_vars)) {
          axis_vars <- c(axis_vars, rv)
        }
      }
      if (!length(axis_vars)) next

      catalog_general[[idx_var]] <- list(
        id = idx_var,
        key = nm,
        mode = "general",
        label = .label_idx(idx_var, nm),
        axis_vars = axis_vars,
        axis_labels = vapply(axis_vars, .label_bloq, character(1))
      )
    }
  }

  if (!length(catalog_indicadores)) {
    for (bk in names(meta_bloques)) {
      bl <- meta_bloques[[bk]]
      bvar <- as.character(bl$salida %||% NA_character_)[1]
      axis_vars <- unique(as.character(bl$vars %||% character(0)))
      axis_vars <- axis_vars[axis_vars %in% names(data_dim)]
      if (!length(axis_vars)) next

      catalog_indicadores[[bk]] <- list(
        id = bk,
        key = bk,
        mode = "indicadores",
        label = .label_bloq(bvar, bk),
        block_var = bvar,
        axis_vars = axis_vars,
        axis_labels = vapply(axis_vars, .label_ind, character(1))
      )
    }
  }

  surv <- instrumento$survey %||% NULL
  so_all <- character(0)
  if (!is.null(surv) && all(c("name", "type") %in% names(surv))) {
    so_all <- as.character(surv$name[grepl("^select_one\\b", tolower(as.character(surv$type)))])
    so_all <- unique(so_all[so_all %in% names(data_dim)])
  }

  sec_map_raw <- ctx$secciones_limpias %||% list()
  if (!is.list(sec_map_raw) || !length(sec_map_raw)) {
    sec_map_raw <- list("Variables disponibles" = so_all)
  }
  if (is.null(names(sec_map_raw)) || !length(names(sec_map_raw))) {
    names(sec_map_raw) <- paste0("Sección ", seq_along(sec_map_raw))
  }

  section_var_map <- lapply(sec_map_raw, function(vs) {
    vv <- unique(as.character(vs))
    vv <- vv[vv %in% names(data_dim)]
    vv
  })
  section_var_map <- section_var_map[vapply(section_var_map, length, integer(1)) > 0]

  if (!length(section_var_map)) {
    fallback_vars <- if (length(so_all)) so_all else character(0)
    if (length(fallback_vars)) {
      section_var_map <- list("Variables disponibles" = fallback_vars)
    }
  }

  sec_names <- names(section_var_map)
  principal_sec_choices <- c("Todas las secciones" = "__all__", stats::setNames(sec_names, sec_names))
  filtro_sec_choices <- c("Todas las secciones" = "__all__", stats::setNames(sec_names, sec_names))

  .vars_for_section <- function(sec) {
    s <- as.character(sec %||% "__all__")[1]
    if (identical(s, "__all__")) {
      vv <- unique(unlist(section_var_map, use.names = FALSE))
    } else {
      vv <- section_var_map[[s]] %||% character(0)
    }
    vv <- vv[vv %in% names(data_dim)]
    labs <- vapply(vv, .label_var, character(1))
    list(vars = vv, labels = labs)
  }

  .choices_for_section <- function(sec, empty_label = "Sin selección") {
    out <- .vars_for_section(sec)
    c(stats::setNames("", empty_label), stats::setNames(out$vars, out$labels))
  }

  modes_disponibles <- shiny::reactive({
    modes <- character(0)
    if (length(catalog_general)) modes <- c(modes, "general")
    if (length(catalog_indicadores)) modes <- c(modes, "indicadores")
    unique(modes)
  })

  mode_activo <- shiny::reactive({
    modes <- modes_disponibles()
    if (!length(modes)) return("general")
    m <- if (isTRUE(input$dim_vista_indicadores)) "indicadores" else "general"
    if (m %in% modes) m else modes[1]
  })

  shiny::observe({
    modes <- modes_disponibles()
    if (!length(modes)) return()
    if (length(modes) == 1L) {
      shiny::updateCheckboxInput(
        session,
        "dim_vista_indicadores",
        value = identical(modes[1], "indicadores")
      )
    }
  })

  shiny::observe({
    mode <- mode_activo()
    obj_map <- if (identical(mode, "indicadores")) catalog_indicadores else catalog_general
    if (!length(obj_map)) {
      shiny::updateSelectizeInput(session, "dim_objetivo", choices = c(), selected = character(0), server = FALSE)
      return()
    }

    ids <- names(obj_map)
    labs <- vapply(obj_map, function(x) as.character(x$label %||% x$id), character(1))
    choices <- stats::setNames(ids, labs)

    cur <- as.character(input$dim_objetivo %||% "")[1]
    sel <- if (cur %in% ids) cur else ids[1]
    lbl <- if (identical(mode, "indicadores")) "Bloque a analizar" else "Índice a analizar"

    shiny::updateSelectizeInput(session, "dim_objetivo", label = lbl, choices = choices, selected = sel, server = FALSE)
  })

  output$dim_objetivo_help_ui <- shiny::renderUI({
    mode <- mode_activo()
    txt <- if (identical(mode, "indicadores")) {
      "Muestra cada bloque por sus preguntas recodificadas (aperturas)."
    } else {
      "Muestra el índice elegido y sus componentes para comparar brechas entre grupos."
    }
    shiny::p(class = "rel-sidebar-hint", txt)
  })

  shiny::observe({
    psec_cur <- as.character(input$dim_principal_seccion %||% "__all__")[1]
    psec_sel <- if (psec_cur %in% c("__all__", sec_names)) psec_cur else "__all__"
    shiny::updateSelectizeInput(
      session, "dim_principal_seccion",
      choices = principal_sec_choices,
      selected = psec_sel,
      server = FALSE
    )

    pvars <- .vars_for_section(psec_sel)$vars
    pcur <- as.character(input$dim_principal_var %||% "")[1]
    psel <- if (pcur %in% c("", pvars)) pcur else ""
    shiny::updateSelectizeInput(
      session, "dim_principal_var",
      choices = .choices_for_section(psec_sel, empty_label = "Sin cruce"),
      selected = psel,
      server = FALSE
    )

    fsec_cur <- as.character(input$dim_filtro_seccion %||% "__all__")[1]
    fsec_sel <- if (fsec_cur %in% c("__all__", sec_names)) fsec_cur else "__all__"
    shiny::updateSelectizeInput(
      session, "dim_filtro_seccion",
      choices = filtro_sec_choices,
      selected = fsec_sel,
      server = FALSE
    )

    fvars <- .vars_for_section(fsec_sel)$vars
    fcur <- as.character(input$dim_filtro_var %||% "")[1]
    fsel <- if (fcur %in% c("", fvars)) fcur else ""
    shiny::updateSelectizeInput(
      session, "dim_filtro_var",
      choices = .choices_for_section(fsec_sel, empty_label = "Sin filtro"),
      selected = fsel,
      server = FALSE
    )
  })

  shiny::observeEvent(input$dim_filtro_var, {
    fv <- as.character(input$dim_filtro_var %||% "")[1]
    if (!nzchar(fv) || !(fv %in% names(data_dim))) {
      shiny::updateCheckboxGroupInput(session, "dim_filtro_categorias", choices = character(0), selected = character(0))
      return()
    }

    w0 <- .safe_weights(data_dim)
    cats <- .categorias_var(data_dim, fv, w0, max_levels = 40L)
    if (!nrow(cats$rows)) {
      shiny::updateCheckboxGroupInput(session, "dim_filtro_categorias", choices = character(0), selected = character(0))
      return()
    }

    choices <- stats::setNames(cats$rows$value, cats$rows$label)
    shiny::updateCheckboxGroupInput(session, "dim_filtro_categorias", choices = choices, selected = cats$rows$value)
  }, ignoreInit = FALSE)

  output$dim_filtro_categorias_ui <- shiny::renderUI({
    fv <- as.character(input$dim_filtro_var %||% "")[1]
    if (!nzchar(fv) || !(fv %in% names(data_dim))) return(NULL)
    shiny::checkboxGroupInput(
      inputId = "dim_filtro_categorias",
      label = "Categorías",
      choices = character(0),
      selected = character(0)
    )
  })

  shiny::observeEvent(input$dim_limpiar_filtros, {
    shiny::updateSelectizeInput(session, "dim_filtro_seccion", selected = "__all__")
    shiny::updateSelectizeInput(session, "dim_filtro_var", selected = "")
  })

  data_filtrada <- shiny::reactive({
    df <- data_dim
    fv <- as.character(input$dim_filtro_var %||% "")[1]
    if (nzchar(fv) && fv %in% names(df)) {
      if (is.null(input$dim_filtro_categorias)) return(df)

      cats <- as.character(input$dim_filtro_categorias %||% character(0))
      if (!length(cats)) return(df[0, , drop = FALSE])

      xv <- trimws(as.character(df[[fv]]))
      keep <- !is.na(xv) & xv %in% cats
      df <- df[keep, , drop = FALSE]
    }
    df
  })

  objetivo_activo <- shiny::reactive({
    mode <- mode_activo()
    obj_map <- if (identical(mode, "indicadores")) catalog_indicadores else catalog_general
    if (!length(obj_map)) return(NULL)

    id <- as.character(input$dim_objetivo %||% "")[1]
    if (!nzchar(id) || !(id %in% names(obj_map))) id <- names(obj_map)[1]
    obj_map[[id]]
  })

  score_payload <- shiny::reactive({
    df <- data_filtrada()
    obj <- objetivo_activo()

    if (!nrow(df) || is.null(obj) || !length(obj$axis_vars)) {
      return(list(
        score_plot = data.frame(),
        score_heat = data.frame(),
        axis_order_plot = character(0),
        axis_order_heat = character(0),
        mode = NA_character_,
        objective = NA_character_,
        principal_label = "",
        principal_var = "",
        principal_hidden = 0L
      ))
    }

    axis_vars <- as.character(obj$axis_vars)
    axis_vars <- axis_vars[axis_vars %in% names(df)]
    if (!length(axis_vars)) {
      return(list(
        score_plot = data.frame(),
        score_heat = data.frame(),
        axis_order_plot = character(0),
        axis_order_heat = character(0),
        mode = as.character(obj$mode %||% ""),
        objective = as.character(obj$label %||% obj$id),
        principal_label = "",
        principal_var = "",
        principal_hidden = 0L
      ))
    }

    axis_labels <- as.character(obj$axis_labels %||% axis_vars)
    axis_labels <- axis_labels[match(axis_vars, as.character(obj$axis_vars))]

    w <- .safe_weights(df)
    pv <- as.character(input$dim_principal_var %||% "")[1]
    include_total <- isTRUE(input$dim_show_total)

    groups <- list()
    hidden_main <- 0L

    if (include_total) {
      groups[[length(groups) + 1L]] <- list(
        label = "Total",
        mask = rep(TRUE, nrow(df)),
        base = sum(w, na.rm = TRUE)
      )
    }

    if (nzchar(pv) && pv %in% names(df)) {
      cats <- .categorias_var(df, pv, w, max_levels = max_categorias_principal)
      hidden_main <- as.integer(cats$hidden_levels)
      xv <- trimws(as.character(df[[pv]]))

      for (i in seq_len(nrow(cats$rows))) {
        val <- as.character(cats$rows$value[i])
        groups[[length(groups) + 1L]] <- list(
          label = as.character(cats$rows$label[i]),
          mask = !is.na(xv) & xv == val,
          base = as.numeric(cats$rows$base[i])
        )
      }
    }

    if (!length(groups)) {
      groups[[1]] <- list(
        label = "Total",
        mask = rep(TRUE, nrow(df)),
        base = sum(w, na.rm = TRUE)
      )
    }

    obj_mode <- as.character(obj$mode %||% "")
    obj_summary_var <- if (identical(obj_mode, "general")) {
      as.character(obj$id %||% "")
    } else {
      as.character(obj$block_var %||% "")
    }
    if (!nzchar(obj_summary_var) || !(obj_summary_var %in% names(df))) {
      obj_summary_var <- ""
    }

    out_axis <- list()
    out_total <- list()
    for (g in groups) {
      gw <- w * as.numeric(g$mask)
      for (j in seq_along(axis_vars)) {
        v <- axis_vars[j]
        mu <- .weighted_mean(df[[v]], gw)
        out_axis[[length(out_axis) + 1L]] <- data.frame(
          axis_var = v,
          axis_label = as.character(axis_labels[j]),
          grupo = as.character(g$label),
          tipo = "apertura",
          score_raw = as.numeric(mu),
          base = as.numeric(g$base),
          stringsAsFactors = FALSE
        )
      }

      mu_total <- if (nzchar(obj_summary_var)) {
        .weighted_mean(df[[obj_summary_var]], gw)
      } else {
        X <- as.data.frame(df[, axis_vars, drop = FALSE])
        X[] <- lapply(X, function(z) suppressWarnings(as.numeric(z)))
        row_mu <- rowMeans(X, na.rm = TRUE)
        row_mu[!is.finite(row_mu)] <- NA_real_
        .weighted_mean(row_mu, gw)
      }

      out_total[[length(out_total) + 1L]] <- data.frame(
        axis_var = "__total_cruce__",
        axis_label = "Total cruce",
        grupo = as.character(g$label),
        tipo = "total_cruce",
        score_raw = as.numeric(mu_total),
        base = as.numeric(g$base),
        stringsAsFactors = FALSE
      )
    }

    sc_plot <- dplyr::bind_rows(out_axis)
    sc_total <- dplyr::bind_rows(out_total)
    sc_plot$score_round <- .round_half_up(sc_plot$score_raw, 0)
    sc_total$score_round <- .round_half_up(sc_total$score_raw, 0)
    sc_heat <- dplyr::bind_rows(sc_total, sc_plot)

    list(
      score_plot = sc_plot,
      score_heat = sc_heat,
      axis_order_plot = as.character(axis_labels),
      axis_order_heat = c("Total cruce", as.character(axis_labels)),
      mode = as.character(obj$mode %||% ""),
      objective = as.character(obj$label %||% obj$id),
      principal_label = .label_var(pv),
      principal_var = pv,
      principal_hidden = hidden_main
    )
  })

  .group_order <- function(sc) {
    if (!nrow(sc)) return(character(0))
    base_df <- sc |>
      dplyr::distinct(.data$grupo, .data$base)

    others_df <- base_df[base_df$grupo != "Total", , drop = FALSE]
    others <- as.character(others_df$grupo[order(-others_df$base, as.character(others_df$grupo))])

    if ("Total" %in% as.character(base_df$grupo)) {
      unique(c("Total", others))
    } else {
      unique(others)
    }
  }

  visual_mode_resolved <- shiny::reactive({
    p <- score_payload()
    n_ejes <- length(unique(as.character(p$axis_order_plot %||% character(0))))
    if (n_ejes >= radar_min_ejes) "radar" else "barras"
  })

  group_levels <- shiny::reactive({
    sc <- score_payload()$score_plot
    .group_order(sc)
  })

  focus_group <- shiny::reactiveVal("")

  shiny::observe({
    lv <- group_levels()
    if (!length(lv)) {
      focus_group("")
      return()
    }
    cur <- as.character(focus_group() %||% "")[1]
    if (!nzchar(cur) || !(cur %in% lv)) focus_group(lv[1])
  })

  shiny::observeEvent(input$dim_focus_next, {
    lv <- group_levels()
    if (length(lv) <= 1L) return()
    cur <- as.character(focus_group() %||% "")[1]
    idx <- which(lv == cur)[1]
    if (is.na(idx)) idx <- 1L
    nxt <- if (idx >= length(lv)) 1L else idx + 1L
    focus_group(lv[nxt])
  })

  focus_enabled <- shiny::reactive({
    isTRUE(input$dim_focus_enable) && length(group_levels()) > 1L
  })

  output$dim_focus_controls_ui <- shiny::renderUI({
    lv <- group_levels()
    if (length(lv) <= 1L) return(NULL)

    sc <- score_payload()$score_plot
    gf <- as.character(focus_group() %||% lv[1])[1]
    b <- sc$base[match(gf, as.character(sc$grupo))]
    b <- suppressWarnings(as.numeric(b[1]))

    shiny::div(
      class = "dim-focus-wrap",
      shiny::div(
        class = "toggle-row dim-focus-toggle",
        shiny::span(class = "toggle-label", "Comparar"),
        shiny::tags$label(
          class = "switch",
          if (isTRUE(input$dim_focus_enable)) {
            shiny::tags$input(id = "dim_focus_enable", type = "checkbox", checked = "checked")
          } else {
            shiny::tags$input(id = "dim_focus_enable", type = "checkbox")
          },
          shiny::tags$span(class = "slider")
        ),
        shiny::span(class = "toggle-label", "Enfoque")
      ),
      if (isTRUE(input$dim_focus_enable)) {
        shiny::div(
          class = "rel-iter-level-control",
          shiny::actionButton(
            inputId = "dim_focus_next",
            label = NULL,
            icon = shiny::icon("repeat"),
            class = "rel-iter-circle-btn",
            title = "Siguiente grupo"
          ),
          shiny::div(
            class = "rel-iter-level-chip",
            shiny::div(class = "rel-iter-level-name", gf),
            shiny::div(
              class = "rel-iter-level-meta",
              paste0("N ", format(round(b, 0), big.mark = ",", scientific = FALSE))
            )
          )
        )
      }
    )
  })

  output$dim_main_title_ui <- shiny::renderUI({
    ttl <- "Comparación del indicador"
    shiny::div(class = "cardbox-title", ttl)
  })

  output$dim_heatmap_subtitle_ui <- shiny::renderUI({
    p <- score_payload()
    if (!nrow(p$score_heat)) {
      return(shiny::div(class = "cardbox-subtitle", "Sin datos disponibles con la selección actual."))
    }

    sec_mode <- if (identical(p$mode, "indicadores")) "Vista Indicadores" else "Vista General"
    principal_txt <- if (nzchar(p$principal_var)) {
      paste0("Cruce: ", p$principal_label, ". ")
    } else {
      "Sin cruce. "
    }

    cuts_lab <- .range_labels(sem_cortes[1], sem_cortes[2])

    shiny::div(
      class = "cardbox-subtitle",
      paste0(
        sec_mode, " | Objetivo: ", p$objective, " | ",
        principal_txt,
        "Rangos: ", cuts_lab[1], ", ", cuts_lab[2], ", ", cuts_lab[3], ".",
        if (isTRUE(p$principal_hidden > 0L)) paste0(" +", p$principal_hidden, " categorías no visibles por legibilidad.") else ""
      )
    )
  })

  output$dim_main_subtitle_ui <- shiny::renderUI({
    p <- score_payload()
    if (!nrow(p$score_plot)) {
      return(shiny::div(class = "cardbox-subtitle", "Sin datos disponibles con la selección actual."))
    }

    shiny::div(
      class = "cardbox-subtitle",
      paste0(
        if (identical(p$mode, "indicadores")) "Vista indicadores" else "Vista general",
        " | Objetivo: ", p$objective,
        " | Total: ", if (isTRUE(input$dim_show_total)) "incluido" else "oculto"
      )
    )
  })

  output$dim_heatmap_ui <- shiny::renderUI({
    plotly::plotlyOutput("dim_heatmap_plot", height = "460px")
  })

  output$dim_heatmap_legend_ui <- shiny::renderUI({
    p <- score_payload()
    if (!nrow(p$score_heat)) return(NULL)

    cuts_lab <- .range_labels(sem_cortes[1], sem_cortes[2])
    shiny::div(
      class = "dim-heat-legend",
      shiny::div(
        class = "dim-heat-legend-item",
        shiny::span(class = "dim-heat-legend-swatch", style = paste0("background:", sem_col_rojo, ";")),
        shiny::span(class = "dim-heat-legend-text", cuts_lab[1])
      ),
      shiny::div(
        class = "dim-heat-legend-item",
        shiny::span(class = "dim-heat-legend-swatch", style = paste0("background:", sem_col_amb, ";")),
        shiny::span(class = "dim-heat-legend-text", cuts_lab[2])
      ),
      shiny::div(
        class = "dim-heat-legend-item",
        shiny::span(class = "dim-heat-legend-swatch", style = paste0("background:", sem_col_ver, ";")),
        shiny::span(class = "dim-heat-legend-text", cuts_lab[3])
      )
    )
  })

  output$dim_main_plot_ui <- shiny::renderUI({
    h <- if (identical(visual_mode_resolved(), "radar")) "600px" else "560px"
    plotly::plotlyOutput("dim_main_plot_plot", height = h)
  })

  .heat_colorscale <- function() {
    list(
      list(0.000000, sem_col_na),
      list(0.249999, sem_col_na),
      list(0.250000, sem_col_rojo),
      list(0.499999, sem_col_rojo),
      list(0.500000, sem_col_amb),
      list(0.749999, sem_col_amb),
      list(0.750000, sem_col_ver),
      list(1.000000, sem_col_ver)
    )
  }

  output$dim_heatmap_plot <- plotly::renderPlotly({
    p <- score_payload()
    sc <- p$score_heat

    if (!nrow(sc)) {
      return(
        plotly::plot_ly(height = 460) |>
          plotly::layout(
            annotations = list(list(text = "Sin datos para mostrar", showarrow = FALSE)),
            margin = list(l = 10, r = 10, t = 10, b = 10),
            xaxis = list(visible = FALSE),
            yaxis = list(visible = FALSE)
          ) |>
          plotly::config(displayModeBar = FALSE, responsive = TRUE)
      )
    }

    grupos_ord <- .group_order(sc)
    if (!length(grupos_ord)) grupos_ord <- unique(as.character(sc$grupo))

    axis_order <- as.character(p$axis_order_heat %||% unique(sc$axis_label))
    axis_order <- axis_order[axis_order %in% unique(sc$axis_label)]
    if (!length(axis_order)) axis_order <- unique(as.character(sc$axis_label))

    sc$grupo <- factor(sc$grupo, levels = grupos_ord)
    sc$axis_label <- factor(sc$axis_label, levels = rev(axis_order))
    sc$cat_code <- dplyr::case_when(
      is.na(sc$score_raw) ~ 0,
      sc$score_raw < sem_cortes[1] ~ 1,
      sc$score_raw < sem_cortes[2] ~ 2,
      TRUE ~ 3
    )
    sc$texto <- ifelse(is.na(sc$score_raw), "", .fmt_int(sc$score_round))

    cuts_lab <- .range_labels(sem_cortes[1], sem_cortes[2])
    sc$estado <- dplyr::case_when(
      is.na(sc$score_raw) ~ "Sin dato",
      sc$score_raw < sem_cortes[1] ~ cuts_lab[1],
      sc$score_raw < sem_cortes[2] ~ cuts_lab[2],
      TRUE ~ cuts_lab[3]
    )

    max_chars <- max(nchar(as.character(axis_order), type = "width"), na.rm = TRUE)
    left_margin <- .clamp(36 + 7 * max_chars, 130, 320)

    plotly::plot_ly(
      data = sc,
      x = ~grupo,
      y = ~axis_label,
      z = ~cat_code,
      text = ~texto,
      type = "heatmap",
      texttemplate = "%{text}",
      textfont = list(size = 11, color = "#122842"),
      xgap = 2,
      ygap = 2,
      colorscale = .heat_colorscale(),
      zmin = 0,
      zmax = 3,
      showscale = FALSE,
      hovertemplate = paste0(
        "<b>%{y}</b><br>",
        "Grupo: %{x}<br>",
        "Score: %{customdata}<br>",
        "Rango: %{meta}<extra></extra>"
      ),
      customdata = ~ifelse(is.na(score_raw), "Sin dato", .fmt_int(score_round)),
      meta = ~estado
      ) |>
      plotly::layout(
        margin = list(l = left_margin, r = 26, t = 8, b = 70),
        xaxis = list(title = "", tickangle = -18, tickfont = list(size = 11, color = "#20324d")),
        yaxis = list(title = "", tickfont = list(size = 11, color = "#20324d")),
        legend = list(title = list(text = ""))
      ) |>
      plotly::config(displayModeBar = FALSE, responsive = TRUE)
  })

  output$dim_main_plot_plot <- plotly::renderPlotly({
    p <- score_payload()
    sc <- p$score_plot

    if (!nrow(sc)) {
      return(
        plotly::plot_ly(height = 560) |>
          plotly::layout(
            annotations = list(list(text = "Sin datos para mostrar", showarrow = FALSE)),
            margin = list(l = 10, r = 10, t = 10, b = 10),
            xaxis = list(visible = FALSE),
            yaxis = list(visible = FALSE)
          ) |>
          plotly::config(displayModeBar = FALSE, responsive = TRUE)
      )
    }

    mode_plot <- visual_mode_resolved()
    grupos_ord_all <- .group_order(sc)
    if (!length(grupos_ord_all)) grupos_ord_all <- unique(as.character(sc$grupo))
    grupos_ord <- grupos_ord_all

    axis_order <- as.character(p$axis_order_plot %||% unique(sc$axis_label))
    axis_order <- axis_order[axis_order %in% unique(sc$axis_label)]
    if (!length(axis_order)) axis_order <- unique(as.character(sc$axis_label))

    if (isTRUE(focus_enabled())) {
      gf <- as.character(focus_group() %||% grupos_ord[1])[1]
      grupos_ord <- intersect(gf, grupos_ord)
      if (!length(grupos_ord)) grupos_ord <- unique(as.character(sc$grupo))[1]
    }

    cols_group_all <- .group_colors(grupos_ord_all)
    cols_group <- cols_group_all[grupos_ord]

    if (identical(mode_plot, "barras")) {
      max_chars <- max(nchar(as.character(axis_order), type = "width"), na.rm = TRUE)
      left_margin <- .clamp(36 + 7 * max_chars, 130, 320)

      pbar <- plotly::plot_ly(type = "bar", orientation = "h")
      for (g in grupos_ord) {
        dfg <- sc[as.character(sc$grupo) == g, , drop = FALSE]
        dfg <- dfg[match(axis_order, as.character(dfg$axis_label)), , drop = FALSE]

        pbar <- pbar |>
          plotly::add_trace(
            x = as.numeric(dfg$score_round),
            y = axis_order,
            name = g,
            marker = list(color = as.character(cols_group[[g]] %||% "#1F4E85")),
            hovertemplate = paste0(
              "<b>", g, "</b><br>",
              "%{y}: %{x}<extra></extra>"
            )
          )
      }

      return(
        pbar |>
          plotly::layout(
            barmode = "group",
            margin = list(l = left_margin, r = 26, t = 18, b = 78),
            xaxis = list(title = "Score (0-100)", range = c(0, 100), tickfont = list(size = 11, color = "#20324d")),
            yaxis = list(
              title = "",
              autorange = "reversed",
              tickfont = list(size = 11, color = "#20324d"),
              categoryorder = "array",
              categoryarray = axis_order
            ),
            legend = list(
              orientation = "h",
              y = -0.18,
              x = 0.5,
              xanchor = "center",
              entrywidthmode = if (length(grupos_ord_all) >= 5) "fraction" else NULL,
              entrywidth = if (length(grupos_ord_all) >= 5) 0.18 else NULL,
              title = list(text = "")
            )
          ) |>
          plotly::config(displayModeBar = FALSE, responsive = TRUE)
      )
    }

    theta_wrap <- .wrap_axis_label(axis_order, width = 16L)

    prad <- plotly::plot_ly(type = "scatterpolar", mode = "lines+markers")
    for (g in grupos_ord) {
      dfg <- sc[as.character(sc$grupo) == g, , drop = FALSE]
      dfg <- dfg[match(axis_order, as.character(dfg$axis_label)), , drop = FALSE]
      vals <- as.numeric(dfg$score_round)
      vals[!is.finite(vals)] <- NA_real_

      vals_poly <- c(vals, vals[1])
      theta_poly <- c(theta_wrap, theta_wrap[1])
      col_line <- as.character(cols_group[[g]] %||% "#1F4E85")
      is_total <- identical(g, "Total")

      prad <- prad |>
        plotly::add_trace(
          r = vals_poly,
          theta = theta_poly,
          name = g,
          fill = "toself",
          fillcolor = .add_alpha(col_line, if (is_total) 0.30 else 0.18),
          line = list(width = if (is_total) 3 else 2, color = col_line),
          marker = list(size = if (is_total) 6.8 else 5.2, color = col_line),
          hovertemplate = paste0(
            "<b>", g, "</b><br>",
            "%{theta}: %{r}<extra></extra>"
          )
        )
    }

    prad |>
      plotly::layout(
        polar = list(
          bgcolor = "rgba(255,255,255,0)",
          radialaxis = list(
            range = c(0, 100),
            tickmode = "array",
            tickvals = c(20, 40, 60, 80, 100),
            showticklabels = FALSE,
            ticks = "",
            gridcolor = "rgba(0,36,87,0.16)",
            linecolor = "rgba(0,36,87,0.16)"
          ),
          angularaxis = list(
            tickfont = list(size = 11, color = "#243a56"),
            rotation = 90,
            direction = "clockwise",
            linecolor = "rgba(0,36,87,0.16)",
            gridcolor = "rgba(0,36,87,0.10)"
          )
        ),
        legend = list(
          orientation = "h",
          y = -0.16,
          x = 0.5,
          xanchor = "center",
          entrywidthmode = if (length(grupos_ord_all) >= 5) "fraction" else NULL,
          entrywidth = if (length(grupos_ord_all) >= 5) 0.18 else NULL,
          title = list(text = "")
        ),
        margin = list(l = 54, r = 54, t = 64, b = 98)
      ) |>
      plotly::config(displayModeBar = FALSE, responsive = TRUE)
  })
}
