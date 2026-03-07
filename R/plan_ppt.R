# =============================================================================
# NUEVO API — PPT "PLAN" (declarativo)
# - presets se definen en un chunk previo con p_presets()
# - p_* crea ELEMENTOS (gráficos / texto / base) con overrides por diapositiva
# - p_slide_* crea SLIDES (layout fijo, sin strings sueltos)
# - reporte_ppt_plan() recolecta diapo_### o recibe plan explícito y exporta
# =============================================================================

#' @title Reporte PowerPoint basado en "plan" (p_* + diapo_###)
#'
#' @description
#' Genera un archivo **.pptx** a partir de un **plan de diapositivas** compuesto por:
#' - **elementos** `p_*()` (gráficos / texto / base),
#' - **slides** `p_slide_*()` (contenedores con layout fijo),
#' - y objetos `diapo_###` (convención para recolección automática).
#'
#' El flujo recomendado es:
#' 1) Definir un objeto de presets con `p_presets()` en un chunk previo.
#' 2) Definir `diapo_001 <- p_slide_*(...)`, `diapo_002 <- ...` (uno o varios chunks).
#' 3) Llamar a `reporte_ppt_plan(presets = presets, ...)` para recolectar y exportar.
#'
#' @param data `data.frame` o `tibble` con las variables (o dummies) a reportar.
#' @param instrumento Objeto de instrumento con al menos `survey` (y opcionalmente `choices`,
#'   `orders_list`). Si es `NULL`, se busca el atributo `instrumento_reporte` en `data`.
#' @param path_ppt Ruta del `.pptx` de salida.
#'
#' @param presets Objeto creado con `p_presets()`. Contiene estilos base por tipo de gráfico
#'   (barras_agrupadas, barras_apiladas, multi_apiladas, pie, donut, numerico, texto/base, etc.)
#'   y configuración de layouts/placeholders.
#'
#' @param plan Lista de slides ya construidos con `p_plan()` o `list(diapo_001=..., ...)`.
#'   Si es `NULL`, se recolectan objetos `diapo_###` desde `env_diapos`.
#' @param env_diapos Entorno donde se buscarán objetos `diapo_###` cuando `plan` sea `NULL`.
#'   Por defecto se usa `parent.frame()`.
#' @param strict_diapos Si `TRUE`, errores en vez de warnings cuando los `diapo_###` no son
#'   consecutivos o cuando hay inconsistencias estructurales (por ejemplo, slot requerido vacío).
#'
#' @param template_pptx Ruta a plantilla `.pptx`. Si es `NULL`, se intenta usar una plantilla
#'   interna del paquete y, si no existe, la plantilla por defecto de PowerPoint.
#' @param master Nombre del master de la plantilla (por defecto `"Office Theme"`).
#'
#' @param mensajes_progreso Si `TRUE`, imprime mensajes de avance durante el proceso.
#' @param solo_lista Si `TRUE`, no se escribe el archivo y solo se retorna el objeto de salida.
#'
#' @return Invisiblemente una lista con:
#' \describe{
#'   \item{doc}{Objeto `officer::pptx` (solo si `solo_lista = TRUE` o para depuración).}
#'   \item{plan}{Plan normalizado de slides (lista).}
#'   \item{log}{Tabla con decisiones/alertas por slide y por elemento.}
#' }
#'
#' @export
reporte_ppt_plan <- function(
    data,
    instrumento       = NULL,
    path_ppt          = "reporte_ppt_plan.pptx",
    presets           = NULL,
    plan              = NULL,
    env_diapos        = parent.frame(),
    strict_diapos     = FALSE,
    template_pptx     = getOption("prosecnur.template_pptx", NA_character_),
    master            = "Office Theme",
    mensajes_progreso = TRUE,
    solo_lista        = FALSE
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y

  # -----------------------
  # 0) Validaciones mínimas
  # -----------------------
  if (!is.data.frame(data)) stop("`data` debe ser un data.frame o tibble.", call. = FALSE)

  if (!requireNamespace("officer", quietly = TRUE) ||
      !requireNamespace("rvg", quietly = TRUE)) {
    stop("Se requieren los paquetes 'officer' y 'rvg'.", call. = FALSE)
  }

  if (is.null(instrumento)) {
    instrumento <- attr(data, "instrumento_reporte", exact = TRUE)
    if (is.null(instrumento)) {
      stop("No se proporcionó `instrumento` y `data` no tiene atributo `instrumento_reporte`.", call. = FALSE)
    }
  }

  survey      <- instrumento$survey %||% NULL
  choices     <- instrumento$choices %||% NULL
  orders_list <- instrumento$orders_list %||% NULL

  if (is.null(survey) || !"name" %in% names(survey)) {
    stop("`instrumento$survey` debe existir y contener al menos la columna `name`.", call. = FALSE)
  }

  # -----------------------
  # 0.1) Presets (tu contrato)
  # -----------------------
  presets <- presets %||% list()
  # defaults mínimos si el usuario no pasó nada
  presets$barras_apiladas <- presets$barras_apiladas %||% list(args = list())
  if (is.null(presets$barras_apiladas$args) || !is.list(presets$barras_apiladas$args)) {
    presets$barras_apiladas$args <- list()
  }
  # defaults de seguridad
  presets$barras_apiladas$args$usar_canvas <- presets$barras_apiladas$args$usar_canvas %||% TRUE
  presets$barras_apiladas$args$exportar    <- presets$barras_apiladas$args$exportar %||% "rplot"
  # defaults para BASE auto (si el usuario no declara base)
  presets$base <- presets$base %||% list()
  presets$base$args <- presets$base$args %||% list()

  presets$base$args$formato   <- presets$base$args$formato %||% "Base: %s"
  presets$base$args$sufijo_auto <- presets$base$args$sufijo_auto %||% NULL

  # defaults para que nunca falle el acceso a $args
  presets$barras_agrupadas <- presets$barras_agrupadas %||% list(args = list())
  presets$barras_agrupadas$args <- presets$barras_agrupadas$args %||% list()

  presets$barras_numericas <- presets$barras_numericas %||% list(args = list())
  presets$barras_numericas$args <- presets$barras_numericas$args %||% list()

  presets$radar_tabla <- presets$radar_tabla %||% list(args = list())
  presets$radar_tabla$args <- presets$radar_tabla$args %||% list()

  # ------------------------------------------------------------
  # HERENCIA: base$args (solo estilo) -> todos los presets$args
  # ------------------------------------------------------------
  base_style <- presets$base$args %||% list()

  # NO pasar estos al graficador (son solo para texto base auto)
  base_style$formato     <- NULL
  base_style$sufijo_auto <- NULL

  # a qué presets se les hereda
  targets <- intersect(
    names(presets),
    c("barras_apiladas", "multi_apiladas", "barras_agrupadas",
      "barras_numericas", "pie", "donut", "radar_tabla")
  )

  for (nm in targets) {
    presets[[nm]]$args <- modifyList(base_style, presets[[nm]]$args %||% list())
  }

  # defaults multi_apiladas
  presets$multi_apiladas <- presets$multi_apiladas %||% list(args = list())
  if (is.null(presets$multi_apiladas$args) || !is.list(presets$multi_apiladas$args)) {
    presets$multi_apiladas$args <- list()
  }

  # heredar defaults de barras_apiladas si quieres (opcional)
  presets$multi_apiladas$args$usar_canvas <- presets$multi_apiladas$args$usar_canvas %||% TRUE
  presets$multi_apiladas$args$exportar    <- presets$multi_apiladas$args$exportar %||% "rplot"

  # defaults pie/donut
  presets$pie   <- presets$pie   %||% list(args = list())
  presets$donut <- presets$donut %||% list(args = list())

  presets$pie$args   <- presets$pie$args   %||% list()
  presets$donut$args <- presets$donut$args %||% list()

  # herencia: donut hereda pie
  presets$donut$args <- .merge_args(presets$pie$args, presets$donut$args)

  # defaults de seguridad
  presets$pie$args$usar_canvas   <- presets$pie$args$usar_canvas   %||% TRUE
  presets$pie$args$exportar      <- presets$pie$args$exportar      %||% "rplot"
  presets$donut$args$usar_canvas <- presets$donut$args$usar_canvas %||% presets$pie$args$usar_canvas
  presets$donut$args$exportar    <- presets$donut$args$exportar    %||% presets$pie$args$exportar

  # ---------------------------------------------------------------------------
  # 1) Helpers — PPT strict con contrato interno (.PPT_CONTRACT)
  # ---------------------------------------------------------------------------
  .layout_exists <- function(layout_name) {
    layout_name %in% layout_info$layout
  }

  .add_slide_strict <- function(doc, layout_name) {
    if (!.layout_exists(layout_name)) {
      stop("La plantilla NO tiene el layout requerido: '", layout_name, "'.", call. = FALSE)
    }
    officer::add_slide(doc, layout = layout_name, master = master)
  }

  .ph_loc <- function(type, type_idx = NULL) {
    if (is.null(type_idx)) return(officer::ph_location_type(type = type))
    tryCatch(
      officer::ph_location_type(type = type, type_idx = type_idx),
      error = function(e) tryCatch(
        officer::ph_location_type(type = type, id = type_idx),
        error = function(e2) officer::ph_location_type(type = type)
      )
    )
  }

  .ph_with_strict <- function(doc, value, spec) {
    if (is.null(spec) || is.null(spec$type)) {
      stop("Placeholder spec inválido (NULL o sin $type).", call. = FALSE)
    }
    loc <- .ph_loc(spec$type, spec$type_idx %||% NULL)
    out <- tryCatch(officer::ph_with(doc, value = value, location = loc), error = identity)
    if (inherits(out, "error")) {
      stop(
        "No se pudo insertar en placeholder type='", spec$type,
        "' type_idx=", spec$type_idx %||% "NULL",
        ". Error: ", conditionMessage(out),
        call. = FALSE
      )
    }
    out
  }

  # ---------------------------------------------------------------------------
  # 2) Helpers — Plan (recolección diapo_###)
  # ---------------------------------------------------------------------------
  .collect_diapos <- function(env, strict = FALSE) {
    nms <- ls(envir = env, all.names = TRUE)
    nms <- nms[grepl("^diapo_\\d{3}$", nms)]
    if (!length(nms)) return(list())

    ord <- order(as.integer(sub("^diapo_(\\d{3})$", "\\1", nms)))
    nms <- nms[ord]
    objs <- mget(nms, envir = env, inherits = FALSE)

    if (isTRUE(strict)) {
      ids <- as.integer(sub("^diapo_(\\d{3})$", "\\1", names(objs)))
      if (length(ids) > 1) {
        dif <- diff(ids)
        if (any(dif != 1L)) stop("strict_diapos=TRUE: los `diapo_###` no son consecutivos.", call. = FALSE)
      }
    }
    objs
  }

  # ---------------------------------------------------------------------------
  # 3) Helpers — Instrumento / tablas / títulos
  # ---------------------------------------------------------------------------
  .pct_enteros_100 <- function(n) {
    n <- as.numeric(n)
    n[is.na(n)] <- 0
    tot <- sum(n)
    if (!is.finite(tot) || tot <= 0) return(rep(0L, length(n)))
    raw <- n / tot * 100
    fl  <- floor(raw)
    resid <- as.integer(round(100 - sum(fl)))
    frac <- raw - fl
    if (resid > 0) {
      idx <- head(order(frac, decreasing = TRUE), resid)
      fl[idx] <- fl[idx] + 1L
    } else if (resid < 0) {
      idx <- head(order(frac, decreasing = FALSE), abs(resid))
      fl[idx] <- pmax(0L, fl[idx] - 1L)
    }
    fl
  }

  .list_name_of_var <- function(var) {
    if ("list_name" %in% names(survey)) {
      x <- survey$list_name[survey$name == var]
      if (length(x)) return(x[1])
    }
    if ("list_norm" %in% names(survey)) {
      x <- survey$list_norm[survey$name == var]
      if (length(x)) return(x[1])
    }
    NA_character_
  }

  .title_of_var <- function(var) {
    if (exists("titulo_var", mode = "function", inherits = TRUE)) {
      return(titulo_var(
        var,
        dic_vars        = NULL,
        labels_override = NULL,
        orders_list     = orders_list,
        df              = data
      ))
    }
    var
  }

  .tab_freq <- function(var) {
    if (!is.character(var) || length(var) != 1L || !nzchar(trimws(var))) {
      stop("`.tab_freq()` requiere `var` como character(1). Recibido length=", length(var), call. = FALSE)
    }
    var <- trimws(var)

    freq_table_spss(
      data,
      var,
      survey        = survey,
      sm_vars_force = NULL,
      orders_list   = orders_list,
      mostrar_todo  = FALSE
    )
  }

  # ---------------------------------------------------------------------------
  # 4) Helpers — paleta_<listname> auto desde env_diapos
  # ---------------------------------------------------------------------------
  .paleta_auto <- function(list_name, env = env_diapos) {
    if (is.na(list_name) || !nzchar(list_name)) return(NULL)
    obj_name <- paste0("paleta_", list_name)
    if (!exists(obj_name, envir = env, inherits = TRUE)) return(NULL)
    pal <- get(obj_name, envir = env, inherits = TRUE)
    if (!is.atomic(pal) || is.null(names(pal))) return(NULL)
    pal
  }

  .base_auto_from_var <- function(var, sufijo_auto = NULL, formato = "Base: %s") {
    if (!is.character(var) || length(var) != 1L || !nzchar(trimws(var))) return(NULL)
    var <- trimws(var)

    tab <- .tab_freq(var)
    if (is.null(tab) || !nrow(tab)) return(NULL)

    N_total <- NA_real_
    if ("Opciones" %in% names(tab) && "n" %in% names(tab)) {
      idx_tot <- which(tab$Opciones == "Total")
      if (length(idx_tot)) N_total <- suppressWarnings(as.numeric(tab$n[idx_tot[1]]))
    }

    tab2 <- tab |>
      dplyr::filter(.data$Opciones != "Total") |>
      dplyr::filter(!is.na(.data$n) & .data$n > 0)

    if (!nrow(tab2)) return(NULL)
    if (!is.finite(N_total)) N_total <- sum(tab2$n, na.rm = TRUE)
    if (!is.finite(N_total)) return(NULL)

    N_pretty <- format(N_total, big.mark = ",", scientific = FALSE)

    # SOLO AUTO: sufijo opcional
    suf <- NULL
    if (!is.null(sufijo_auto) && is.character(sufijo_auto) && length(sufijo_auto) == 1L) {
      sufijo_auto <- trimws(sufijo_auto)
      if (nzchar(sufijo_auto)) suf <- sufijo_auto
    }

    base_core <- if (is.null(suf)) N_pretty else paste(N_pretty, suf)
    sprintf(formato, base_core)
  }

  # ---------------------------------------------------------------------------
  # 5) Renders
  # ---------------------------------------------------------------------------


  # Dispatcher genérico: renderiza cualquier ppt_element
  .render_element <- function(el) {

    if (is.null(el) || !inherits(el, "ppt_element")) {
      stop(".render_element(): `el` debe ser `ppt_element`.", call. = FALSE)
    }

    etype <- el$.element_type %||% NA_character_
    if (is.na(etype) || !nzchar(etype)) {
      stop(".render_element(): elemento sin `.element_type`.", call. = FALSE)
    }

    fn_name <- paste0(".render_", etype)
    if (!exists(fn_name, mode = "function", inherits = TRUE)) {
      stop("No existe renderer para etype='", etype, "' (se esperaba ", fn_name, "()).", call. = FALSE)
    }
    fn <- get(fn_name, mode = "function", inherits = TRUE)

    # presets por tipo (si no existen, lista vacía)
    pa_apiladas <- presets$barras_apiladas$args %||% list()
    pa_multi    <- presets$multi_apiladas$args  %||% list()
    pa_agrup    <- presets$barras_agrupadas$args %||% list()
    pa_num      <- presets$barras_numericas$args %||% list()
    pa_pie      <- presets$pie$args %||% list()
    pa_donut    <- presets$donut$args %||% list()
    pa_radar    <- presets$radar_tabla$args %||% list()

    # helper: llamar pasando SOLO args que la función soporte
    .call_keep_formals <- function(fun, args) {
      fml <- names(formals(fun))
      if ("..." %in% fml) return(do.call(fun, args))
      do.call(fun, args[names(args) %in% fml])
    }

    # Caso especial: multiapiladas (firma distinta)
    if (identical(etype, "barras_multiapiladas")) {
      # firma esperada: (el, preset_args_multi, preset_args_single)
      args <- list(
        el                = el,
        preset_args_multi  = pa_multi,
        preset_args_single = pa_apiladas
      )
      out <- tryCatch(.call_keep_formals(fn, args), error = identity)
      if (inherits(out, "error")) {
        stop(
          "Renderer encontrado (", fn_name, ") pero falló al ejecutarse: ",
          conditionMessage(out),
          call. = FALSE
        )
      }
      return(out)
    }

    # Mapeo estándar: (el, preset_args)
    preset_args <- switch(
      etype,
      barras_apiladas  = pa_apiladas,
      barras_agrupadas = pa_agrup,
      numerico         = pa_num,
      pie              = pa_pie,
      donut            = pa_donut,
      radar_tabla      = pa_radar,
      # default: si hay nuevos etypes, se intenta pasar lista vacía
      list()
    )

    args <- list(el = el, preset_args = preset_args)
    out <- tryCatch(.call_keep_formals(fn, args), error = identity)

    if (inherits(out, "error")) {
      # fallback final: intentar SOLO con `el` (por si un renderer nuevo no usa presets)
      out2 <- tryCatch(do.call(fn, list(el = el)), error = identity)
      if (!inherits(out2, "error")) return(out2)

      stop(
        "Renderer encontrado (", fn_name, ") pero falló al ejecutarse: ",
        conditionMessage(out),
        call. = FALSE
      )
    }

    out
  }

  .render_barras_apiladas <- function(el, preset_args) {
    var <- el$var
    tab <- .tab_freq(var)
    if (is.null(tab) || !nrow(tab)) return(NULL)

    # N desde Total si existe
    N_total <- NA_real_
    if ("Opciones" %in% names(tab) && "n" %in% names(tab)) {
      idx_tot <- which(tab$Opciones == "Total")
      if (length(idx_tot)) N_total <- suppressWarnings(as.numeric(tab$n[idx_tot[1]]))
    }

    tab <- tab |>
      dplyr::filter(.data$Opciones != "Total") |>
      dplyr::filter(!is.na(.data$n) & .data$n > 0)

    if (!nrow(tab)) return(NULL)
    if (!is.finite(N_total)) N_total <- sum(tab$n, na.rm = TRUE)

    pct_int  <- .pct_enteros_100(tab$n)
    cols_pct <- paste0("pct_", seq_len(nrow(tab)))

    df_wide <- tibble::tibble(
      categoria = .title_of_var(var),
      N         = N_total
    )
    for (i in seq_along(cols_pct)) df_wide[[cols_pct[i]]] <- pct_int[i] / 100

    etiquetas_grupos <- stats::setNames(as.character(tab$Opciones), cols_pct)

    # paleta auto (paleta_<listname>)
    ln <- .list_name_of_var(var)
    colores_grupos <- .paleta_auto(ln, env_diapos)

    if (!exists("graficar_barras_apiladas", mode = "function", inherits = TRUE)) {
      stop("No existe `graficar_barras_apiladas()` en el entorno/paquete.", call. = FALSE)
    }

    # base args mínimos + preset_args + overrides
    base_args <- list(
      data             = df_wide,
      var_categoria    = "categoria",
      var_n            = "N",
      cols_porcentaje  = cols_pct,
      etiquetas_grupos = etiquetas_grupos,
      escala_valor     = "proporcion_1",
      colores_grupos   = colores_grupos,
      titulo           = NULL,
      subtitulo        = NULL,
      nota_pie         = NULL
    )

    # merge: base_args <- preset_args <- overrides (overrides manda)
    preset_args <- preset_args %||% list()
    overrides   <- el$overrides %||% list()

    args <- .merge_args(base_args, preset_args, overrides)
    fun  <- graficar_barras_apiladas
    args <- .keep_formals(fun, args)
    suppressWarnings(do.call(fun, args))
  }


  .render_barras_multiapiladas <- function(el, preset_args_multi, preset_args_single) {

    `%||%` <- function(x, y) if (!is.null(x)) x else y

    modo <- el$modo %||% "var"

    # ============================================================
    # helpers locales
    # ============================================================
    .clean_chr <- function(x) {
      x <- as.character(x)
      x[is.na(x)] <- ""
      trimws(x)
    }

    # ============================================================
    # MODO "var"
    # ============================================================
    if (identical(modo, "var")) {

      vars <- el$vars
      if (!is.character(vars) || length(vars) < 1L) return(NULL)
      vars <- trimws(vars); vars <- vars[nzchar(vars)]
      if (!length(vars)) return(NULL)

      # regla fuerte: 1 list_name para todo el bloque
      lns <- vapply(vars, .list_name_of_var, character(1))
      lns <- unique(lns[!is.na(lns) & nzchar(lns)])
      if (length(lns) != 1L) {
        stop("multiapiladas (modo='var'): las vars no comparten un único list_name. Encontrados: ",
             paste(lns, collapse = " | "), call. = FALSE)
      }
      ln <- lns[1]

      colores_grupos <- .paleta_auto(ln, env_diapos)

      rows <- list()
      all_opts <- character(0)
      tabs_by_v <- list()
      N_by_v <- numeric(0)

      for (v in vars) {
        tab <- .tab_freq(v)
        if (is.null(tab) || !nrow(tab)) next

        N_total <- NA_real_
        if ("Opciones" %in% names(tab) && "n" %in% names(tab)) {
          idx_tot <- which(tab$Opciones == "Total")
          if (length(idx_tot)) N_total <- suppressWarnings(as.numeric(tab$n[idx_tot[1]]))
        }

        tab <- tab |>
          dplyr::filter(.data$Opciones != "Total") |>
          dplyr::filter(!is.na(.data$n) & .data$n > 0)

        if (!nrow(tab)) next
        if (!is.finite(N_total)) N_total <- sum(tab$n, na.rm = TRUE)

        tabs_by_v[[v]] <- tab
        N_by_v[v] <- N_total
        all_opts <- union(all_opts, as.character(tab$Opciones))
      }

      if (!length(tabs_by_v)) return(NULL)

      # ordenar niveles: paleta -> choices -> lo observado
      niveles_formales <- character(0)
      if (!is.null(colores_grupos) && is.atomic(colores_grupos) && !is.null(names(colores_grupos))) {
        niveles_formales <- names(colores_grupos)
      } else if (!is.null(choices) && "list_name" %in% names(choices) && "label" %in% names(choices)) {
        niveles_formales <- as.character(choices$label[choices$list_name == ln])
      }
      niveles_formales <- niveles_formales[!is.na(niveles_formales) & nzchar(niveles_formales)]
      if (length(niveles_formales)) all_opts <- intersect(niveles_formales, all_opts)

      cols_pct <- paste0("pct_", seq_along(all_opts))
      etiquetas_grupos <- stats::setNames(all_opts, cols_pct)

      for (v in vars) {
        tab <- tabs_by_v[[v]]
        if (is.null(tab)) next

        label_v <- .title_of_var(v)
        if (requireNamespace("stringr", quietly = TRUE)) {
          label_v <- stringr::str_wrap(label_v, width = el$wrap_y %||% 50)
        }

        pct_int <- .pct_enteros_100(tab$n)
        names(pct_int) <- as.character(tab$Opciones)

        row <- tibble::tibble(
          categoria = label_v,
          N         = unname(N_by_v[v])
        )
        for (i in seq_along(all_opts)) {
          opt <- all_opts[i]
          row[[cols_pct[i]]] <- (pct_int[opt] %||% 0) / 100
        }
        rows[[length(rows) + 1]] <- row
      }

      if (!length(rows)) return(NULL)
      df_block <- dplyr::bind_rows(rows)

      base_args <- list(
        data             = df_block,
        var_categoria    = "categoria",
        var_n            = "N",
        cols_porcentaje  = cols_pct,
        etiquetas_grupos = etiquetas_grupos,
        escala_valor     = "proporcion_1",
        colores_grupos   = colores_grupos,
        titulo           = NULL,
        subtitulo        = NULL,
        nota_pie         = NULL
      )

      # ============================================================
      # NUEVO: TOP TWO BOX (alias del wrapper -> args nativos)
      #   - NO depende del "preset"
      #   - fuerza barra_extra_preset="top2box"
      # ============================================================
      if (isTRUE(el$top2box)) {
        base_args$mostrar_barra_extra <- TRUE
        base_args$barra_extra_preset  <- "top2box"

        # Si el usuario no pasa labels, el graficador usa defaults (tail cols)
        if (!is.null(el$top2box_labels) && length(el$top2box_labels)) {
          base_args$top2box_labels <- el$top2box_labels
        }
        if (is.null(base_args$titulo_barra_extra) || !nzchar(base_args$titulo_barra_extra)) {
          base_args$titulo_barra_extra <- "TOP TWO BOX"
        }
      }

      preset_args_multi  <- preset_args_multi  %||% list()
      preset_args_single <- preset_args_single %||% list()
      overrides          <- el$overrides %||% list()

      args <- .merge_args(base_args, preset_args_single, preset_args_multi, overrides)
      fun  <- graficar_barras_apiladas
      args <- .keep_formals(fun, args)
      return(suppressWarnings(do.call(fun, args)))
    }

    # ============================================================
    # MODO "cruce" (NUEVO)
    #   - 1 fila por nivel del cruce
    #   - segmentos = opciones de `var`
    # ============================================================
    if (identical(modo, "cruce")) {

      var   <- el$var %||% NULL
      cruce <- el$cruce %||% NULL

      if (!is.character(var) || length(var) != 1L || !nzchar(trimws(var))) return(NULL)
      if (!is.character(cruce) || length(cruce) != 1L || !nzchar(trimws(cruce))) {
        stop("multiapiladas (modo='cruce'): falta `cruce` (character(1)).", call. = FALSE)
      }
      var   <- trimws(var)
      cruce <- trimws(cruce)

      # --- segmentos: opciones de var (y paleta de var)
      ln_var <- .list_name_of_var(var)
      if (is.na(ln_var) || !nzchar(ln_var)) {
        stop("multiapiladas (modo='cruce'): no se encontró list_name para `var`=", var, call. = FALSE)
      }
      colores_grupos <- .paleta_auto(ln_var, env_diapos)

      # --- niveles del cruce (keys para filtrar + labels para mostrar) usando instrumento
      cm <- .radar_cruce_map(
        data        = data,
        cruce       = cruce,
        survey      = survey,
        orders_list = orders_list,
        env_paletas = env_diapos
      )
      lvls_keys   <- cm$keys
      lvls_labels <- cm$labels

      lvls_keys   <- .clean_chr(lvls_keys);   lvls_keys   <- lvls_keys[nzchar(lvls_keys)]
      lvls_labels <- .clean_chr(lvls_labels); lvls_labels <- lvls_labels[nzchar(lvls_labels)]

      # fallback si algo raro
      if (!length(lvls_keys) || !length(lvls_labels)) {
        x <- .clean_chr(data[[cruce]])
        lvls_keys <- sort(unique(x[nzchar(x)]))
        lvls_labels <- lvls_keys
      }

      # --- primero, descubrir el set de opciones (segmentos) de var (sobre total)
      tab_total <- .tab_freq(var)
      if (is.null(tab_total) || !nrow(tab_total)) return(NULL)

      tab_total <- tab_total |>
        dplyr::filter(.data$Opciones != "Total") |>
        dplyr::filter(!is.na(.data$n) & .data$n > 0)

      if (!nrow(tab_total)) return(NULL)

      all_opts <- as.character(tab_total$Opciones)

      # ordenar opciones: paleta -> choices -> observado
      if (!is.null(colores_grupos) && is.atomic(colores_grupos) && !is.null(names(colores_grupos))) {
        pref <- names(colores_grupos)
        pref <- pref[!is.na(pref) & nzchar(pref)]
        if (length(pref)) all_opts <- intersect(pref, all_opts)
      } else if (!is.null(choices) && "list_name" %in% names(choices) && "label" %in% names(choices)) {
        pref <- as.character(choices$label[choices$list_name == ln_var])
        pref <- pref[!is.na(pref) & nzchar(pref)]
        if (length(pref)) all_opts <- intersect(pref, all_opts)
      }

      cols_pct <- paste0("pct_", seq_along(all_opts))
      etiquetas_grupos <- stats::setNames(all_opts, cols_pct)

      # --- construir 1 fila por nivel del cruce
      rows <- list()

      x_cruce <- .clean_chr(data[[cruce]])

      for (j in seq_along(lvls_keys)) {

        key_j <- lvls_keys[j]
        lab_j <- lvls_labels[j]

        mask <- nzchar(x_cruce) & (x_cruce == .clean_chr(key_j))

        dsub <- data[mask, , drop = FALSE]
        if (!nrow(dsub)) next

        tab <- freq_table_spss(
          dsub,
          var,
          survey        = survey,
          sm_vars_force = NULL,
          orders_list   = orders_list,
          mostrar_todo  = FALSE
        )

        if (is.null(tab) || !nrow(tab)) next

        # N desde Total si existe
        N_total <- NA_real_
        if ("Opciones" %in% names(tab) && "n" %in% names(tab)) {
          idx_tot <- which(tab$Opciones == "Total")
          if (length(idx_tot)) N_total <- suppressWarnings(as.numeric(tab$n[idx_tot[1]]))
        }

        tab <- tab |>
          dplyr::filter(.data$Opciones != "Total") |>
          dplyr::filter(!is.na(.data$n) & .data$n > 0)

        if (!nrow(tab)) next
        if (!is.finite(N_total)) N_total <- sum(tab$n, na.rm = TRUE)

        # pct enteros a 100 dentro del grupo
        pct_int <- .pct_enteros_100(tab$n)
        names(pct_int) <- as.character(tab$Opciones)

        cat_j <- as.character(lab_j)
        if (requireNamespace("stringr", quietly = TRUE)) {
          cat_j <- stringr::str_wrap(cat_j, width = el$wrap_y %||% 50)
        }

        row <- tibble::tibble(
          categoria = cat_j,
          N         = N_total
        )
        for (i in seq_along(all_opts)) {
          opt <- all_opts[i]
          row[[cols_pct[i]]] <- (pct_int[opt] %||% 0) / 100
        }

        rows[[length(rows) + 1]] <- row
      }

      if (!length(rows)) return(NULL)
      df_block <- dplyr::bind_rows(rows)

      base_args <- list(
        data             = df_block,
        var_categoria    = "categoria",
        var_n            = "N",
        cols_porcentaje  = cols_pct,
        etiquetas_grupos = etiquetas_grupos,
        escala_valor     = "proporcion_1",
        colores_grupos   = colores_grupos,
        titulo           = NULL,
        subtitulo        = NULL,
        nota_pie         = NULL
      )

      # ============================================================
      # NUEVO: TOP TWO BOX (alias del wrapper -> args nativos)
      # ============================================================
      if (isTRUE(el$top2box)) {
        base_args$mostrar_barra_extra <- TRUE
        base_args$barra_extra_preset  <- "top2box"
        if (!is.null(el$top2box_labels) && length(el$top2box_labels)) {
          base_args$top2box_labels <- el$top2box_labels
        }
        if (is.null(base_args$titulo_barra_extra) || !nzchar(base_args$titulo_barra_extra)) {
          base_args$titulo_barra_extra <- "TOP TWO BOX"
        }
      }

      preset_args_multi  <- preset_args_multi  %||% list()
      preset_args_single <- preset_args_single %||% list()
      overrides          <- el$overrides %||% list()

      args <- .merge_args(base_args, preset_args_single, preset_args_multi, overrides)
      fun  <- graficar_barras_apiladas
      args <- .keep_formals(fun, args)
      return(suppressWarnings(do.call(fun, args)))
    }

    stop("multiapiladas: modo no soportado: ", modo, call. = FALSE)
  }


  .render_barras_agrupadas <- function(el, preset_args) {

    var <- el$var
    tab <- .tab_freq(var)
    if (is.null(tab) || !nrow(tab)) return(NULL)

    # N desde Total si existe
    N_total <- NA_real_
    if ("Opciones" %in% names(tab) && "n" %in% names(tab)) {
      idx_tot <- which(tab$Opciones == "Total")
      if (length(idx_tot)) N_total <- suppressWarnings(as.numeric(tab$n[idx_tot[1]]))
    }

    tab <- tab |>
      dplyr::filter(.data$Opciones != "Total") |>
      dplyr::filter(!is.na(.data$n) & .data$n > 0)

    if (!nrow(tab)) return(NULL)

    if (!is.finite(N_total)) N_total <- sum(tab$n, na.rm = TRUE)
    if (!is.finite(N_total) || N_total <= 0) return(NULL)

    # ----------------------------
    # LONG: 1 fila por opción
    # (esto evita: eje Y con "título" y colores distintos por opción)
    # ----------------------------
    df_long <- tibble::tibble(
      categoria = as.character(tab$Opciones),
      N         = N_total,
      pct       = as.numeric(tab$n) / N_total
    )

    etiquetas_series <- c(pct = "Porcentaje")

    if (!exists("graficar_barras_agrupadas", mode = "function", inherits = TRUE)) {
      stop("No existe `graficar_barras_agrupadas()` en el entorno/paquete.", call. = FALSE)
    }

    base_args <- list(
      data             = df_long,
      var_categoria    = "categoria",
      var_n            = "N",
      cols_porcentaje  = "pct",
      etiquetas_series = etiquetas_series,
      titulo           = NULL,
      subtitulo        = NULL,
      nota_pie         = NULL
    )

    preset_args <- preset_args %||% list()
    overrides   <- el$overrides %||% list()

    # limpiar cosas que NO aplican a agrupadas (por si vienen de presets genéricos)
    preset_args$var_grupo      <- NULL
    preset_args$colores_grupos <- NULL
    overrides$var_grupo        <- NULL
    overrides$colores_grupos   <- NULL

    args <- .merge_args(base_args, preset_args, overrides)
    fun  <- graficar_barras_agrupadas
    args <- .keep_formals(fun, args)

    suppressWarnings(do.call(fun, args))
  }

  .render_pie <- function(el, preset_args, tipo_pie = c("pie", "donut")) {
    tipo_pie <- match.arg(tipo_pie)

    var <- el$var
    tab <- .tab_freq(var)
    if (is.null(tab) || !nrow(tab)) return(NULL)

    tab <- tab |>
      dplyr::filter(.data$Opciones != "Total") |>
      dplyr::filter(!is.na(.data$n) & .data$n > 0)

    if (!nrow(tab)) return(NULL)

    df_long <- tab |>
      dplyr::transmute(
        opcion = as.character(.data$Opciones),
        n      = as.numeric(.data$n)
      ) |>
      dplyr::mutate(
        pct = .data$n / sum(.data$n, na.rm = TRUE)  # proporción 0-1
      )

    ln <- .list_name_of_var(var)
    colores_grupos <- .paleta_auto(ln, env_diapos)

    if (!exists("graficar_pie", mode = "function", inherits = TRUE)) {
      stop("No existe `graficar_pie()` en el entorno/paquete.", call. = FALSE)
    }

    base_args <- list(
      data           = df_long,
      var_categoria  = "opcion",
      var_pct        = "pct",
      tipo_pie       = tipo_pie,
      colores_categorias = colores_grupos,
      titulo         = NULL,
      subtitulo      = NULL,
      nota_pie       = NULL
    )

    preset_args <- preset_args %||% list()
    overrides   <- el$overrides %||% list()

    args <- .merge_args(base_args, preset_args, overrides)

    fun  <- graficar_pie
    args <- .keep_formals(fun, args)

    suppressWarnings(do.call(fun, args))
  }

  .render_donut <- function(el, preset_args) {
    .render_pie(el, preset_args = preset_args, tipo_pie = "donut")
  }

  .render_numerico <- function(el, preset_args) {

    `%||%` <- function(x, y) if (!is.null(x)) x else y

    var <- el$var
    if (is.null(var) || !nzchar(var)) return(NULL)

    preset_args <- preset_args %||% list()
    overrides   <- el$overrides %||% list()

    cruce <- overrides$cruce %||% el$cruce %||% preset_args$cruce %||% NULL
    preset_args$cruce <- NULL
    overrides$cruce   <- NULL

    df <- NULL
    if (!is.null(el$data) && is.data.frame(el$data)) df <- el$data
    if (is.null(df) && !is.null(el$df) && is.data.frame(el$df)) df <- el$df

    if (is.null(df) && exists(".df", inherits = TRUE)) {
      tmp <- get(".df", inherits = TRUE)
      if (is.data.frame(tmp)) df <- tmp
    }
    if (is.null(df) && exists("data", inherits = TRUE)) {
      tmp <- get("data", inherits = TRUE)
      if (is.data.frame(tmp)) df <- tmp
    }

    if (is.null(df) || !is.data.frame(df)) {
      stop("`.render_numerico`: no se encontró un data.frame válido en `el$data/el$df` ni en el entorno.", call. = FALSE)
    }
    if (!var %in% names(df)) return(NULL)

    if (!is.null(cruce)) {
      if (!is.character(cruce) || length(cruce) != 1L || !nzchar(cruce)) cruce <- NULL
      if (!is.null(cruce) && !cruce %in% names(df)) {
        stop("`.render_numerico`: el cruce '", cruce, "' no existe en `df`.", call. = FALSE)
      }
    }

    .get_inst <- function() {
      cand <- list(el$instrumento, el$inst, el$rp_inst)
      cand <- cand[!vapply(cand, is.null, logical(1))]
      for (obj in cand) if (is.list(obj) && !is.null(obj$survey)) return(obj)

      for (nm in c(".inst", "inst", "instrumento", "rp_inst")) {
        if (exists(nm, inherits = TRUE)) {
          obj <- get(nm, inherits = TRUE)
          if (is.list(obj) && !is.null(obj$survey)) return(obj)
        }
      }
      NULL
    }

    .labels_from_inst <- function(inst, varname) {
      if (is.null(inst) || is.null(inst$survey)) return(NULL)
      surv <- inst$survey
      if (!("name" %in% names(surv))) return(NULL)

      ln <- NA_character_
      if ("list_name" %in% names(surv)) {
        tmp <- surv$list_name[surv$name == varname]
        if (length(tmp)) ln <- tmp[1]
      } else if ("list_norm" %in% names(surv)) {
        tmp <- surv$list_norm[surv$name == varname]
        if (length(tmp)) ln <- tmp[1]
      }
      if (is.na(ln) || !nzchar(ln)) return(NULL)

      ch <- inst$choices_raw %||% inst$choices %||% NULL
      if (is.null(ch) || !("list_name" %in% names(ch)) || !("name" %in% names(ch))) return(NULL)

      lab_col <- NULL
      if ("label::Spanish (ES)" %in% names(ch)) lab_col <- "label::Spanish (ES)"
      if (is.null(lab_col) && "label" %in% names(ch)) lab_col <- "label"
      if (is.null(lab_col)) return(NULL)

      sub <- ch[ch$list_name == ln, , drop = FALSE]
      if (!nrow(sub)) return(NULL)

      codes  <- as.character(sub$name)
      labels <- as.character(sub[[lab_col]])
      out <- stats::setNames(labels, codes)
      attr(out, "levels_labels") <- labels
      out
    }

    .apply_cruce_labels <- function(x_cruce, inst, cruce_name) {

      if (requireNamespace("haven", quietly = TRUE) &&
          inherits(x_cruce, "haven_labelled")) {
        x_chr <- as.character(haven::as_factor(x_cruce, levels = "labels"))
        lvls  <- unique(x_chr)
        return(list(x = x_chr, lvls = lvls))
      }

      if (is.factor(x_cruce)) {
        x_chr <- as.character(x_cruce)
        return(list(x = x_chr, lvls = levels(x_cruce)))
      }

      map <- .labels_from_inst(inst, cruce_name)
      if (!is.null(map)) {
        x_chr <- as.character(x_cruce)
        x_lab <- ifelse(x_chr %in% names(map), unname(map[x_chr]), x_chr)

        lvls <- attr(map, "levels_labels")
        lvls <- lvls[!is.na(lvls) & nzchar(lvls)]
        extras <- setdiff(unique(x_lab), lvls)
        lvls2  <- c(lvls, extras)

        return(list(x = x_lab, lvls = lvls2))
      }

      x_chr <- as.character(x_cruce)
      return(list(x = x_chr, lvls = unique(x_chr)))
    }

    x_raw <- df[[var]]
    if (is.factor(x_raw)) x_raw <- as.character(x_raw)
    x <- suppressWarnings(as.numeric(x_raw))

    nombre_serie   <- preset_args$nombre_serie   %||% overrides$nombre_serie   %||% "v1"
    etiqueta_serie <- preset_args$etiqueta_serie %||% overrides$etiqueta_serie %||% "Media"

    preset_args$nombre_serie   <- NULL
    preset_args$etiqueta_serie <- NULL
    overrides$nombre_serie     <- NULL
    overrides$etiqueta_serie   <- NULL

    if (is.null(cruce)) {

      x2 <- x[is.finite(x)]
      if (!length(x2)) return(NULL)

      N <- length(x2)
      m <- mean(x2, na.rm = TRUE)
      if (!is.finite(m)) return(NULL)

      cat_label <- tryCatch(.title_of_var(var), error = function(e) var)
      if (is.null(cat_label) || !nzchar(cat_label)) cat_label <- var

      df_wide <- tibble::tibble(
        categoria = cat_label,
        N         = N
      )
      df_wide[[nombre_serie]] <- m

    } else {

      inst <- .get_inst()
      cr <- .apply_cruce_labels(df[[cruce]], inst, cruce)

      d2 <- tibble::tibble(
        .cruce = cr$x,
        .x     = x
      )

      d2 <- d2[is.finite(d2$.x), , drop = FALSE]
      d2 <- d2[!is.na(d2$.cruce) & nzchar(trimws(as.character(d2$.cruce))), , drop = FALSE]
      if (!nrow(d2)) return(NULL)

      df_wide <- d2 |>
        dplyr::group_by(.data$.cruce) |>
        dplyr::summarise(
          N  = dplyr::n(),
          .m = mean(.data$.x, na.rm = TRUE),
          .groups = "drop"
        ) |>
        dplyr::rename(categoria = .data$.cruce)

      df_wide[[nombre_serie]] <- df_wide$.m
      df_wide$.m <- NULL

      lvls <- cr$lvls
      if (!is.null(lvls) && length(lvls)) {
        df_wide$categoria <- factor(df_wide$categoria, levels = lvls)
      }
    }

    if (!nrow(df_wide) || all(!is.finite(df_wide[[nombre_serie]]))) return(NULL)

    base_args <- list(
      data             = df_wide,
      var_categoria    = "categoria",
      var_n            = "N",
      vars_valor       = nombre_serie,
      etiquetas_series = stats::setNames(etiqueta_serie, nombre_serie),

      titulo           = NULL,
      subtitulo        = NULL,
      nota_pie         = NULL,

      usar_canvas      = TRUE,
      exportar         = "rplot"
    )

    for (k in c("titulo","subtitulo","nota_pie","title","subtitle","caption","main","sub")) {
      if (!is.null(preset_args[[k]])) preset_args[[k]] <- NULL
      if (!is.null(overrides[[k]]))   overrides[[k]]   <- NULL
    }

    if (!exists("graficar_barras_numericas", mode = "function", inherits = TRUE)) {
      stop("No existe `graficar_barras_numericas()` en el entorno/paquete.", call. = FALSE)
    }

    fun  <- graficar_barras_numericas
    args <- .merge_args(base_args, preset_args, overrides)
    args <- .keep_formals(fun, args)

    tryCatch(
      suppressWarnings(do.call(fun, args)),
      error = function(e) {
        message("⚠️ .render_numerico(): ", conditionMessage(e))
        NULL
      }
    )
  }

  .render_radar_tabla <- function(el, preset_args) {

    if (!exists("graficar_radar", mode = "function", inherits = TRUE)) {
      stop("No existe `graficar_radar()` en el entorno/paquete.", call. = FALSE)
    }

    modo  <- el$modo %||% "sm"
    cruce <- el$cruce %||% NULL
    titulo_tabla <- el$titulo_tabla %||% if (modo == "sm") "Opciones" else "Top 2 Box"

    if (identical(modo, "sm")) {

      omit_codes  <- el$sm_omit_codes  %||% preset_args$sm_omit_codes  %||% NULL
      omit_labels <- el$sm_omit_labels %||% preset_args$sm_omit_labels %||% NULL
      omit_na     <- el$sm_omit_na     %||% preset_args$sm_omit_na     %||% TRUE

      d_radar <- .radar_build_sm(
        var         = el$var,
        cruce       = cruce,
        top_n       = el$top_n %||% NULL,

        sm_omit_codes  = omit_codes,
        sm_omit_labels = omit_labels,
        sm_omit_na     = omit_na,

        data        = data,
        survey      = survey,
        orders_list = orders_list,
        env_paletas = env_diapos
      )
    } else if (identical(modo, "box")) {
      d_radar <- .radar_build_box(
        vars        = el$vars,
        cruce       = cruce,
        box_labels  = el$box_labels,
        titulo_tabla = titulo_tabla,
        data        = data,
        survey      = survey,
        orders_list = orders_list,
        env_paletas = env_diapos
      )
    } else {
      stop("radar_tabla: modo no soportado: ", modo, call. = FALSE)
    }

    if (is.null(d_radar) || !nrow(d_radar)) return(NULL)

    base_args <- list(
      data         = d_radar,
      var_eje      = "eje",
      var_grupo    = "grupo",
      var_valor    = "valor",
      titulo_tabla = titulo_tabla
    )

    # -----------------------------
    # FIX: pasar paleta del CRUCE
    # -----------------------------
    pal_series <- attr(d_radar, "palette", exact = TRUE)

    if (!is.null(pal_series) && is.atomic(pal_series) && length(pal_series) && !is.null(names(pal_series))) {

      # asegurar que los nombres calcen con los niveles reales de `grupo`
      grupos_lvl <- NULL
      if ("grupo" %in% names(d_radar)) {
        if (is.factor(d_radar$grupo)) grupos_lvl <- levels(d_radar$grupo)
        else grupos_lvl <- sort(unique(as.character(d_radar$grupo)))
      }
      if (length(grupos_lvl)) {
        pal_series <- pal_series[names(pal_series) %in% grupos_lvl]
      }

      # inyectar en el argumento correcto según cómo se llame en graficar_radar()
      fml <- names(formals(graficar_radar))

      if ("colores_series" %in% fml) {
        base_args$colores_series <- pal_series
      } else if ("colores_grupos" %in% fml) {
        base_args$colores_grupos <- pal_series
      } else if ("colores_lineas" %in% fml) {
        base_args$colores_lineas <- pal_series
      } else if ("palette" %in% fml) {
        base_args$palette <- pal_series
      } else if ("paleta" %in% fml) {
        base_args$paleta <- pal_series
      } else {
        # último recurso: meterlo en overrides por si tu graficar_radar lo recoge allí
        overrides$colores_series <- overrides$colores_series %||% pal_series
        overrides$colores_grupos <- overrides$colores_grupos %||% pal_series
        overrides$colores_lineas <- overrides$colores_lineas %||% pal_series
      }
    }

    preset_args <- preset_args %||% list()
    overrides   <- el$overrides %||% list()

    args <- .merge_args(base_args, preset_args, overrides)
    fun  <- graficar_radar
    args <- .keep_formals(fun, args)

    suppressWarnings(do.call(fun, args))
  }

  # ---------------------------------------------------------------------------
  # 6) Normalizar plan
  # ---------------------------------------------------------------------------
  if (is.null(plan)) {
    plan_accum <- NULL
    if (exists(.ppt_plan_name, envir = env_diapos, inherits = TRUE)) {
      cand <- get(.ppt_plan_name, envir = env_diapos, inherits = TRUE)
      if (is.list(cand) && length(cand)) {
        plan_accum <- cand
        class(plan_accum) <- unique(c("ppt_plan","list", class(plan_accum)))
      }
    }

    if (!is.null(plan_accum) && length(plan_accum)) {
      plan <- plan_accum
      .validate_plan(plan, strict = strict_diapos)

    } else {
      diapos <- .collect_diapos(env = env_diapos, strict = strict_diapos)
      if (!length(diapos)) {
        plan <- structure(list(), class = c("ppt_plan", "list"))
      } else {
        plan <- unname(diapos)
        class(plan) <- c("ppt_plan", "list")
        attr(plan, "diapo_names") <- names(diapos)
      }
      .validate_plan(plan, strict = strict_diapos)
    }

  } else {
    if (!is.list(plan)) stop("`plan` debe ser una lista de slides.", call. = FALSE)
    .validate_plan(plan, strict = strict_diapos)
  }

  if (!length(plan)) stop("No hay diapositivas...", call. = FALSE)

  # ---------------------------------------------------------------------------
  # 7) Abrir plantilla / doc (solo si exporta)
  # ---------------------------------------------------------------------------
  if (isTRUE(solo_lista)) {
    doc <- NULL
  } else {

    # Si el usuario no pasó template_pptx (NULL/NA/"") -> intentar interna
    if (is.null(template_pptx) || is.na(template_pptx) || !nzchar(template_pptx)) {

      template_interno <- system.file("plantillas/plantilla_16_9.pptx", package = "prosecnur")

      if (nzchar(template_interno) && file.exists(template_interno)) {
        if (isTRUE(mensajes_progreso)) message("Usando plantilla interna: ", template_interno)
        doc <- officer::read_pptx(path = template_interno)
      } else {
        if (isTRUE(mensajes_progreso)) message("No se encontró plantilla interna. Usando PPT default.")
        doc <- officer::read_pptx()
      }

    } else {
      # Plantilla externa explícita
      if (!file.exists(template_pptx)) stop("No existe `template_pptx`: ", template_pptx, call. = FALSE)
      if (isTRUE(mensajes_progreso)) message("Usando plantilla externa: ", template_pptx)
      doc <- officer::read_pptx(path = template_pptx)
    }
  }

  layout_info <- tryCatch(officer::layout_summary(doc), error = function(e) NULL)
  if (is.null(layout_info) || !nrow(layout_info)) {
    stop("No se pudo leer `layout_summary()` del PPT.", call. = FALSE)
  }

  .pick_layout <- function(candidates) {
    hit <- candidates[candidates %in% layout_info$layout][1]
    if (length(hit) == 0 || is.na(hit)) return(NA_character_)
    hit
  }

  # Preferencias
  layout_graficos <- .pick_layout(c("Graficos2", "Graficos"))
  layout_doble    <- .pick_layout(c("Graficos_2columnas"))
  layout_title      <- .pick_layout(c("Title Slide"))
  layout_poblacion4 <- .pick_layout(c("poblacion_4"))
  layout_text_right <- .pick_layout(c("right_grafico_texto"))
  layout_text_left  <- .pick_layout(c("left_grafico_texto"))
  layout_text_right2 <- .pick_layout(c("right_2graficos_texto"))
  layout_text_left2  <- .pick_layout(c("left_2graficos_texto"))
  layout_poblacion_2         <- .pick_layout(c("poblacion_2"))
  layout_poblacion_5         <- .pick_layout(c("poblacion_5"))
  layout_poblacion_6         <- .pick_layout(c("poblacion_6"))

  if (is.na(layout_graficos)) {
    stop("La plantilla NO tiene layout requerido: 'Graficos' o 'Graficos2'.", call. = FALSE)
  }
  if (is.na(layout_doble)) {
    stop("La plantilla NO tiene layout requerido: 'Graficos_2columnas'.", call. = FALSE)
  }
  if (is.na(layout_title)) {
    stop("La plantilla NO tiene layout requerido: 'Title Slide'.", call. = FALSE)
  }
  if (is.na(layout_poblacion4)) {
    stop("La plantilla NO tiene layout requerido: 'poblacion_4'.", call. = FALSE)
  }
  if (is.na(layout_text_right)) {
    stop("La plantilla NO tiene layout requerido: 'right_grafico_texto'.", call. = FALSE)
  }
  if (is.na(layout_text_left)) {
    stop("La plantilla NO tiene layout requerido: 'left_grafico_texto'.", call. = FALSE)
  }
  if (is.na(layout_text_right2)) {
    stop("La plantilla NO tiene layout requerido: 'right_2graficos_texto'.", call. = FALSE)
  }
  if (is.na(layout_text_left2)) {
    stop("La plantilla NO tiene layout requerido: 'left_2graficos_texto'.", call. = FALSE)
  }
  if (is.na(layout_poblacion_2)) stop("La plantilla NO tiene layout requerido: 'poblacion_2'.", call. = FALSE)
  if (is.na(layout_poblacion_5)) stop("La plantilla NO tiene layout requerido: 'poblacion_5'.", call. = FALSE)
  if (is.na(layout_poblacion_6)) stop("La plantilla NO tiene layout requerido: 'poblacion_6'.", call. = FALSE)



  PPT_CONTRACT <- .PPT_CONTRACT
  PPT_CONTRACT$slide_1$layout  <- layout_graficos
  PPT_CONTRACT$slide_2$layout  <- layout_doble
  PPT_CONTRACT$title_slide$layout  <- layout_title
  PPT_CONTRACT$poblacion_4$layout  <- layout_poblacion4
  PPT_CONTRACT$text_r$layout <- layout_text_right
  PPT_CONTRACT$text_l$layout <- layout_text_left
  PPT_CONTRACT$text_r2$layout <- layout_text_right2
  PPT_CONTRACT$text_l2$layout <- layout_text_left2
  PPT_CONTRACT$poblacion_2$layout <- layout_poblacion_2
  PPT_CONTRACT$poblacion_5$layout <- layout_poblacion_5
  PPT_CONTRACT$poblacion_6$layout <- layout_poblacion_6

  # ---------------------------------------------------------------------------
  # 8) Render + export (estricto con .PPT_CONTRACT)
  # ---------------------------------------------------------------------------
  log_rows <- list()
  rendered <- list()

  for (i in seq_along(plan)) {

    slide <- plan[[i]]
    if (!inherits(slide, "ppt_slide")) {
      stop("Cada slide debe tener clase `ppt_slide`.", call. = FALSE)
    }

    stype <- slide$.slide_type %||% NA_character_

    if (isTRUE(mensajes_progreso)) {
      .msg_diapo(
        i, length(plan), stype,
        el_plot = NULL,
        mensajes_progreso = mensajes_progreso
      )
    }

    # ---- TITLE SLIDE ---------------------------------------------------------
    if (identical(stype, "title_slide")) {

      contract <- PPT_CONTRACT$title_slide
      slots <- slide$slots %||% list()

      ttl  <- slots$title      %||% slide$title %||% NULL
      sub  <- slots$subtitle   %||% NULL
      dt   <- slots$date       %||% NULL
      ml   <- slots$meta_left  %||% NULL
      mr   <- slots$meta_right %||% NULL
      mln  <- slots$meta_line  %||% NULL

      if (!isTRUE(solo_lista)) {

        doc <- .add_slide_strict(doc, contract$layout)

        # title (requerido)
        if (!is.null(ttl) && nzchar(trimws(ttl))) {
          doc <- .ph_with_strict(doc, ttl, contract$slots$title)
        } else {
          stop("title_slide requiere `title` no vacío.", call. = FALSE)
        }

        # opcionales (solo si vienen)
        if (!is.null(sub) && nzchar(trimws(sub))) {
          doc <- .ph_with_strict(doc, sub, contract$slots$subtitle)
        }
        if (!is.null(dt) && nzchar(trimws(dt))) {
          doc <- .ph_with_strict(doc, dt, contract$slots$date)
        }
        if (!is.null(ml) && nzchar(trimws(ml))) {
          doc <- .ph_with_strict(doc, ml, contract$slots$meta_left)
        }
        if (!is.null(mr) && nzchar(trimws(mr))) {
          doc <- .ph_with_strict(doc, mr, contract$slots$meta_right)
        }
        if (!is.null(mln) && nzchar(trimws(mln))) {
          doc <- .ph_with_strict(doc, mln, contract$slots$meta_line)
        }
      }

      log_rows[[length(log_rows) + 1]] <- tibble::tibble(
        slide_i    = i,
        slide_type = "title_slide",
        element    = NA_character_,
        var        = NA_character_
      )
      next
    }

    # ---- SECTION -------------------------------------------------------------
    if (identical(stype, "section")) {

      contract <- PPT_CONTRACT$section
      title    <- slide$title %||% ""
      subtitle <- slide$subtitle %||% NULL

      if (!isTRUE(solo_lista)) {
        doc <- .add_slide_strict(doc, contract$layout)
        doc <- .ph_with_strict(doc, title, contract$slots$title)
        if (!is.null(subtitle) && nzchar(subtitle)) {
          doc <- .ph_with_strict(doc, subtitle, contract$slots$subtitle)
        }
      }

      if (isTRUE(mensajes_progreso)) {
        message(sprintf("  • sección: %s", slide$title %||% "<sin título>"))
      }

      log_rows[[length(log_rows) + 1]] <- tibble::tibble(
        slide_i    = i,
        slide_type = "section",
        element    = NA_character_,
        var        = NA_character_
      )
      next
    }

    # ---- SLIDE_1 -------------------------------------------------------------
    if (identical(stype, "slide_1")) {

      contract <- PPT_CONTRACT$slide_1

      title_slide <- slide$title %||% NULL
      slots       <- slide$slots %||% list()
      el_plot     <- slots$plot %||% NULL

      if (!inherits(el_plot, "ppt_element")) {
        stop("En `p_slide_1()`, `plot` debe ser `ppt_element`.", call. = FALSE)
      }

      etype <- el_plot$.element_type %||% NA_character_

      if (isTRUE(mensajes_progreso)) {
        .msg_diapo(i, length(plan), stype, el_plot = el_plot, mensajes_progreso = mensajes_progreso)
        message("  • gráficos a crear: 1")
      }

      p <- .render_element(el_plot)

      if (is.null(p)) {
        vv <- el_plot$var %||% paste(el_plot$vars %||% character(0), collapse = ", ")
        stop("No se pudo renderizar elemento: ", etype, " (", vv, ").", call. = FALSE)
      }

      rendered[[length(rendered) + 1]] <- p

      # Resolver título del slide si no viene
      if (is.null(title_slide)) {
        title_slide <- el_plot$title_slide %||% {
          if (!is.null(el_plot$var)) .title_of_var(el_plot$var) else {
            v1 <- el_plot$vars %||% NULL
            if (!is.null(v1) && length(v1)) .title_of_var(v1[1]) else NULL
          }
        }
      }

      if (!isTRUE(solo_lista)) {

        doc <- .add_slide_strict(doc, contract$layout)

        if (!is.null(title_slide) && nzchar(title_slide)) {
          doc <- .ph_with_strict(doc, title_slide, contract$slots$title)
        }

        doc <- .ph_with_strict(
          doc,
          rvg::dml(ggobj = p, bg = "transparent"),
          contract$slots$plot
        )

        # BASE (manual o auto)
        base_txt <- slots$base %||% NULL

        if (is.null(base_txt)) {
          var_base <- el_plot$var %||% {
            v1 <- el_plot$vars %||% NULL
            if (!is.null(v1) && length(v1)) v1[1] else NULL
          }
          base_txt <- .base_auto_from_var(
            var         = var_base,
            sufijo_auto = presets$base$args$sufijo_auto %||% NULL,
            formato     = presets$base$args$formato %||% "Base: %s"
          )
        }

        if (is.null(base_txt)) base_txt <- " "
        doc <- .ph_with_strict(doc, as.character(base_txt)[1], contract$slots$base)

        # RIGHT (usa footer o deja en blanco)
        right_obj <- slots$footer %||% NULL

        right_txt <- NULL
        if (inherits(right_obj, "ppt_element_text")) right_txt <- right_obj$text %||% NULL
        if (is.character(right_obj) && length(right_obj) == 1L) right_txt <- right_obj

        if (is.null(right_txt) || !nzchar(trimws(right_txt))) right_txt <- " "
        doc <- .ph_with_strict(doc, right_txt, contract$slots$right)
      }

      log_rows[[length(log_rows) + 1]] <- tibble::tibble(
        slide_i    = i,
        slide_type = "slide_1",
        element    = el_plot$.element_type %||% NA_character_,
        var        = el_plot$var %||% {
          v1 <- el_plot$vars %||% NULL
          if (!is.null(v1) && length(v1)) v1[1] else NA_character_
        }
      )
      next
    }

    # ---- SLIDE_2 -------------------------------------------------------------
    if (identical(stype, "slide_2")) {

      contract <- PPT_CONTRACT$slide_2

      title_slide <- slide$title %||% NULL
      slots       <- slide$slots %||% list()

      el_left  <- slots$left  %||% NULL
      el_right <- slots$right %||% NULL

      if (!inherits(el_left, "ppt_element") || !inherits(el_right, "ppt_element")) {
        stop("En `p_slide_2()`, `left` y `right` deben ser `ppt_element`.", call. = FALSE)
      }

      pL <- .render_element(el_left)
      pR <- .render_element(el_right)

      if (is.null(pL)) stop("No se pudo renderizar left: ",  el_left$.element_type  %||% "<NA>", call. = FALSE)
      if (is.null(pR)) stop("No se pudo renderizar right: ", el_right$.element_type %||% "<NA>", call. = FALSE)

      rendered[[length(rendered) + 1]] <- pL
      rendered[[length(rendered) + 1]] <- pR

      if (!isTRUE(solo_lista)) {

        doc <- .add_slide_strict(doc, contract$layout)

        if (!is.null(title_slide) && nzchar(title_slide)) {
          doc <- .ph_with_strict(doc, title_slide, contract$slots$title)
        }

        doc <- .ph_with_strict(doc, rvg::dml(ggobj = pL, bg = "transparent"), contract$slots$left)
        doc <- .ph_with_strict(doc, rvg::dml(ggobj = pR, bg = "transparent"), contract$slots$right)

        # BASE auto desde left si no se declara
        base_txt <- slots$base %||% NULL
        if (is.null(base_txt)) {
          base_txt <- .base_auto_from_var(
            var         = el_left$var,
            sufijo_auto = presets$base$args$sufijo_auto %||% NULL,
            formato     = presets$base$args$formato %||% "Base: %s"
          )
        }
        if (is.null(base_txt)) base_txt <- " "
        doc <- .ph_with_strict(doc, as.character(base_txt)[1], contract$slots$base)

        rt_txt <- slots$right_text %||% NULL
        if (!is.null(rt_txt) && is.character(rt_txt) && length(rt_txt) == 1L) {
          doc <- .ph_with_strict(doc, rt_txt, contract$slots$right_text)
        } else {
          doc <- .ph_with_strict(doc, " ", contract$slots$right_text)
        }
      }

      log_rows[[length(log_rows) + 1]] <- tibble::tibble(
        slide_i    = i,
        slide_type = "slide_2",
        element    = paste0(
          el_left$.element_type  %||% "<NA>", " + ",
          el_right$.element_type %||% "<NA>"
        ),
        var = paste0(
          (el_left$var  %||% paste(el_left$vars  %||% character(0), collapse = ",")),
          " | ",
          (el_right$var %||% paste(el_right$vars %||% character(0), collapse = ","))
        )
      )
      next
    }

    # ---- POBLACION_4 (4 gráficos 2x2) ----------------------------------------
    if (identical(stype, "poblacion_4")) {

      contract <- PPT_CONTRACT$poblacion_4
      slots    <- slide$slots %||% list()

      # título (opcional)
      title_slide <- slots$title %||% slide$title %||% NULL

      # elementos requeridos (4)
      el_ul <- slots$up_left      %||% NULL
      el_ur <- slots$up_right     %||% NULL
      el_bl <- slots$bottom_left  %||% NULL
      el_br <- slots$bottom_right %||% NULL

      if (!inherits(el_ul, "ppt_element")) stop("poblacion_4: `up_left` debe ser `ppt_element`.", call. = FALSE)
      if (!inherits(el_ur, "ppt_element")) stop("poblacion_4: `up_right` debe ser `ppt_element`.", call. = FALSE)
      if (!inherits(el_bl, "ppt_element")) stop("poblacion_4: `bottom_left` debe ser `ppt_element`.", call. = FALSE)
      if (!inherits(el_br, "ppt_element")) stop("poblacion_4: `bottom_right` debe ser `ppt_element`.", call. = FALSE)

      pUL <- .render_element(el_ul)
      pUR <- .render_element(el_ur)
      pBL <- .render_element(el_bl)
      pBR <- .render_element(el_br)

      if (is.null(pUL)) stop("poblacion_4: no se pudo renderizar up_left (",      el_ul$.element_type %||% "<NA>", ").", call. = FALSE)
      if (is.null(pUR)) stop("poblacion_4: no se pudo renderizar up_right (",     el_ur$.element_type %||% "<NA>", ").", call. = FALSE)
      if (is.null(pBL)) stop("poblacion_4: no se pudo renderizar bottom_left (",  el_bl$.element_type %||% "<NA>", ").", call. = FALSE)
      if (is.null(pBR)) stop("poblacion_4: no se pudo renderizar bottom_right (", el_br$.element_type %||% "<NA>", ").", call. = FALSE)

      rendered[[length(rendered) + 1]] <- pUL
      rendered[[length(rendered) + 1]] <- pUR
      rendered[[length(rendered) + 1]] <- pBL
      rendered[[length(rendered) + 1]] <- pBR

      if (!isTRUE(solo_lista)) {

        doc <- .add_slide_strict(doc, contract$layout)

        if (!is.null(title_slide) && nzchar(trimws(title_slide))) {
          doc <- .ph_with_strict(doc, title_slide, contract$slots$title)
        }

        doc <- .ph_with_strict(doc, rvg::dml(ggobj = pUL, bg = "transparent"), contract$slots$up_left)
        doc <- .ph_with_strict(doc, rvg::dml(ggobj = pUR, bg = "transparent"), contract$slots$up_right)
        doc <- .ph_with_strict(doc, rvg::dml(ggobj = pBL, bg = "transparent"), contract$slots$bottom_left)
        doc <- .ph_with_strict(doc, rvg::dml(ggobj = pBR, bg = "transparent"), contract$slots$bottom_right)

        # tag (usa body idx 1) — opcional
        tag_txt <- slots$tag %||% NULL
        if (!is.null(tag_txt) && is.character(tag_txt) && length(tag_txt) == 1L && nzchar(trimws(tag_txt))) {
          doc <- .ph_with_strict(doc, tag_txt, contract$slots$tag)
        }

        # center_note (usa body idx 2) — opcional
        cn_txt <- slots$center_note %||% NULL
        if (!is.null(cn_txt) && is.character(cn_txt) && length(cn_txt) == 1L && nzchar(trimws(cn_txt))) {
          doc <- .ph_with_strict(doc, cn_txt, contract$slots$center_note)
        }

        # base (usa body idx 3) — opcional/auto
        base_txt <- slots$base %||% NULL
        if (is.null(base_txt)) {
          # por defecto: intentar armar base desde el primer elemento
          var_base <- el_ul$var %||% {
            v1 <- el_ul$vars %||% NULL
            if (!is.null(v1) && length(v1)) v1[1] else NULL
          }
          base_txt <- .base_auto_from_var(
            var         = var_base,
            sufijo_auto = presets$base$args$sufijo_auto %||% NULL,
            formato     = presets$base$args$formato %||% "Base: %s"
          )
        }
        if (is.null(base_txt) || !nzchar(trimws(base_txt))) base_txt <- " "
        doc <- .ph_with_strict(doc, as.character(base_txt)[1], contract$slots$base)
      }

      log_rows[[length(log_rows) + 1]] <- tibble::tibble(
        slide_i    = i,
        slide_type = "poblacion_4",
        element    = paste(
          el_ul$.element_type %||% "<NA>",
          el_ur$.element_type %||% "<NA>",
          el_bl$.element_type %||% "<NA>",
          el_br$.element_type %||% "<NA>",
          sep = " | "
        ),
        var = paste(
          el_ul$var %||% paste(el_ul$vars %||% character(0), collapse = ","),
          el_ur$var %||% paste(el_ur$vars %||% character(0), collapse = ","),
          el_bl$var %||% paste(el_bl$vars %||% character(0), collapse = ","),
          el_br$var %||% paste(el_br$vars %||% character(0), collapse = ","),
          sep = " || "
        )
      )
      next
    }

    # ---- TEXT_R (gráfico izquierda, texto derecha) ------------------------------
    if (identical(stype, "text_r")) {

      contract <- PPT_CONTRACT$text_r
      slots    <- slide$slots %||% list()

      title_slide <- slots$title %||% slide$title %||% NULL

      el_plot <- slots$plot %||% NULL
      if (!inherits(el_plot, "ppt_element")) {
        stop("text_r: `plot` debe ser `ppt_element`.", call. = FALSE)
      }

      # render plot
      if (isTRUE(mensajes_progreso)) {
        .msg_diapo(i, length(plan), stype, el_plot = el_plot, mensajes_progreso = mensajes_progreso)
        message("  • gráficos a crear: 1")
      }

      p <- .render_element(el_plot)
      if (is.null(p)) {
        vv <- el_plot$var %||% paste(el_plot$vars %||% character(0), collapse = ", ")
        stop("text_r: no se pudo renderizar plot (", el_plot$.element_type %||% "<NA>", " | ", vv, ").", call. = FALSE)
      }
      rendered[[length(rendered) + 1]] <- p

      # inferir título si no viene
      if (is.null(title_slide)) {
        title_slide <- el_plot$title_slide %||% {
          if (!is.null(el_plot$var)) .title_of_var(el_plot$var) else {
            v1 <- el_plot$vars %||% NULL
            if (!is.null(v1) && length(v1)) .title_of_var(v1[1]) else NULL
          }
        }
      }

      if (!isTRUE(solo_lista)) {

        doc <- .add_slide_strict(doc, contract$layout)

        if (!is.null(title_slide) && nzchar(trimws(title_slide))) {
          doc <- .ph_with_strict(doc, title_slide, contract$slots$title)
        }

        # tag opcional
        tag_txt <- slots$tag %||% NULL
        if (!is.null(tag_txt) && nzchar(trimws(as.character(tag_txt)[1]))) {
          doc <- .ph_with_strict(doc, as.character(tag_txt)[1], contract$slots$tag)
        }

        # plot
        doc <- .ph_with_strict(
          doc,
          rvg::dml(ggobj = p, bg = "transparent"),
          contract$slots$plot
        )

        # texto derecha (slots$text es character(1) en tu p_slide_text_r)
        tx <- slots$text %||% NULL
        if (is.null(tx) || !nzchar(trimws(as.character(tx)[1]))) tx <- " "
        doc <- .ph_with_strict(doc, as.character(tx)[1], contract$slots$text)

        # base (manual o auto)
        base_txt <- slots$base %||% NULL
        if (is.null(base_txt)) {
          var_base <- el_plot$var %||% {
            v1 <- el_plot$vars %||% NULL
            if (!is.null(v1) && length(v1)) v1[1] else NULL
          }
          base_txt <- .base_auto_from_var(
            var         = var_base,
            sufijo_auto = presets$base$args$sufijo_auto %||% NULL,
            formato     = presets$base$args$formato %||% "Base: %s"
          )
        }
        if (is.null(base_txt) || !nzchar(trimws(as.character(base_txt)[1]))) base_txt <- " "
        doc <- .ph_with_strict(doc, as.character(base_txt)[1], contract$slots$base)

        # footer opcional
        ft <- slots$footer %||% NULL
        if (is.null(ft) || !nzchar(trimws(as.character(ft)[1]))) ft <- " "
        doc <- .ph_with_strict(doc, as.character(ft)[1], contract$slots$footer)
      }

      log_rows[[length(log_rows) + 1]] <- tibble::tibble(
        slide_i    = i,
        slide_type = "text_r",
        element    = el_plot$.element_type %||% NA_character_,
        var        = el_plot$var %||% {
          v1 <- el_plot$vars %||% NULL
          if (!is.null(v1) && length(v1)) v1[1] else NA_character_
        }
      )
      next
    }

    # ---- TEXT_L (texto izquierda, gráfico derecha) ------------------------------
    if (identical(stype, "text_l")) {

      contract <- PPT_CONTRACT$text_l
      slots    <- slide$slots %||% list()

      title_slide <- slots$title %||% slide$title %||% NULL

      el_plot <- slots$plot %||% NULL
      if (!inherits(el_plot, "ppt_element")) {
        stop("text_l: `plot` debe ser `ppt_element`.", call. = FALSE)
      }

      if (isTRUE(mensajes_progreso)) {
        .msg_diapo(i, length(plan), stype, el_plot = el_plot, mensajes_progreso = mensajes_progreso)
        message("  • gráficos a crear: 1")
      }

      p <- .render_element(el_plot)
      if (is.null(p)) {
        vv <- el_plot$var %||% paste(el_plot$vars %||% character(0), collapse = ", ")
        stop("text_l: no se pudo renderizar plot (", el_plot$.element_type %||% "<NA>", " | ", vv, ").", call. = FALSE)
      }
      rendered[[length(rendered) + 1]] <- p

      if (is.null(title_slide)) {
        title_slide <- el_plot$title_slide %||% {
          if (!is.null(el_plot$var)) .title_of_var(el_plot$var) else {
            v1 <- el_plot$vars %||% NULL
            if (!is.null(v1) && length(v1)) .title_of_var(v1[1]) else NULL
          }
        }
      }

      if (!isTRUE(solo_lista)) {

        doc <- .add_slide_strict(doc, contract$layout)

        if (!is.null(title_slide) && nzchar(trimws(title_slide))) {
          doc <- .ph_with_strict(doc, title_slide, contract$slots$title)
        }

        # tag opcional
        tag_txt <- slots$tag %||% NULL
        if (!is.null(tag_txt) && nzchar(trimws(as.character(tag_txt)[1]))) {
          doc <- .ph_with_strict(doc, as.character(tag_txt)[1], contract$slots$tag)
        }

        # texto izquierda
        tx <- slots$text %||% NULL
        if (is.null(tx) || !nzchar(trimws(as.character(tx)[1]))) tx <- " "
        doc <- .ph_with_strict(doc, as.character(tx)[1], contract$slots$text)

        # plot derecha
        doc <- .ph_with_strict(
          doc,
          rvg::dml(ggobj = p, bg = "transparent"),
          contract$slots$plot
        )

        # base (manual o auto)
        base_txt <- slots$base %||% NULL
        if (is.null(base_txt)) {
          var_base <- el_plot$var %||% {
            v1 <- el_plot$vars %||% NULL
            if (!is.null(v1) && length(v1)) v1[1] else NULL
          }
          base_txt <- .base_auto_from_var(
            var         = var_base,
            sufijo_auto = presets$base$args$sufijo_auto %||% NULL,
            formato     = presets$base$args$formato %||% "Base: %s"
          )
        }
        if (is.null(base_txt) || !nzchar(trimws(as.character(base_txt)[1]))) base_txt <- " "
        doc <- .ph_with_strict(doc, as.character(base_txt)[1], contract$slots$base)

        # footer opcional
        ft <- slots$footer %||% NULL
        if (is.null(ft) || !nzchar(trimws(as.character(ft)[1]))) ft <- " "
        doc <- .ph_with_strict(doc, as.character(ft)[1], contract$slots$footer)
      }

      log_rows[[length(log_rows) + 1]] <- tibble::tibble(
        slide_i    = i,
        slide_type = "text_l",
        element    = el_plot$.element_type %||% NA_character_,
        var        = el_plot$var %||% {
          v1 <- el_plot$vars %||% NULL
          if (!is.null(v1) && length(v1)) v1[1] else NA_character_
        }
      )
      next
    }

    if (identical(stype, "text_r2")) {

      contract <- PPT_CONTRACT$text_r2
      slots    <- slide$slots %||% list()

      title_slide <- slots$title %||% slide$title %||% NULL

      el1 <- slots$plot1 %||% NULL
      el2 <- slots$plot2 %||% NULL

      if (!inherits(el1, "ppt_element")) stop("text_r2: `plot1` debe ser `ppt_element`.", call. = FALSE)
      if (!inherits(el2, "ppt_element")) stop("text_r2: `plot2` debe ser `ppt_element`.", call. = FALSE)

      p1 <- .render_element(el1)
      p2 <- .render_element(el2)

      if (is.null(p1)) stop("text_r2: no se pudo renderizar plot1.", call. = FALSE)
      if (is.null(p2)) stop("text_r2: no se pudo renderizar plot2.", call. = FALSE)

      rendered[[length(rendered) + 1]] <- p1
      rendered[[length(rendered) + 1]] <- p2

      # inferir título si no viene
      if (is.null(title_slide)) {
        title_slide <- el1$title_slide %||% if (!is.null(el1$var)) .title_of_var(el1$var) else NULL
      }

      if (!isTRUE(solo_lista)) {

        doc <- .add_slide_strict(doc, contract$layout)

        if (!is.null(title_slide) && nzchar(trimws(title_slide))) {
          doc <- .ph_with_strict(doc, title_slide, contract$slots$title)
        }

        # tag opcional
        tag_txt <- slots$tag %||% NULL
        if (!is.null(tag_txt) && nzchar(trimws(as.character(tag_txt)[1]))) {
          doc <- .ph_with_strict(doc, as.character(tag_txt)[1], contract$slots$tag)
        }

        # 2 plots
        doc <- .ph_with_strict(doc, rvg::dml(ggobj = p1, bg = "transparent"), contract$slots$plot1)
        doc <- .ph_with_strict(doc, rvg::dml(ggobj = p2, bg = "transparent"), contract$slots$plot2)

        # texto derecha
        tx <- slots$text %||% NULL
        if (is.null(tx) || !nzchar(trimws(as.character(tx)[1]))) tx <- " "
        doc <- .ph_with_strict(doc, as.character(tx)[1], contract$slots$text)

        # base auto (por defecto desde plot1)
        base_txt <- slots$base %||% NULL
        if (is.null(base_txt)) {
          var_base <- el1$var %||% { v1 <- el1$vars %||% NULL; if (!is.null(v1) && length(v1)) v1[1] else NULL }
          base_txt <- .base_auto_from_var(
            var         = var_base,
            sufijo_auto = presets$base$args$sufijo_auto %||% NULL,
            formato     = presets$base$args$formato %||% "Base: %s"
          )
        }
        if (is.null(base_txt) || !nzchar(trimws(as.character(base_txt)[1]))) base_txt <- " "
        doc <- .ph_with_strict(doc, as.character(base_txt)[1], contract$slots$base)

        # footer opcional
        ft <- slots$footer %||% NULL
        if (is.null(ft) || !nzchar(trimws(as.character(ft)[1]))) ft <- " "
        doc <- .ph_with_strict(doc, as.character(ft)[1], contract$slots$footer)
      }

      next
    }

    if (identical(stype, "text_l2")) {

      contract <- PPT_CONTRACT$text_l2
      slots    <- slide$slots %||% list()

      title_slide <- slots$title %||% slide$title %||% NULL

      el1 <- slots$plot1 %||% NULL
      el2 <- slots$plot2 %||% NULL

      if (!inherits(el1, "ppt_element")) stop("text_l2: `plot1` debe ser `ppt_element`.", call. = FALSE)
      if (!inherits(el2, "ppt_element")) stop("text_l2: `plot2` debe ser `ppt_element`.", call. = FALSE)

      p1 <- .render_element(el1)
      p2 <- .render_element(el2)

      if (is.null(p1)) stop("text_l2: no se pudo renderizar plot1.", call. = FALSE)
      if (is.null(p2)) stop("text_l2: no se pudo renderizar plot2.", call. = FALSE)

      rendered[[length(rendered) + 1]] <- p1
      rendered[[length(rendered) + 1]] <- p2

      if (is.null(title_slide)) {
        title_slide <- el1$title_slide %||% if (!is.null(el1$var)) .title_of_var(el1$var) else NULL
      }

      if (!isTRUE(solo_lista)) {

        doc <- .add_slide_strict(doc, contract$layout)

        if (!is.null(title_slide) && nzchar(trimws(title_slide))) {
          doc <- .ph_with_strict(doc, title_slide, contract$slots$title)
        }

        # tag opcional
        tag_txt <- slots$tag %||% NULL
        if (!is.null(tag_txt) && nzchar(trimws(as.character(tag_txt)[1]))) {
          doc <- .ph_with_strict(doc, as.character(tag_txt)[1], contract$slots$tag)
        }

        # texto izquierda
        tx <- slots$text %||% NULL
        if (is.null(tx) || !nzchar(trimws(as.character(tx)[1]))) tx <- " "
        doc <- .ph_with_strict(doc, as.character(tx)[1], contract$slots$text)

        # 2 plots
        doc <- .ph_with_strict(doc, rvg::dml(ggobj = p1, bg = "transparent"), contract$slots$plot1)
        doc <- .ph_with_strict(doc, rvg::dml(ggobj = p2, bg = "transparent"), contract$slots$plot2)

        # base auto desde plot1
        base_txt <- slots$base %||% NULL
        if (is.null(base_txt)) {
          var_base <- el1$var %||% { v1 <- el1$vars %||% NULL; if (!is.null(v1) && length(v1)) v1[1] else NULL }
          base_txt <- .base_auto_from_var(
            var         = var_base,
            sufijo_auto = presets$base$args$sufijo_auto %||% NULL,
            formato     = presets$base$args$formato %||% "Base: %s"
          )
        }
        if (is.null(base_txt) || !nzchar(trimws(as.character(base_txt)[1]))) base_txt <- " "
        doc <- .ph_with_strict(doc, as.character(base_txt)[1], contract$slots$base)

        # footer opcional
        ft <- slots$footer %||% NULL
        if (is.null(ft) || !nzchar(trimws(as.character(ft)[1]))) ft <- " "
        doc <- .ph_with_strict(doc, as.character(ft)[1], contract$slots$footer)
      }

      next
    }

    # ---- POBLACION_2 ------------------------------------------------------------
    if (identical(stype, "poblacion_2")) {

      contract <- PPT_CONTRACT$poblacion_2
      slots    <- slide$slots %||% list()

      title_slide <- slots$title %||% slide$title %||% NULL
      tag_txt     <- slots$tag   %||% NULL

      el_left  <- slots$left  %||% NULL
      el_right <- slots$right %||% NULL

      if (!inherits(el_left, "ppt_element"))  stop("poblacion_2: `left` debe ser `ppt_element`.", call. = FALSE)
      if (!inherits(el_right, "ppt_element")) stop("poblacion_2: `right` debe ser `ppt_element`.", call. = FALSE)

      pL <- .render_element(el_left)
      pR <- .render_element(el_right)

      if (is.null(pL)) stop("poblacion_2: no se pudo renderizar left.", call. = FALSE)
      if (is.null(pR)) stop("poblacion_2: no se pudo renderizar right.", call. = FALSE)

      rendered[[length(rendered) + 1]] <- pL
      rendered[[length(rendered) + 1]] <- pR

      if (!isTRUE(solo_lista)) {

        doc <- .add_slide_strict(doc, contract$layout)

        if (!is.null(title_slide) && nzchar(trimws(as.character(title_slide)[1]))) {
          doc <- .ph_with_strict(doc, as.character(title_slide)[1], contract$slots$title)
        }

        if (!is.null(tag_txt) && nzchar(trimws(as.character(tag_txt)[1]))) {
          doc <- .ph_with_strict(doc, as.character(tag_txt)[1], contract$slots$tag)
        }

        # OJO: aquí en tu contrato left/right son "body", pero tú quieres meter gráficos:
        # si realmente son placeholders de texto, esto debe ser texto.
        # Si son placeholders de imagen, cambia el contrato a type="pic".
        doc <- .ph_with_strict(doc, rvg::dml(ggobj = pL, bg = "transparent"), contract$slots$left)
        doc <- .ph_with_strict(doc, rvg::dml(ggobj = pR, bg = "transparent"), contract$slots$right)
      }

      next
    }

    # ---- POBLACION_5 ------------------------------------------------------------
    if (identical(stype, "poblacion_5")) {

      contract <- PPT_CONTRACT$poblacion_5
      slots    <- slide$slots %||% list()

      title_slide <- slots$title %||% slide$title %||% NULL

      pics <- lapply(1:5, function(i) slots[[paste0("pic", i)]] %||% NULL)
      for (i in 1:5) if (!inherits(pics[[i]], "ppt_element")) stop("poblacion_5: `pic", i, "` debe ser `ppt_element`.", call. = FALSE)

      plots <- lapply(pics, .render_element)
      for (i in 1:5) if (is.null(plots[[i]])) stop("poblacion_5: no se pudo renderizar pic", i, ".", call. = FALSE)

      rendered <- c(rendered, plots)

      if (!isTRUE(solo_lista)) {

        doc <- .add_slide_strict(doc, contract$layout)

        if (!is.null(title_slide) && nzchar(trimws(as.character(title_slide)[1]))) {
          doc <- .ph_with_strict(doc, as.character(title_slide)[1], contract$slots$title)
        }

        # tag/icon/footer opcionales
        for (nm in c("tag","icon","footer")) {
          tx <- slots[[nm]] %||% NULL
          if (!is.null(tx) && nzchar(trimws(as.character(tx)[1]))) {
            doc <- .ph_with_strict(doc, as.character(tx)[1], contract$slots[[nm]])
          }
        }

        # 5 pics
        for (i in 1:5) {
          doc <- .ph_with_strict(
            doc,
            rvg::dml(ggobj = plots[[i]], bg = "transparent"),
            contract$slots[[paste0("pic", i)]]
          )
        }
      }

      next
    }

    # ---- POBLACION_6 ------------------------------------------------------------
    if (identical(stype, "poblacion_6")) {

      contract <- PPT_CONTRACT$poblacion_6
      slots    <- slide$slots %||% list()

      title_slide <- slots$title %||% slide$title %||% NULL

      pics <- lapply(1:6, function(i) slots[[paste0("pic", i)]] %||% NULL)
      for (i in 1:6) if (!inherits(pics[[i]], "ppt_element")) stop("poblacion_6: `pic", i, "` debe ser `ppt_element`.", call. = FALSE)

      plots <- lapply(pics, .render_element)
      for (i in 1:6) if (is.null(plots[[i]])) stop("poblacion_6: no se pudo renderizar pic", i, ".", call. = FALSE)

      rendered <- c(rendered, plots)

      if (!isTRUE(solo_lista)) {

        doc <- .add_slide_strict(doc, contract$layout)

        if (!is.null(title_slide) && nzchar(trimws(as.character(title_slide)[1]))) {
          doc <- .ph_with_strict(doc, as.character(title_slide)[1], contract$slots$title)
        }

        for (nm in c("tag","icon","footer")) {
          tx <- slots[[nm]] %||% NULL
          if (!is.null(tx) && nzchar(trimws(as.character(tx)[1]))) {
            doc <- .ph_with_strict(doc, as.character(tx)[1], contract$slots[[nm]])
          }
        }

        for (i in 1:6) {
          doc <- .ph_with_strict(
            doc,
            rvg::dml(ggobj = plots[[i]], bg = "transparent"),
            contract$slots[[paste0("pic", i)]]
          )
        }
      }

      next
    }

    stop("Tipo de slide no implementado: ", stype, call. = FALSE)
  }

  log <- dplyr::bind_rows(log_rows)

  if (!isTRUE(solo_lista)) {
    print(doc, target = path_ppt)
    if (isTRUE(mensajes_progreso)) {
      message("PPT generado en: ", normalizePath(path_ppt, winslash = "/"))
    }
  }

  # ---------------------------------------------------------------------------
  # Limpiar plan acumulado (si se usó diapo())
  # ---------------------------------------------------------------------------
  if (exists(".ppt_plan_clear", mode = "function", inherits = TRUE)) {
    try(.ppt_plan_clear(env_diapos), silent = TRUE)
  }

  invisible(list(
    doc      = if (isTRUE(solo_lista)) NULL else doc,
    plan     = plan,
    rendered = rendered,
    log      = log
  ))
}

# =============================================================================
# PRESETS
# =============================================================================

#' @title Definir presets por tipo de elemento y por layout
#'
#' @description
#' Construye un objeto de presets que centraliza configuraciones por tipo:
#' `barras_agrupadas`, `barras_apiladas`, `multi_apiladas`, `pie`, `donut`,
#' `numerico`, `texto`, `base`, y también configuración de layouts/slots.
#'
#' Los presets se definen en un chunk previo y se pasan a `reporte_ppt_plan(presets=...)`.
#'
#' @param barras_agrupadas Lista de parámetros por defecto para `graficar_barras_agrupadas()`.
#' @param barras_apiladas Lista de parámetros por defecto para `graficar_barras_apiladas()`.
#' @param multi_apiladas Lista de parámetros por defecto para `graficar_barras_apiladas()` en modo bloque.
#' @param pie Lista de parámetros por defecto para `graficar_pie(tipo_pie="pie")`.
#' @param donut Lista de parámetros por defecto para `graficar_pie(tipo_pie="donut")`.
#' @param numerico Lista de parámetros por defecto para elementos KPI numéricos.
#' @param texto Lista de parámetros por defecto para `p_text()` (tipografía, color, etc.).
#' @param base Lista de parámetros por defecto para `p_base()` (formato, prefijos, etc.).
#'
#' @param layouts Lista opcional con mapeos de layout -> slots -> placeholders
#'   (type/type_idx o loc). Si es `NULL`, se usan defaults internos.
#'
#' @param debug Lista opcional de parámetros de depuración (p.ej. bordes canvas).
#'
#' @return Objeto con clase `"ppt_presets"`.
#'
#' @export
p_presets <- function(
    barras_agrupadas = list(),
    barras_apiladas  = list(),
    multi_apiladas   = list(),
    pie              = list(),
    donut            = list(),
    numerico         = list(),
    texto            = list(),
    base             = list(),
    layouts          = NULL,
    debug            = list()
) {
  stop("Implementación pendiente.")
}

