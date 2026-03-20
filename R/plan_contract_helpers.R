# CONTRATO DE LAYPUT Y PLACERHOLDER DE PPT

#' @keywords internal
.PPT_CONTRACT <- list(

  # ------------------------------------------------------------
  # SECTION
  # ------------------------------------------------------------
  section = list(
    layout = "Section Header",
    slots  = list(
      title    = list(type="title", type_idx=NULL),
      subtitle = list(type="body",  type_idx=1)
    )
  ),

  # ------------------------------------------------------------
  # TITLE SLIDE  (nuevo)
  # ------------------------------------------------------------
  title_slide = list(
    layout = "Title Slide",
    slots  = list(
      title      = list(type="ctrTitle", type_idx=1),
      subtitle   = list(type="subTitle", type_idx=1),
      date       = list(type="dt",       type_idx=1),
      meta_left  = list(type="body",     type_idx=2),
      meta_right = list(type="body",     type_idx=1),
      meta_line  = list(type="body",     type_idx=3)
    )
  ),

  # ------------------------------------------------------------
  # SLIDE 1 (1 gráfico)
  # ------------------------------------------------------------
  slide_1 = list(
    layout = "Graficos",
    slots  = list(
      title = list(type="title", type_idx=NULL),
      plot  = list(type="pic",   type_idx=NULL),
      base  = list(type="body",  type_idx=2),
      right = list(type="body",  type_idx=3)
    )
  ),

  # ------------------------------------------------------------
  # SLIDE 2 (2 gráficos)
  # ------------------------------------------------------------
  slide_2 = list(
    layout = "Graficos_2columnas",
    slots  = list(
      title      = list(type="title", type_idx=NULL),
      left       = list(type="pic",   type_idx=2),
      right      = list(type="pic",   type_idx=1),
      base       = list(type="body",  type_idx=2),
      right_text = list(type="body",  type_idx=3)
    )
  ),

  # ------------------------------------------------------------
  # POBLACION_4 (nuevo) — 4 gráficos 2x2 con nombres posicionales
  # ------------------------------------------------------------
  poblacion_4 = list(
    layout = "poblacion_4",
    slots  = list(
      title        = list(type = "title", type_idx = 1),

      up_left      = list(type = "pic",   type_idx = 1),
      up_right     = list(type = "pic",   type_idx = 2),
      bottom_left  = list(type = "pic",   type_idx = 3),
      bottom_right = list(type = "pic",   type_idx = 4),

      # textos disponibles en el layout
      tag          = list(type = "body",  type_idx = 1),
      center_note  = list(type = "body",  type_idx = 2),
      base         = list(type = "body",  type_idx = 3)
    )
  ),

  # ------------------------------------------------------------
  # POBLACION_2 (nuevo) — 2 paneles grandes (body/body) + tag
  # OJO: tu layout_properties mostró:
  #  - title: type="title", type_idx=1
  #  - tag:   body 1 (shape lateral)
  #  - left:  body 2
  #  - right: body 3
  #  - body 4 = Imagen 14 (decorativo; se ignora)
  # ------------------------------------------------------------
  poblacion_2 = list(
    layout = "poblacion_2",
    slots  = list(
      title = list(type = "title", type_idx = 1),
      tag   = list(type = "body",  type_idx = 1),
      left  = list(type = "body",  type_idx = 2),
      right = list(type = "body",  type_idx = 3)
    )
  ),

  # ------------------------------------------------------------
  # POBLACION_5 (nuevo) — 5 pics + footer + tag (+ icon opcional)
  # layout_properties:
  #  - title: title 1
  #  - tag:   body 1 (shape)
  #  - icon:  body 2 (Imagen 14)
  #  - footer:body 3 (Text Placeholder 9)
  #  - pics:  pic 1..5
  # ------------------------------------------------------------
  poblacion_5 = list(
    layout = "poblacion_5",
    slots  = list(
      title  = list(type = "title", type_idx = 1),
      tag    = list(type = "body",  type_idx = 1),
      icon   = list(type = "body",  type_idx = 2),
      footer = list(type = "body",  type_idx = 3),

      pic1   = list(type = "pic",   type_idx = 1),
      pic2   = list(type = "pic",   type_idx = 2),
      pic3   = list(type = "pic",   type_idx = 3),
      pic4   = list(type = "pic",   type_idx = 4),
      pic5   = list(type = "pic",   type_idx = 5)
    )
  ),

  # ------------------------------------------------------------
  # POBLACION_6 (nuevo) — 6 pics + footer + tag (+ icon opcional)
  # layout_properties:
  #  - title: title 1
  #  - tag:   body 1 (shape)
  #  - icon:  body 2 (Imagen 14)
  #  - footer:body 3 (Text Placeholder 9)
  #  - pics:  pic 1..6
  # ------------------------------------------------------------
  poblacion_6 = list(
    layout = "poblacion_6",
    slots  = list(
      title  = list(type = "title", type_idx = 1),
      tag    = list(type = "body",  type_idx = 1),
      icon   = list(type = "body",  type_idx = 2),
      footer = list(type = "body",  type_idx = 3),

      pic1   = list(type = "pic",   type_idx = 1),
      pic2   = list(type = "pic",   type_idx = 2),
      pic3   = list(type = "pic",   type_idx = 3),
      pic4   = list(type = "pic",   type_idx = 4),
      pic5   = list(type = "pic",   type_idx = 5),
      pic6   = list(type = "pic",   type_idx = 6)
    )
  ),

  # ------------------------------------------------------------
  # GRAFICO + TEXTO (nuevo) — gráfico izquierda, texto derecha
  # ------------------------------------------------------------
  text_r = list(
    layout = "right_grafico_texto",
    slots  = list(
      title  = list(type="title", type_idx=1),
      tag    = list(type="body",  type_idx=1),
      plot   = list(type="pic",   type_idx=1),
      text   = list(type="body",  type_idx=2),
      base   = list(type="body",  type_idx=3),
      footer = list(type="body",  type_idx=4)
    )
  ),

  # ------------------------------------------------------------
  # GRAFICO + TEXTO (nuevo) — texto izquierda, gráfico derecha
  # ------------------------------------------------------------
  text_l = list(
    layout = "left_grafico_texto",
    slots  = list(
      title  = list(type="title", type_idx=1),
      tag    = list(type="body",  type_idx=1),
      text   = list(type="body",  type_idx=2),
      plot   = list(type="pic",   type_idx=1),
      base   = list(type="body",  type_idx=3),
      footer = list(type="body",  type_idx=4)
    )
  ),
  # ------------------------------------------------------------
  # 2 GRAFICOS + TEXTO — 2 gráficos + texto a la derecha
  # ------------------------------------------------------------
  text_r2 = list(
    layout = "right_2graficos_texto",   # <- nombre de layout en PPT
    slots  = list(
      title  = list(type="title", type_idx=1),
      tag    = list(type="body",  type_idx=1),

      plot1  = list(type="pic",   type_idx=1),
      plot2  = list(type="pic",   type_idx=2),

      text   = list(type="body",  type_idx=2),
      base   = list(type="body",  type_idx=3),
      footer = list(type="body",  type_idx=4)
    )
  ),

  # ------------------------------------------------------------
  # 2 GRAFICOS + TEXTO — texto a la izquierda + 2 gráficos
  # ------------------------------------------------------------
  text_l2 = list(
    layout = "left_2graficos_texto",    # <- nombre de layout en PPT
    slots  = list(
      title  = list(type="title", type_idx=1),
      tag    = list(type="body",  type_idx=1),

      text   = list(type="body",  type_idx=2),

      plot1  = list(type="pic",   type_idx=1),
      plot2  = list(type="pic",   type_idx=2),

      base   = list(type="body",  type_idx=3),
      footer = list(type="body",  type_idx=4)
    )
  )

)

# =============================================================================
# HELPERS internos (recolección, construcción y validación)
# - MVP: p_slide_section() + p_slide_1() + p_barras_apiladas()
# - Se asume que los slides tienen clase "ppt_slide" y campo .slide_type
# =============================================================================

#' @keywords internal
.collect_diapo_objects <- function(env = parent.frame(), strict = FALSE) {

  if (!is.environment(env)) {
    stop("`.collect_diapo_objects()`: `env` debe ser un environment.", call. = FALSE)
  }

  nms <- ls(envir = env, all.names = TRUE)
  nms <- nms[grepl("^diapo_\\d{3}$", nms)]

  if (!length(nms)) {
    return(list())
  }

  # ordenar por número
  ids <- as.integer(sub("^diapo_(\\d{3})$", "\\1", nms))
  ord <- order(ids)
  nms <- nms[ord]
  ids <- ids[ord]

  # recuperar objetos SIN heredar (para evitar colisiones raras)
  objs <- mget(nms, envir = env, inherits = FALSE)

  # validación ligera: clase ppt_slide
  bad <- vapply(objs, function(x) !inherits(x, "ppt_slide"), logical(1))
  if (any(bad)) {
    msg <- paste0(
      "`.collect_diapo_objects()`: estos objetos `diapo_###` no son `ppt_slide`: ",
      paste(names(objs)[bad], collapse = ", ")
    )
    if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
  }

  # strict: consecutividad (si hay >1)
  if (isTRUE(strict) && length(ids) > 1) {
    dif <- diff(ids)
    if (any(dif != 1L)) {
      stop(
        "strict=TRUE: los `diapo_###` no son consecutivos (hay saltos en la numeración).",
        call. = FALSE
      )
    }
  }

  objs
}

#' @keywords internal
.validate_plan <- function(plan, strict = FALSE) {

  if (!is.list(plan)) {
    stop("`.validate_plan()`: `plan` debe ser una lista de slides.", call. = FALSE)
  }
  if (!length(plan)) return(invisible(TRUE))

  bad_slide <- vapply(plan, function(x) !inherits(x, "ppt_slide"), logical(1))
  if (any(bad_slide)) {
    msg <- paste0(
      "`.validate_plan()`: hay elementos del plan que no son `ppt_slide` en posiciones: ",
      paste(which(bad_slide), collapse = ", ")
    )
    if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
  }

  for (i in seq_along(plan)) {
    s <- plan[[i]]
    if (!inherits(s, "ppt_slide")) next

    stype <- s$.slide_type %||% NA_character_

    # ---- SECTION ------------------------------------------------------------
    if (identical(stype, "section")) {

      ttl <- s$title %||% NULL
      ok  <- !is.null(ttl) && is.character(ttl) && length(ttl) == 1L && nzchar(trimws(ttl))

      if (!ok) {
        msg <- paste0("`.validate_plan()`: section (i=", i, ") requiere `title` no vacío.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
      }

      # subtitle es opcional; si existe debe ser character(1)
      sub <- s$subtitle %||% NULL
      if (!is.null(sub) && !(is.character(sub) && length(sub) == 1L)) {
        msg <- paste0("`.validate_plan()`: section (i=", i, ") `subtitle` debe ser character(1) o NULL.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
      }

      next
    }

    # -------------------------
    # TITLE SLIDE (nuevo)
    # -------------------------
    if (identical(stype, "title_slide")) {
      ttl <- s$slots$title %||% s$title %||% NULL
      ok  <- !is.null(ttl) && is.character(ttl) && length(ttl) == 1L && nzchar(trimws(ttl))
      if (!ok) {
        msg <- paste0("`.validate_plan()`: title_slide (i=", i, ") requiere `title` no vacío.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
      }
      next
    }

    # -------------------------
    # SLIDE_1
    # -------------------------
    if (identical(stype, "slide_1")) {
      slots <- s$slots %||% NULL
      if (is.null(slots) || !is.list(slots)) {
        msg <- paste0("`.validate_plan()`: slide_1 (i=", i, ") requiere `slots` como lista.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
        next
      }
      pl <- slots$plot %||% NULL
      if (is.null(pl) || !inherits(pl, "ppt_element")) {
        msg <- paste0("`.validate_plan()`: slide_1 (i=", i, ") requiere `slots$plot` como `ppt_element`.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
      }
      next
    }

    # -------------------------
    # SLIDE_2
    # -------------------------
    if (identical(stype, "slide_2")) {
      slots <- s$slots %||% NULL
      if (is.null(slots) || !is.list(slots)) {
        msg <- paste0("`.validate_plan()`: slide_2 (i=", i, ") requiere `slots` como lista.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
        next
      }
      el_left  <- slots$left  %||% NULL
      el_right <- slots$right %||% NULL
      if (is.null(el_left) || !inherits(el_left, "ppt_element")) {
        msg <- paste0("`.validate_plan()`: slide_2 (i=", i, ") requiere `slots$left` como `ppt_element`.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
      }
      if (is.null(el_right) || !inherits(el_right, "ppt_element")) {
        msg <- paste0("`.validate_plan()`: slide_2 (i=", i, ") requiere `slots$right` como `ppt_element`.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
      }
      next
    }

    # -------------------------
    # POBLACION_4
    # -------------------------
    if (identical(stype, "poblacion_4")) {
      slots <- s$slots %||% NULL
      if (is.null(slots) || !is.list(slots)) {
        msg <- paste0("`.validate_plan()`: poblacion_4 (i=", i, ") requiere `slots` como lista.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
        next
      }
      need <- c("up_left","up_right","bottom_left","bottom_right")
      for (nm in need) {
        el <- slots[[nm]] %||% NULL
        if (is.null(el) || !inherits(el, "ppt_element")) {
          msg <- paste0("`.validate_plan()`: poblacion_4 (i=", i, ") requiere `slots$", nm, "` como `ppt_element`.")
          if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
        }
      }
      next
    }

    # -------------------------
    # TEXT_R / TEXT_L
    # -------------------------
    if (identical(stype, "text_r") || identical(stype, "text_l")) {
      slots <- s$slots %||% NULL
      if (is.null(slots) || !is.list(slots)) {
        msg <- paste0("`.validate_plan()`: ", stype, " (i=", i, ") requiere `slots` como lista.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
        next
      }
      el_plot <- slots$plot %||% NULL
      if (is.null(el_plot) || !inherits(el_plot, "ppt_element")) {
        msg <- paste0("`.validate_plan()`: ", stype, " (i=", i, ") requiere `slots$plot` como `ppt_element`.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
      }
      # texto puede ser character(1) (lo insertas en PPT)
      tx <- slots$text %||% NULL
      if (!is.null(tx) && !(is.character(tx) && length(tx) == 1L)) {
        msg <- paste0("`.validate_plan()`: ", stype, " (i=", i, ") `slots$text` debe ser character(1) o NULL.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
      }
      next
    }

    # -------------------------
    # POBLACION_2
    # -------------------------
    if (identical(stype, "poblacion_2")) {
      slots <- s$slots %||% NULL
      if (is.null(slots) || !is.list(slots)) {
        msg <- paste0("`.validate_plan()`: poblacion_2 (i=", i, ") requiere `slots` como lista.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
        next
      }
      # Ajusta los nombres si tu layout usa otros
      need <- c("left", "right")
      for (nm in need) {
        el <- slots[[nm]] %||% NULL
        if (is.null(el) || !inherits(el, "ppt_element")) {
          msg <- paste0("`.validate_plan()`: poblacion_2 (i=", i, ") requiere `slots$", nm, "` como `ppt_element`.")
          if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
        }
      }
      next
    }

    # -------------------------
    # POBLACION_5
    # -------------------------
    if (identical(stype, "poblacion_5")) {
      slots <- s$slots %||% NULL
      if (is.null(slots) || !is.list(slots)) {
        msg <- paste0("`.validate_plan()`: poblacion_5 (i=", i, ") requiere `slots` como lista.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
        next
      }
      need <- paste0("pic", 1:5)
      for (nm in need) {
        el <- slots[[nm]] %||% NULL
        if (is.null(el) || !inherits(el, "ppt_element")) {
          msg <- paste0("`.validate_plan()`: poblacion_5 (i=", i, ") requiere `slots$", nm, "` como `ppt_element`.")
          if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
        }
      }
      next
    }

    # -------------------------
    # POBLACION_6
    # -------------------------
    if (identical(stype, "poblacion_6")) {
      slots <- s$slots %||% NULL
      if (is.null(slots) || !is.list(slots)) {
        msg <- paste0("`.validate_plan()`: poblacion_6 (i=", i, ") requiere `slots` como lista.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
        next
      }
      need <- paste0("pic", 1:6)
      for (nm in need) {
        el <- slots[[nm]] %||% NULL
        if (is.null(el) || !inherits(el, "ppt_element")) {
          msg <- paste0("`.validate_plan()`: poblacion_6 (i=", i, ") requiere `slots$", nm, "` como `ppt_element`.")
          if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
        }
      }
      next
    }

    # -------------------------
    # TEXT_R2 / TEXT_L2
    # -------------------------
    if (identical(stype, "text_r2") || identical(stype, "text_l2")) {
      slots <- s$slots %||% NULL
      if (is.null(slots) || !is.list(slots)) {
        msg <- paste0("`.validate_plan()`: ", stype, " (i=", i, ") requiere `slots` como lista.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
        next
      }
      for (nm in c("plot1","plot2")) {
        el <- slots[[nm]] %||% NULL
        if (is.null(el) || !inherits(el, "ppt_element")) {
          msg <- paste0("`.validate_plan()`: ", stype, " (i=", i, ") requiere `slots$", nm, "` como `ppt_element`.")
          if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
        }
      }
      tx <- slots$text %||% NULL
      if (!is.null(tx) && !(is.character(tx) && length(tx) == 1L)) {
        msg <- paste0("`.validate_plan()`: ", stype, " (i=", i, ") `slots$text` debe ser character(1) o NULL.")
        if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
      }
      next
    }


    # -------------------------
    # default
    # -------------------------
    msg <- paste0(
      "`.validate_plan()`: slide type no soportado (i=", i, "): ",
      if (is.na(stype)) "<NA>" else stype
    )
    if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
  }
}

#' @keywords internal
.merge_args <- function(...) {
  out <- list()
  for (lst in list(...)) {
    if (is.null(lst) || !length(lst)) next
    if (is.null(names(lst)) || any(names(lst) == "")) {
      stop("Todos los args deben venir nombrados (sin nombres vacíos).", call. = FALSE)
    }
    out[names(lst)] <- lst
  }
  out
}
#' @keywords internal
.keep_formals <- function(fun, args) {
  fml <- names(formals(fun))
  if ("..." %in% fml) return(args)
  args[names(args) %in% fml]
}

# -----------------------------------------------------------------------------
# LOG helpers (mensajes de progreso)
# -----------------------------------------------------------------------------
.fmt_vars <- function(x) {
  if (is.null(x)) return("<sin vars>")
  if (is.character(x)) {
    x <- trimws(x); x <- x[nzchar(x)]
    if (!length(x)) return("<sin vars>")
    return(paste(x, collapse = ", "))
  }
  if (is.list(x)) {
    vals <- unlist(lapply(x, .fmt_vars), use.names = FALSE)
    vals <- vals[!is.na(vals) & nzchar(trimws(vals)) & vals != "<sin vars>"]
    if (!length(vals)) return("<sin vars>")
    return(paste(vals, collapse = ", "))
  }
  "<sin vars>"
}

.fmt_grafico <- function(el_plot) {
  if (!inherits(el_plot, "ppt_element")) return("<sin elemento>")
  et <- el_plot$.element_type %||% "<NA>"
  # vars o var
  vv <- el_plot$var %||% el_plot$vars %||% NULL
  paste0(et, " | vars: ", .fmt_vars(vv))
}

.msg_diapo <- function(i, n, stype, el_plot = NULL, mensajes_progreso = FALSE) {
  # stype: section/slide_1/slide_2...
  tipo   <- stype %||% "<NA>"
  header <- sprintf("Diapositiva %03d/%03d — %s", i, n, tipo)

  if (!isTRUE(mensajes_progreso)) return(invisible(NULL))

  if (is.null(el_plot)) {
    message(header)
  } else {
    message(header, " — gráfico: ", .fmt_grafico(el_plot))
  }
  invisible(NULL)
}

# ---------------------------------------------------------------------------
# Radar helpers internos (FIX)
# - Cruce: usa keys (códigos) para filtrar y labels para mostrar (tabla/leyenda)
# - Colores de líneas: desde paleta_<list_name_del_cruce> (si existe), usando labels
# ---------------------------------------------------------------------------

.as_chr <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  trimws(x)
}

# ---- paleta auto (paleta_<listname>) desde env_diapos -----------------------
# (usa el mismo patrón que tu .paleta_auto() del exportador)
.paleta_auto_local <- function(list_name, env) {
  if (is.null(env) || !is.environment(env)) env <- parent.frame()
  if (is.na(list_name) || !nzchar(list_name)) return(NULL)
  obj_name <- paste0("paleta_", list_name)
  if (!exists(obj_name, envir = env, inherits = TRUE)) return(NULL)
  pal <- get(obj_name, envir = env, inherits = TRUE)
  if (!is.atomic(pal) || is.null(names(pal))) return(NULL)
  pal
}

# ---- list_name de una var (survey) ------------------------------------------
.list_name_of_var_local <- function(v, survey) {
  if (!is.data.frame(survey)) return(NA_character_)
  if ("list_name" %in% names(survey)) {
    idx <- !is.na(survey$name) & survey$name == v
    x <- survey$list_name[idx]
    x <- x[!is.na(x) & nzchar(x)]
    if (length(x)) return(x[1])
  }
  if ("list_norm" %in% names(survey)) {
    idx <- !is.na(survey$name) & survey$name == v
    x <- survey$list_norm[idx]
    x <- x[!is.na(x) & nzchar(x)]
    if (length(x)) return(x[1])
  }
  NA_character_
}

# ---- map del cruce: keys (para filtrar) + labels (para mostrar) -------------
.radar_cruce_map <- function(data, cruce, survey, orders_list,
                             env_paletas = parent.frame()) {

  # categorías del cruce desde instrumento
  cats <- get_categorias(
    var              = cruce,
    data             = data,
    survey           = survey,
    orders_list      = orders_list,
    opciones_excluir = NULL
  )

  estr_codes  <- .as_chr(cats$codes)
  estr_labels <- .as_chr(cats$labels)

  # fallback si no hay nada en instrumento
  if (!length(estr_codes) || !length(estr_labels)) {
    v <- sort(unique(na.omit(.as_chr(data[[cruce]]))))
    pal <- NULL
    return(list(keys = v, labels = v, palette = pal))
  }

  # valores observados en data
  v_estr <- .as_chr(data[[cruce]])
  v_estr <- v_estr[nzchar(v_estr)]

  # ¿data usa códigos o labels?
  usa_codes  <- any(v_estr %in% estr_codes)
  usa_labels <- any(v_estr %in% estr_labels)

  keys_vec <- if (usa_codes || !usa_labels) estr_codes else estr_labels
  labels_vec <- estr_labels

  # paleta por list_name del CRUCE
  ln_cruce <- .list_name_of_var_local(cruce, survey)
  pal <- .paleta_auto_local(ln_cruce, env = env_paletas)

  # si paleta existe:
  # - idealmente nombres de paleta son labels (como sueles hacer en listas)
  # - si nombres son códigos, mapear a labels
  pal_out <- NULL
  if (!is.null(pal)) {
    pal_names <- .as_chr(names(pal))

    # map code -> label (vector nombrado)
    map_code2lab <- stats::setNames(labels_vec, estr_codes)

    if (any(pal_names %in% labels_vec)) {
      # ya viene por labels
      pal_out <- pal
      # reordenar a labels_vec si se puede
      keep <- labels_vec[labels_vec %in% names(pal_out)]
      if (length(keep)) pal_out <- pal_out[keep]
    } else if (any(pal_names %in% estr_codes)) {
      # viene por códigos: renombrar a labels
      new_names <- ifelse(pal_names %in% names(map_code2lab), unname(map_code2lab[pal_names]), pal_names)
      pal_out <- pal
      names(pal_out) <- new_names
      keep <- labels_vec[labels_vec %in% names(pal_out)]
      if (length(keep)) pal_out <- pal_out[keep]
    } else {
      # nombres no calzan: se usa tal cual (fallback)
      pal_out <- pal
    }
  }

  list(keys = keys_vec, labels = labels_vec, palette = pal_out)
}

# ---- cruce labels safe (fallback) -------------------------------------------
.apply_cruce_labels_safe <- function(df, cruce_name) {
  # Reutiliza el helper ya definido en .render_numerico() si existe en scope:
  if (exists(".apply_cruce_labels", mode = "function", inherits = TRUE)) {
    inst <- NULL
    if (exists(".get_inst", mode = "function", inherits = TRUE)) inst <- .get_inst()
    out <- .apply_cruce_labels(df[[cruce_name]], inst, cruce_name)
    return(list(x = out$x, lvls = out$lvls))
  }
  x <- .as_chr(df[[cruce_name]])
  list(x = x, lvls = unique(x[nzchar(x)]))
}

# Radar SM: devuelve long (eje, grupo, valor) + attr(palette)
.radar_build_sm <- function(var, cruce = NULL, top_n = NULL,
                            sm_omit_codes  = NULL,
                            sm_omit_labels = NULL,
                            sm_omit_na     = TRUE,
                            data, survey, orders_list,
                            env_paletas = parent.frame()) {

  cats <- get_categorias(
    var              = var,
    data             = data,
    survey           = survey,
    orders_list      = orders_list,
    opciones_excluir = NULL
  )

  codes_row  <- .as_chr(cats$codes)
  labels_row <- .as_chr(cats$labels)

  # 1) Omit manual (por codes / labels)
  if (!is.null(sm_omit_codes) || !is.null(sm_omit_labels)) {

    keep <- rep(TRUE, length(codes_row))

    if (!is.null(sm_omit_codes) && length(sm_omit_codes)) {
      oc <- .as_chr(sm_omit_codes)
      keep <- keep & !(codes_row %in% oc)
    }

    if (!is.null(sm_omit_labels) && length(sm_omit_labels)) {
      ol <- .as_chr(sm_omit_labels)
      keep <- keep & !(labels_row %in% ol)
    }

    codes_row  <- codes_row[keep]
    labels_row <- labels_row[keep]
  }

  if (!length(codes_row)) return(NULL)

  # 2) Drop "Total" (RECALCULAR *después* del omit)
  op_chr <- trimws(tolower(as.character(labels_row)))
  cd_chr <- trimws(tolower(as.character(codes_row)))

  drop_total <- (op_chr == "total") | (cd_chr == "total") | is.na(op_chr) | (op_chr == "")

  if (any(drop_total)) {
    codes_row  <- codes_row[!drop_total]
    labels_row <- labels_row[!drop_total]
  }

  if (!length(codes_row)) return(NULL)

  # 3) Tipo de pregunta
  tp <- tipo_pregunta(var, survey = survey, sm_vars_force = NULL, data = data)

  # ---- series (cruce) FIX ---------------------------------------------------
  pal_series <- NULL

  if (is.null(cruce)) {
    lvls_keys   <- "Total"
    lvls_labels <- "Total"
    grupo_x     <- rep("Total", nrow(data))
  } else {

    # map cruce usando instrumento (keys vs labels) + paleta
    cm <- .radar_cruce_map(
      data        = data,
      cruce       = cruce,
      survey      = survey,
      orders_list = orders_list,
      env_paletas = env_paletas
    )

    # keys para filtrar
    lvls_keys <- cm$keys
    # labels para mostrar (leyenda/tabla)
    lvls_labels <- cm$labels
    pal_series <- cm$palette

    # si por alguna razón no hay levels válidos, fallback a Total
    lvls_keys   <- lvls_keys[nzchar(trimws(lvls_keys))]
    lvls_labels <- lvls_labels[nzchar(trimws(lvls_labels))]

    if (!length(lvls_keys) || !length(lvls_labels)) {
      lvls_keys   <- "Total"
      lvls_labels <- "Total"
      grupo_x     <- rep("Total", nrow(data))
      pal_series  <- NULL
    } else {
      # cruce en data (raw) se compara contra keys (códigos si aplica)
      grupo_x <- .as_chr(data[[cruce]])
    }
  }

  # ---- top_n ---------------------------------------------------------------
  if (!is.null(top_n) && length(codes_row) > top_n) {
    n_all <- contar_por_opcion(
      data       = data,
      var        = var,
      codes      = codes_row,
      tp         = tp,
      mask       = rep(TRUE, nrow(data)),
      weight_col = "peso"
    )
    ord <- order(n_all, decreasing = TRUE)
    keep <- head(ord, top_n)
    codes_row  <- codes_row[keep]
    labels_row <- labels_row[keep]
  }

  # ---- construir long -------------------------------------------------------
  out_rows <- list()

  if (identical(lvls_keys, "Total")) {

    mask_g <- rep(TRUE, nrow(data))

    n_vec <- contar_por_opcion(
      data       = data,
      var        = var,
      codes      = codes_row,
      tp         = tp,
      mask       = mask_g,
      weight_col = "peso"
    )

    N_g <- denominador_validos(
      data       = data,
      var        = var,
      codes      = codes_row,
      tp         = tp,
      mask       = mask_g,
      weight_col = "peso"
    )

    pct <- if (is.finite(N_g) && N_g > 0) as.numeric(n_vec) / N_g else rep(NA_real_, length(n_vec))

    out_rows[[1]] <- tibble::tibble(
      eje   = as.character(labels_row),
      grupo = "Total",
      valor = as.numeric(pct)
    )

    d <- dplyr::bind_rows(out_rows)
    if (!nrow(d)) return(NULL)
    d$grupo <- factor(d$grupo, levels = "Total")
    attr(d, "palette") <- pal_series
    return(d)
  }

  # loop por niveles: filtrar con KEY, mostrar LABEL
  for (j in seq_along(lvls_keys)) {

    key_j <- lvls_keys[j]
    lab_j <- lvls_labels[j]

    mask_g <- (!is.na(grupo_x) & .as_chr(grupo_x) == .as_chr(key_j))

    n_vec <- contar_por_opcion(
      data       = data,
      var        = var,
      codes      = codes_row,
      tp         = tp,
      mask       = mask_g,
      weight_col = "peso"
    )

    N_g <- denominador_validos(
      data       = data,
      var        = var,
      codes      = codes_row,
      tp         = tp,
      mask       = mask_g,
      weight_col = "peso"
    )

    pct <- if (is.finite(N_g) && N_g > 0) as.numeric(n_vec) / N_g else rep(NA_real_, length(n_vec))

    out_rows[[length(out_rows) + 1]] <- tibble::tibble(
      eje   = as.character(labels_row),
      grupo = as.character(lab_j),   # <- LABEL visible
      valor = as.numeric(pct)
    )
  }

  d <- dplyr::bind_rows(out_rows)
  if (!nrow(d)) return(NULL)

  d$grupo <- factor(d$grupo, levels = lvls_labels)

  # adjuntar paleta para series (por labels)
  if (!is.null(pal_series) && !is.null(names(pal_series))) {
    keep <- lvls_labels[lvls_labels %in% names(pal_series)]
    if (length(keep)) pal_series <- pal_series[keep]
  } else {
    pal_series <- NULL
  }
  attr(d, "palette") <- pal_series
  d
}

# ---------------------------------------------------------------------------
# Radar BOX (Top/Bottom box): devuelve long + attr(palette)
# ---------------------------------------------------------------------------
.radar_build_box <- function(vars, cruce = NULL, box_labels,
                             data, survey, orders_list,
                             titulo_tabla = "Top 2 Box",
                             env_paletas = parent.frame()) {

  vars <- trimws(as.character(vars)); vars <- vars[nzchar(vars)]
  if (!length(vars)) return(NULL)

  # asegurar 1 list_name para el set de respuestas (vars)
  lns <- vapply(vars, .list_name_of_var_local, character(1), survey = survey)
  lns <- unique(lns[!is.na(lns) & nzchar(lns)])
  if (length(lns) != 1L) {
    stop("radar(box): `vars` no comparten un único list_name. Encontrados: ",
         paste(lns, collapse = " | "), call. = FALSE)
  }

  cats0 <- get_categorias(
    var              = vars[1],
    data             = data,
    survey           = survey,
    orders_list      = orders_list,
    opciones_excluir = NULL
  )
  codes_all  <- .as_chr(cats0$codes)
  labels_all <- .as_chr(cats0$labels)
  if (!length(codes_all)) return(NULL)

  # map labels -> codes para box
  codes_box <- codes_all[labels_all %in% box_labels]
  if (length(codes_box) != length(box_labels)) {
    stop(
      "radar(box): no se mapearon correctamente los códigos desde `box_labels`.\n",
      "Labels pedidos: ", paste(box_labels, collapse = " | "),
      "\nLabels disponibles: ", paste(unique(labels_all), collapse = " | "),
      call. = FALSE
    )
  }

  # ---- cruce FIX + paleta por CRUCE ----------------------------------------
  pal_series <- NULL

  if (is.null(cruce)) {
    lvls_keys   <- "Total"
    lvls_labels <- "Total"
    grupo_x     <- rep("Total", nrow(data))
  } else {

    cm <- .radar_cruce_map(
      data        = data,
      cruce       = cruce,
      survey      = survey,
      orders_list = orders_list,
      env_paletas = env_paletas
    )

    lvls_keys   <- cm$keys
    lvls_labels <- cm$labels
    pal_series  <- cm$palette

    lvls_keys   <- lvls_keys[nzchar(trimws(lvls_keys))]
    lvls_labels <- lvls_labels[nzchar(trimws(lvls_labels))]

    if (!length(lvls_keys) || !length(lvls_labels)) {
      lvls_keys   <- "Total"
      lvls_labels <- "Total"
      grupo_x     <- rep("Total", nrow(data))
      pal_series  <- NULL
    } else {
      grupo_x <- .as_chr(data[[cruce]])
    }
  }

  .count_in_codes <- function(v, mask, codes_keep, weight_col = "peso") {
    w <- get_pesos(data, weight_col)
    v_codes <- .as_chr(data[[v]])
    ok <- mask & nzchar(v_codes) & (v_codes %in% codes_keep)
    sum(w[ok], na.rm = TRUE)
  }

  out_rows <- list()

  for (v in vars) {

    tpv <- tipo_pregunta(v, survey = survey, sm_vars_force = NULL, data = data)
    if (!identical(tpv, "so")) tpv <- "so"

    eje_lbl <- label_variable(
      v,
      dic_vars = dplyr::select(survey, name, label),
      labels_override = NULL,
      data = data
    )

    cats_v <- get_categorias(
      var              = v,
      data             = data,
      survey           = survey,
      orders_list      = orders_list,
      opciones_excluir = NULL
    )
    codes_v  <- .as_chr(cats_v$codes)
    labels_v <- .as_chr(cats_v$labels)

    # map box por variable (por si el set cambia)
    codes_box_v <- codes_v[labels_v %in% box_labels]
    if (length(codes_box_v) != length(box_labels)) codes_box_v <- codes_box

    if (identical(lvls_keys, "Total")) {

      mask_g <- rep(TRUE, nrow(data))

      N_g <- denominador_validos(
        data       = data,
        var        = v,
        codes      = codes_v,
        tp         = tpv,
        mask       = mask_g,
        weight_col = "peso"
      )

      n_box <- .count_in_codes(v, mask_g, codes_box_v, weight_col = "peso")
      pct <- if (is.finite(N_g) && N_g > 0) as.numeric(n_box) / N_g else NA_real_

      out_rows[[length(out_rows) + 1]] <- tibble::tibble(
        eje   = as.character(eje_lbl),
        grupo = "Total",
        valor = as.numeric(pct)
      )

      next
    }

    # loop por niveles: filtrar con KEY, mostrar LABEL
    for (j in seq_along(lvls_keys)) {

      key_j <- lvls_keys[j]
      lab_j <- lvls_labels[j]

      mask_g <- (!is.na(grupo_x) & .as_chr(grupo_x) == .as_chr(key_j))

      N_g <- denominador_validos(
        data       = data,
        var        = v,
        codes      = codes_v,
        tp         = tpv,
        mask       = mask_g,
        weight_col = "peso"
      )

      n_box <- .count_in_codes(v, mask_g, codes_box_v, weight_col = "peso")
      pct <- if (is.finite(N_g) && N_g > 0) as.numeric(n_box) / N_g else NA_real_

      out_rows[[length(out_rows) + 1]] <- tibble::tibble(
        eje   = as.character(eje_lbl),
        grupo = as.character(lab_j),   # <- LABEL visible
        valor = as.numeric(pct)
      )
    }
  }

  d <- dplyr::bind_rows(out_rows)
  if (!nrow(d)) return(NULL)

  d$grupo <- factor(d$grupo, levels = lvls_labels)

  # adjuntar paleta para series (por labels)
  if (!is.null(pal_series) && !is.null(names(pal_series))) {
    keep <- lvls_labels[lvls_labels %in% names(pal_series)]
    if (length(keep)) pal_series <- pal_series[keep]
  } else {
    pal_series <- NULL
  }
  attr(d, "palette") <- pal_series
  d
}

# =============================================================================
# PLAN acumulativo por chunks: diapo() / .ppt_plan_env
# =============================================================================

.ppt_plan_name <- ".ppt_plan_accum"

#' @keywords internal
.ppt_plan_env <- function(env = parent.frame()) {
  if (!is.environment(env)) stop("`.ppt_plan_env()`: `env` debe ser environment.", call. = FALSE)

  if (!exists(.ppt_plan_name, envir = env, inherits = FALSE)) {
    init <- structure(list(), class = c("ppt_plan", "list"))
    assign(.ppt_plan_name, init, envir = env)
  }

  get(.ppt_plan_name, envir = env, inherits = FALSE)
}

#' @keywords internal
.ppt_plan_set <- function(plan, env = parent.frame()) {
  if (!is.environment(env)) stop("`.ppt_plan_set()`: `env` debe ser environment.", call. = FALSE)
  if (!is.list(plan)) stop("`.ppt_plan_set()`: `plan` debe ser lista.", call. = FALSE)
  class(plan) <- unique(c("ppt_plan","list", class(plan)))
  assign(.ppt_plan_name, plan, envir = env)
  invisible(plan)
}

#' @keywords internal
.ppt_plan_push <- function(slide, env = parent.frame()) {
  if (is.null(slide) || !inherits(slide, "ppt_slide")) {
    stop("`.ppt_plan_push()`: `slide` debe ser `ppt_slide`.", call. = FALSE)
  }
  plan <- .ppt_plan_env(env)
  plan[[length(plan) + 1L]] <- slide
  .ppt_plan_set(plan, env)
}

#' @keywords internal
.ppt_plan_clear <- function(env = parent.frame()) {
  .ppt_plan_set(structure(list(), class = c("ppt_plan","list")), env)
}

#' @export
diapo <- function(slide, env = parent.frame()) {
  .ppt_plan_push(slide, env = env)
  slide
}

#' Obtener el plan acumulado sin limpiarlo
#'
#' Devuelve el plan de diapositivas acumulado por [diapo()] en el entorno
#' indicado. A diferencia de [reporte_ppt_plan()], no borra el acumulador.
#' Útil para capturar el plan y pasarlo luego a [plan_word()]
#' sin necesidad de guardar el objeto pesado devuelto por [reporte_ppt_plan()].
#'
#' @param env Entorno donde se buscará el acumulador (por defecto el entorno
#'   del llamador, igual que [diapo()]).
#' @return Objeto `ppt_plan` (lista de `ppt_slide`).
#' @export
p_get_plan <- function(env = parent.frame()) {
  .ppt_plan_env(env)
}


# =============================================================================
# RESET PPT — limpiar acumulados de diapo() + objetos diapo_###
# - Se una al INICIO del script / qmd antes de definir diapo_### otra vez.
# =============================================================================
#' @export
p_reset <- function(
    env = parent.frame(),
    drop_diapos = TRUE,     # borra diapo_### del env
    drop_plan   = TRUE,     # borra plan acumulado si existe
    drop_misc   = TRUE,     # borra caches comunes (rendered/logs) si existen
    verbose     = TRUE
) {

  # helper
  .rm_if_exists <- function(nm, envir) {
    if (exists(nm, envir = envir, inherits = FALSE)) {
      rm(list = nm, envir = envir)
      TRUE
    } else FALSE
  }

  removed <- character(0)

  # 1) Borrar diapo_###
  if (isTRUE(drop_diapos)) {
    nms <- ls(envir = env, all.names = TRUE)
    di <- nms[grepl("^diapo_\\d{3}$", nms)]
    if (length(di)) {
      rm(list = di, envir = env)
      removed <- c(removed, di)
    }
  }

  # 2) Borrar plan acumulado (múltiples estrategias)
  if (isTRUE(drop_plan)) {

    # a) si existe una función oficial para limpiar, úsala
    if (exists(".ppt_plan_clear", mode = "function", inherits = TRUE)) {
      try(.ppt_plan_clear(env), silent = TRUE)
      removed <- c(removed, "<.ppt_plan_clear()>")
    }

    # b) posibles nombres típicos de plan acumulado en el env
    candidates <- c(
      ".ppt_plan", "ppt_plan", "plan_ppt",
      ".plan", "plan", ".ppt_plan_accum",
      ".ppt_plan_obj", ".ppt_plan_cache"
    )
    for (nm in candidates) {
      if (.rm_if_exists(nm, env)) removed <- c(removed, nm)
    }

    # c) si tienes un “nombre del plan” guardado en .ppt_plan_name
    if (exists(".ppt_plan_name", envir = env, inherits = TRUE)) {
      nm_plan <- try(get(".ppt_plan_name", envir = env, inherits = TRUE), silent = TRUE)
      if (is.character(nm_plan) && length(nm_plan) == 1L && nzchar(nm_plan)) {
        if (.rm_if_exists(nm_plan, env)) removed <- c(removed, nm_plan)
      }
    }
  }

  # 3) Borrar caches comunes (por si guardaste cosas auxiliares)
  if (isTRUE(drop_misc)) {
    misc <- c("rendered", "ppt_rendered", ".ppt_rendered", "ppt_log", ".ppt_log")
    for (nm in misc) {
      if (.rm_if_exists(nm, env)) removed <- c(removed, nm)
    }
  }

  if (isTRUE(verbose)) {
    if (!length(removed)) message("✅ ppt_reset(): nada que limpiar (todo ya estaba limpio).")
    else message("🧹 ppt_reset(): limpiado -> ", paste(unique(removed), collapse = ", "))
  }

  invisible(unique(removed))
}



`%||%` <- function(x, y) if (!is.null(x)) x else y
