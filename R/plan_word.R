# =============================================================================
# CONTRATO DE LAYOUT Y PLACEHOLDER — WORD (SIMPLE)
# - Word NO replica layouts complejos de PPT.
# - Estructura: (Portada opcional) + (Secciones opcionales) + Bloques 1 a 1
#   donde cada bloque es: TÍTULO (párrafo) + IMAGEN (PNG) + BASE (párrafo) + FOOTER (párrafo)
# - Los gráficos se renderizan como PNG (NO editables).
# =============================================================================

#' @keywords internal
.WORD_CONTRACT <- list(

  # ------------------------------------------------------------
  # TITLE (portada / cabecera del documento) — opcional
  # ------------------------------------------------------------
  title = list(
    style = "Title",
    slots = list(
      title      = list(kind = "paragraph", style = "Title"),
      subtitle   = list(kind = "paragraph", style = "Subtitle"),
      date       = list(kind = "paragraph", style = "Date"),
      meta_left  = list(kind = "paragraph", style = "MetaLeft"),
      meta_right = list(kind = "paragraph", style = "MetaRight"),
      meta_line  = list(kind = "paragraph", style = "MetaLine")
    )
  ),

  # ------------------------------------------------------------
  # SECTION (encabezado de sección) — opcional
  # ------------------------------------------------------------
  section = list(
    style = "Heading 1",
    slots = list(
      title    = list(kind = "paragraph", style = "Heading 1"),
      subtitle = list(kind = "paragraph", style = "Heading 2")
    )
  ),

  # ------------------------------------------------------------
  # BLOCK_1 (bloque estándar) — SIEMPRE 1 gráfico/tabla/kpi por bloque
  # ------------------------------------------------------------
  block_1 = list(
    layout = "block_1",
    slots  = list(
      title  = list(kind = "paragraph", style = "Heading 2"),
      main   = list(kind = "figure"),
      base   = list(kind = "paragraph", style = "BodyText"),
      footer = list(kind = "paragraph", style = "FooterText")
    )
  )
)

# =============================================================================
# UTILIDADES
# =============================================================================

`%||%` <- function(x, y) if (!is.null(x)) x else y

.w_norm_text1 <- function(x, blank = NULL) {
  if (is.null(x)) return(NULL)
  x <- as.character(x)[1]
  if (!nzchar(trimws(x))) return(blank)
  x
}

# =============================================================================
# CLASES: word_block / word_plan
# - Para Word, el "elemento" que se imprime como imagen es un `ppt_element`
#   (tus p_* existentes), porque se reutiliza el dispatcher .render_element().
# =============================================================================

#' @keywords internal
.w_as_block <- function(x) { class(x) <- c("word_block", "list"); x }

#' @keywords internal
.w_as_plan <- function(x)  { class(x) <- c("word_plan",  "list"); x }

# =============================================================================
# BLOQUES (WORD) — API MINIMAL
# =============================================================================

#' @title Bloque de portada / cabecera (opcional)
#' @export
w_title <- function(
    title,
    subtitle   = NULL,
    date       = NULL,
    meta_left  = NULL,
    meta_right = NULL,
    meta_line  = NULL,
    meta       = list()
) {
  title <- .w_norm_text1(title)
  if (is.null(title)) stop("`w_title()`: `title` debe ser texto no vacío.", call. = FALSE)

  if (!is.list(meta)) stop("`w_title()`: `meta` debe ser lista.", call. = FALSE)

  .w_as_block(list(
    .block_type = "title",
    slots = list(
      title      = title,
      subtitle   = .w_norm_text1(subtitle,   blank = NULL),
      date       = .w_norm_text1(date,       blank = NULL),
      meta_left  = .w_norm_text1(meta_left,  blank = NULL),
      meta_right = .w_norm_text1(meta_right, blank = NULL),
      meta_line  = .w_norm_text1(meta_line,  blank = NULL)
    ),
    meta = meta
  ))
}

#' @title Bloque de sección (opcional)
#' @export
w_section <- function(title, subtitle = NULL, meta = list()) {
  title <- .w_norm_text1(title)
  if (is.null(title)) stop("`w_section()`: `title` debe ser texto no vacío.", call. = FALSE)

  if (!is.list(meta)) stop("`w_section()`: `meta` debe ser lista.", call. = FALSE)

  .w_as_block(list(
    .block_type = "section",
    slots = list(
      title    = title,
      subtitle = .w_norm_text1(subtitle, blank = NULL)
    ),
    meta = meta
  ))
}

#' @title Bloque estándar (título + PNG + base + footer)
#' @param plot `ppt_element` (tus p_* existentes).
#' @export
w_block_1 <- function(plot, title = NULL, base = NULL, footer = NULL, overrides = list(), meta = list()) {

  if (is.null(plot) || !inherits(plot, "ppt_element")) {
    stop("`w_block_1()`: `plot` debe ser `ppt_element` (de tus p_*).", call. = FALSE)
  }

  title  <- .w_norm_text1(title,  blank = NULL)
  base   <- .w_norm_text1(base,   blank = NULL)
  footer <- .w_norm_text1(footer, blank = NULL)

  if (!is.list(overrides)) stop("`w_block_1()`: `overrides` debe ser lista.", call. = FALSE)
  if (!is.list(meta))      stop("`w_block_1()`: `meta` debe ser lista.", call. = FALSE)

  # aplicar overrides sin mutar el elemento original
  el <- plot
  el$overrides <- modifyList(el$overrides %||% list(), overrides)

  .w_as_block(list(
    .block_type = "block_1",
    plot  = el,
    slots = list(
      title  = title,
      main   = el,
      base   = base,
      footer = footer
    ),
    meta = meta
  ))
}

# =============================================================================
# PLAN: declarativo y/o acumulativo por chunks
# =============================================================================

#' @export
w_plan <- function(..., blocks = NULL) {
  out <- if (!is.null(blocks)) {
    if (!is.list(blocks)) stop("`w_plan()`: `blocks` debe ser lista.", call. = FALSE)
    blocks
  } else list(...)

  if (!length(out)) return(.w_as_plan(list()))

  bad <- vapply(out, function(x) !inherits(x, "word_block"), logical(1))
  if (any(bad)) stop("`w_plan()`: todos deben ser `word_block`. Malos: ", paste(which(bad), collapse = ", "), call. = FALSE)

  .w_as_plan(out)
}

.word_plan_name <- ".word_plan_accum"

#' @keywords internal
.word_plan_env <- function(env = parent.frame()) {
  if (!is.environment(env)) stop("`.word_plan_env()`: `env` debe ser environment.", call. = FALSE)
  if (!exists(.word_plan_name, envir = env, inherits = FALSE)) {
    assign(.word_plan_name, .w_as_plan(list()), envir = env)
  }
  get(.word_plan_name, envir = env, inherits = FALSE)
}

#' @keywords internal
.word_plan_set <- function(plan, env = parent.frame()) {
  if (!is.environment(env)) stop("`.word_plan_set()`: `env` debe ser environment.", call. = FALSE)
  if (!is.list(plan)) stop("`.word_plan_set()`: `plan` debe ser lista.", call. = FALSE)
  plan <- .w_as_plan(plan)
  assign(.word_plan_name, plan, envir = env)
  invisible(plan)
}

#' @keywords internal
.word_plan_push <- function(block, env = parent.frame()) {
  if (is.null(block) || !inherits(block, "word_block")) {
    stop("`.word_plan_push()`: `block` debe ser `word_block`.", call. = FALSE)
  }
  plan <- .word_plan_env(env)
  plan[[length(plan) + 1L]] <- block
  .word_plan_set(plan, env)
}

#' @keywords internal
.word_plan_clear <- function(env = parent.frame()) .word_plan_set(.w_as_plan(list()), env)

#' @export
w_add <- function(block, env = parent.frame()) {
  .word_plan_push(block, env = env)
  block
}

# -----------------------------------------------------------------------------
# Recolección alternativa: convención bloque_###
# -----------------------------------------------------------------------------
#' @keywords internal
.collect_word_objects <- function(env = parent.frame(), strict = FALSE) {

  if (!is.environment(env)) {
    stop("`.collect_word_objects()`: `env` debe ser un environment.", call. = FALSE)
  }

  nms <- ls(envir = env, all.names = TRUE)
  nms <- nms[grepl("^bloque_\\d{3}$", nms)]

  if (!length(nms)) return(list())

  ids <- as.integer(sub("^bloque_(\\d{3})$", "\\1", nms))
  ord <- order(ids)
  nms <- nms[ord]
  ids <- ids[ord]

  objs <- mget(nms, envir = env, inherits = FALSE)

  bad <- vapply(objs, function(x) !inherits(x, "word_block"), logical(1))
  if (any(bad)) {
    msg <- paste0(
      "`.collect_word_objects()`: estos objetos `bloque_###` no son `word_block`: ",
      paste(names(objs)[bad], collapse = ", ")
    )
    if (isTRUE(strict)) stop(msg, call. = FALSE) else warning(msg, call. = FALSE)
  }

  if (isTRUE(strict) && length(ids) > 1) {
    dif <- diff(ids)
    if (any(dif != 1L)) {
      stop("strict=TRUE: los `bloque_###` no son consecutivos (hay saltos).", call. = FALSE)
    }
  }

  objs
}

# =============================================================================
# PRESETS WORD
# =============================================================================

#' @title Definir presets para Word (PNG + estilos de párrafo)
#' @export
w_presets <- function(
    image = list(width_in = 6.6, height_in = 3.9, dpi = 300, bg = "white"),
    title_style = list(style_name = NULL, font = "Calibri", size = 12, bold = TRUE,  italic = FALSE, color = "000000"),
    base_style  = list(style_name = NULL, font = "Calibri", size = 9,  bold = FALSE, italic = FALSE, color = "404040",
                       formato = "Base: %s", sufijo_auto = NULL),
    footer_style = list(style_name = NULL, font = "Calibri", size = 9, bold = FALSE, italic = FALSE, color = "404040"),
    figure_numbering = list(enabled = TRUE, prefix = "Gráfico", sep = ". ", prepend_on_manual = FALSE),
    pagebreak_between = FALSE,
    pagebreak_after_title = TRUE
) {

  image$width_in  <- as.numeric(image$width_in  %||% 6.6)
  image$height_in <- as.numeric(image$height_in %||% 3.9)
  image$dpi       <- as.integer(image$dpi %||% 300L)
  image$bg        <- as.character(image$bg %||% "white")[1]

  base_style$formato <- as.character(base_style$formato %||% "Base: %s")[1]

  out <- list(
    image = image,
    title_style = title_style,
    base_style = base_style,
    footer_style = footer_style,
    figure_numbering = figure_numbering,
    pagebreak_between = isTRUE(pagebreak_between),
    pagebreak_after_title = isTRUE(pagebreak_after_title)
  )
  class(out) <- c("word_presets", "list")
  out
}

# =============================================================================
# RESET WORD — limpiar acumulados (w_add) + objetos bloque_###
# =============================================================================
#' @export
w_reset <- function(
    env = parent.frame(),
    drop_blocks = TRUE,
    drop_plan   = TRUE,
    drop_misc   = TRUE,
    verbose     = TRUE
) {
  .rm_if_exists <- function(nm, envir) {
    if (exists(nm, envir = envir, inherits = FALSE)) { rm(list = nm, envir = envir); TRUE } else FALSE
  }

  removed <- character(0)

  if (isTRUE(drop_blocks)) {
    nms <- ls(envir = env, all.names = TRUE)
    bl  <- nms[grepl("^bloque_\\d{3}$", nms)]
    if (length(bl)) { rm(list = bl, envir = env); removed <- c(removed, bl) }
  }

  if (isTRUE(drop_plan)) {
    if (exists(".word_plan_clear", mode = "function", inherits = TRUE)) {
      try(.word_plan_clear(env), silent = TRUE)
      removed <- c(removed, "<.word_plan_clear()>")
    }
    candidates <- c(".word_plan_accum", ".word_plan", "word_plan", "plan_word", ".plan", "plan")
    for (nm in candidates) if (.rm_if_exists(nm, env)) removed <- c(removed, nm)
    if (.rm_if_exists(.word_plan_name, env)) removed <- c(removed, .word_plan_name)
  }

  if (isTRUE(drop_misc)) {
    misc <- c("word_rendered", ".word_rendered", "word_log", ".word_log")
    for (nm in misc) if (.rm_if_exists(nm, env)) removed <- c(removed, nm)
  }

  if (isTRUE(verbose)) {
    if (!length(removed)) message("✅ w_reset(): nada que limpiar.")
    else message("🧹 w_reset(): limpiado -> ", paste(unique(removed), collapse = ", "))
  }

  invisible(unique(removed))
}

# =============================================================================
# REPORTE WORD (PLAN) — SIMPLE
# - Reutiliza tu dispatcher del PPT: `.render_element(ppt_element)` -> ggplot
# - Inserta el ggplot como PNG en Word
# =============================================================================

#' @export
reporte_word_plan <- function(
    data,
    instrumento        = NULL,
    path_docx          = "reporte_word_plan.docx",
    presets            = NULL,
    plan               = NULL,
    env_bloques        = parent.frame(),
    strict_bloques     = FALSE,
    template_docx      = getOption("prosecnur.template_docx", NA_character_),
    mensajes_progreso  = TRUE,
    solo_lista         = FALSE
) {

  # -----------------------
  # 0) Validaciones
  # -----------------------
  if (!is.data.frame(data)) stop("`data` debe ser data.frame/tibble.", call. = FALSE)

  if (!requireNamespace("officer", quietly = TRUE)) stop("Se requiere 'officer'.", call. = FALSE)
  if (!requireNamespace("ggplot2", quietly = TRUE)) stop("Se requiere 'ggplot2'.", call. = FALSE)

  # dplyr/tibble solo si se arma log bonito (si no están, se usa base)
  has_dplyr  <- requireNamespace("dplyr",  quietly = TRUE)
  has_tibble <- requireNamespace("tibble", quietly = TRUE)

  if (is.null(instrumento)) {
    instrumento <- attr(data, "instrumento_reporte", exact = TRUE)
    if (is.null(instrumento)) stop("No se proporcionó `instrumento` y falta atributo `instrumento_reporte`.", call. = FALSE)
  }

  presets <- presets %||% w_presets()
  if (!inherits(presets, "word_presets")) stop("`presets` debe venir de `w_presets()`.", call. = FALSE)

  # -----------------------
  # Helpers: render ppt_element -> ggplot
  # -----------------------
  .render_element_plot <- function(el) {
    if (is.null(el) || !inherits(el, "ppt_element")) {
      stop(".render_element_plot(): `el` debe ser `ppt_element`.", call. = FALSE)
    }
    if (!exists(".render_element", mode = "function", inherits = TRUE)) {
      stop(
        "No existe `.render_element()` en el entorno. ",
        "Word reutiliza el dispatcher del PPT: `.render_element(ppt_element)` -> ggplot.",
        call. = FALSE
      )
    }
    .render_element(el)
  }

  .plot_to_png <- function(p, file, width_in, height_in, dpi, bg) {
    tryCatch({
      ggplot2::ggsave(
        filename = file, plot = p,
        width = width_in, height = height_in, units = "in",
        dpi = dpi, bg = bg
      )
      TRUE
    }, error = function(e) FALSE)
  }

  # -----------------------
  # Helpers: estilos
  # -----------------------
  .fp_text_from <- function(st) {
    officer::fp_text(
      font.size   = st$size %||% 11,
      font.family = st$font %||% "Calibri",
      bold        = isTRUE(st$bold %||% FALSE),
      italic      = isTRUE(st$italic %||% FALSE),
      color       = st$color %||% "000000"
    )
  }

  .add_par <- function(doc, text, st) {
    text <- .w_norm_text1(text, blank = NULL)
    if (is.null(text)) return(doc)

    if (!is.null(st$style_name) && nzchar(st$style_name)) {
      return(officer::body_add_par(doc, value = text, style = st$style_name))
    }

    fp   <- .fp_text_from(st)
    fpar <- officer::fpar(officer::ftext(text, prop = fp))
    officer::body_add_fpar(doc, value = fpar)
  }

  # -----------------------
  # Helpers: título auto (si no viene manual)
  # -----------------------
  .title_auto <- function(el, i) {
    t0 <- el$title_slide %||% el$title_block %||% NULL
    if (is.null(t0) && !is.null(el$var))  t0 <- el$var
    if (is.null(t0) && !is.null(el$vars) && length(el$vars)) t0 <- el$vars[1]

    fn <- presets$figure_numbering %||% list()
    if (isTRUE(fn$enabled)) {
      pref <- fn$prefix %||% "Gráfico"
      sep  <- fn$sep %||% ". "
      head <- paste0(pref, " ", i, sep)
      if (!is.null(t0) && nzchar(trimws(as.character(t0)[1]))) return(paste0(head, trimws(as.character(t0)[1])))
      return(head)
    }

    if (!is.null(t0) && nzchar(trimws(as.character(t0)[1]))) return(trimws(as.character(t0)[1]))
    " "
  }

  # -----------------------
  # 1) Normalizar plan
  # -----------------------
  if (is.null(plan)) {

    # prioridad: plan acumulado por w_add()
    plan_accum <- NULL
    if (exists(.word_plan_name, envir = env_bloques, inherits = TRUE)) {
      cand <- get(.word_plan_name, envir = env_bloques, inherits = TRUE)
      if (is.list(cand) && length(cand)) plan_accum <- cand
    }

    if (!is.null(plan_accum) && length(plan_accum)) {
      plan <- plan_accum
      class(plan) <- c("word_plan", "list")
    } else {
      # fallback: objetos bloque_###
      bloques <- .collect_word_objects(env = env_bloques, strict = strict_bloques)
      if (!length(bloques)) stop("No hay bloques para Word (plan vacío).", call. = FALSE)
      plan <- unname(bloques)
      class(plan) <- c("word_plan", "list")
      attr(plan, "bloque_names") <- names(bloques)
    }

  } else {
    if (!is.list(plan)) stop("`plan` debe ser lista de `word_block`.", call. = FALSE)
    class(plan) <- c("word_plan", "list")
  }

  bad <- vapply(plan, function(x) !inherits(x, "word_block"), logical(1))
  if (any(bad)) stop("Plan: hay elementos que no son `word_block` en posiciones: ", paste(which(bad), collapse = ", "), call. = FALSE)

  # -----------------------
  # 2) Abrir docx
  # -----------------------
  if (isTRUE(solo_lista)) {
    doc <- NULL
  } else {
    if (is.null(template_docx) || is.na(template_docx) || !nzchar(template_docx)) {
      doc <- officer::read_docx()
    } else {
      if (!file.exists(template_docx)) stop("No existe `template_docx`: ", template_docx, call. = FALSE)
      doc <- officer::read_docx(path = template_docx)
    }
  }

  # -----------------------
  # 3) Render + escribir bloques
  # -----------------------
  tmp_dir <- file.path(tempdir(), paste0("word_plan_", format(Sys.time(), "%Y%m%d_%H%M%S")))
  dir.create(tmp_dir, recursive = TRUE, showWarnings = FALSE)

  img_w   <- presets$image$width_in
  img_h   <- presets$image$height_in
  img_dpi <- presets$image$dpi
  img_bg  <- presets$image$bg

  rendered_files <- character(0)
  log_rows <- vector("list", length(plan))

  # contador de gráficos (para numeración) SOLO para block_1
  g_i <- 0L

  for (i in seq_along(plan)) {

    btype <- plan[[i]]$.block_type %||% NA_character_

    if (isTRUE(mensajes_progreso)) {
      message(sprintf("Bloque %03d/%03d — %s", i, length(plan), btype %||% "<NA>"))
    }

    # ---- TITLE --------------------------------------------------------------
    if (identical(btype, "title")) {

      if (!isTRUE(solo_lista)) {
        doc <- .add_par(doc, plan[[i]]$slots$title,      presets$title_style)
        doc <- .add_par(doc, plan[[i]]$slots$subtitle,   presets$title_style)
        doc <- .add_par(doc, plan[[i]]$slots$date,       presets$title_style)
        doc <- .add_par(doc, plan[[i]]$slots$meta_left,  presets$title_style)
        doc <- .add_par(doc, plan[[i]]$slots$meta_right, presets$title_style)
        doc <- .add_par(doc, plan[[i]]$slots$meta_line,  presets$title_style)

        if (isTRUE(presets$pagebreak_after_title) && i < length(plan)) {
          doc <- officer::body_add_break(doc)
        }
      }

      log_rows[[i]] <- if (has_tibble) tibble::tibble(
        block_i = i, block_type = "title", element = NA_character_, var = NA_character_, png = NA_character_
      ) else list(block_i = i, block_type = "title", element = NA, var = NA, png = NA)

      next
    }

    # ---- SECTION ------------------------------------------------------------
    if (identical(btype, "section")) {

      if (!isTRUE(solo_lista)) {
        # Heading 1/2 por estilo de plantilla si existe; si no, texto normal con preset de título
        doc <- officer::body_add_par(doc, value = plan[[i]]$slots$title, style = "heading 1")
        if (!is.null(plan[[i]]$slots$subtitle)) {
          doc <- officer::body_add_par(doc, value = plan[[i]]$slots$subtitle, style = "heading 2")
        }
      }

      log_rows[[i]] <- if (has_tibble) tibble::tibble(
        block_i = i, block_type = "section", element = NA_character_, var = NA_character_, png = NA_character_
      ) else list(block_i = i, block_type = "section", element = NA, var = NA, png = NA)

      next
    }

    # ---- BLOCK_1 ------------------------------------------------------------
    if (!identical(btype, "block_1")) {
      stop("Bloque no soportado en renderer Word simple: ", btype %||% "<NA>", call. = FALSE)
    }

    g_i <- g_i + 1L

    el <- plan[[i]]$plot %||% plan[[i]]$slots$main %||% NULL
    if (is.null(el) || !inherits(el, "ppt_element")) {
      stop("block_1 requiere `plot`/`slots$main` como `ppt_element`.", call. = FALSE)
    }

    p <- .render_element_plot(el)
    if (is.null(p)) stop("No se pudo renderizar `ppt_element` en block_1 (i=", i, ").", call. = FALSE)

    file_png <- file.path(tmp_dir, sprintf("graf_%03d.png", g_i))
    ok_png <- .plot_to_png(p, file_png, width_in = img_w, height_in = img_h, dpi = img_dpi, bg = img_bg)
    if (!isTRUE(ok_png) || !file.exists(file_png)) stop("No se pudo exportar PNG (i=", i, ").", call. = FALSE)

    rendered_files <- c(rendered_files, file_png)

    # título
    title_txt <- plan[[i]]$slots$title %||% NULL
    if (is.null(title_txt) || !nzchar(trimws(title_txt))) {
      title_txt <- .title_auto(el, g_i)
    } else {
      title_txt <- trimws(as.character(title_txt)[1])
      if (isTRUE(presets$figure_numbering$enabled) && isTRUE(presets$figure_numbering$prepend_on_manual)) {
        pref <- presets$figure_numbering$prefix %||% "Gráfico"
        sep  <- presets$figure_numbering$sep %||% ". "
        title_txt <- paste0(pref, " ", g_i, sep, title_txt)
      }
    }

    base_txt   <- plan[[i]]$slots$base   %||% NULL
    footer_txt <- plan[[i]]$slots$footer %||% NULL

    if (!isTRUE(solo_lista)) {

      # 1) título
      doc <- .add_par(doc, title_txt, presets$title_style)

      # 2) imagen
      doc <- officer::body_add_img(doc, src = file_png, width = img_w, height = img_h)

      # 3) base / footer
      if (!is.null(base_txt))   doc <- .add_par(doc, base_txt,   presets$base_style)
      if (!is.null(footer_txt)) doc <- .add_par(doc, footer_txt, presets$footer_style)

      # 4) salto opcional entre bloques (solo para block_1)
      if (isTRUE(presets$pagebreak_between) && i < length(plan)) {
        doc <- officer::body_add_break(doc)
      }
    }

    log_rows[[i]] <- if (has_tibble) tibble::tibble(
      block_i = i,
      block_type = "block_1",
      element = el$.element_type %||% NA_character_,
      var = el$var %||% (el$vars %||% NA_character_)[1],
      png = file_png
    ) else list(
      block_i = i,
      block_type = "block_1",
      element = el$.element_type %||% NA,
      var = el$var %||% (el$vars %||% NA)[1],
      png = file_png
    )
  }

  log <- if (has_dplyr && has_tibble) dplyr::bind_rows(log_rows) else log_rows

  if (!isTRUE(solo_lista)) {
    officer::print(doc, target = path_docx)
    if (isTRUE(mensajes_progreso)) message("DOCX generado en: ", normalizePath(path_docx, winslash = "/"))
  }

  # limpiar plan acumulado (si se usó w_add)
  if (exists(".word_plan_clear", mode = "function", inherits = TRUE)) {
    try(.word_plan_clear(env_bloques), silent = TRUE)
  }

  invisible(list(
    doc      = if (isTRUE(solo_lista)) NULL else doc,
    plan     = plan,
    rendered = rendered_files,
    log      = log
  ))
}
