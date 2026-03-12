# =============================================================================
# PLAN / SLIDES (contenedores con layout FIJO) — v2 (CONSISTENTE)
# - Elimina duplicados (p_slide_2pop/5pop/6pop) y unifica contratos.
# - Normaliza validaciones y textos opcionales.
# - Implementa p_slide_1_left / p_slide_1_right.
# - Corrige bug en p_numerico() (cruce = cruce).
# =============================================================================

# ---- Helpers internos --------------------------------------------------------

`%||%` <- function(x, y) if (!is.null(x)) x else y

.ppt_norm_text1 <- function(x, blank = NULL) {
  if (is.null(x)) return(NULL)
  x <- as.character(x)[1]
  if (!nzchar(trimws(x))) return(blank)
  x
}

.ppt_chk_meta <- function(meta) {
  if (!is.list(meta)) stop("`meta` debe ser una lista.", call. = FALSE)
  invisible(TRUE)
}

.ppt_chk_element <- function(x, nm) {
  if (is.null(x) || !inherits(x, "ppt_element")) {
    stop("`", nm, "` debe ser un `ppt_element` (p_*()).", call. = FALSE)
  }
  invisible(TRUE)
}

.ppt_chk_element_or_text <- function(x, nm) {
  ok <- inherits(x, c("ppt_element", "ppt_element_text")) ||
    (is.character(x) && length(x) == 1L)
  if (!ok) {
    stop("`", nm, "` debe ser un `p_*()` compatible o `character(1)`.", call. = FALSE)
  }
  invisible(TRUE)
}

.ppt_as_slide <- function(slide) {
  class(slide) <- c("ppt_slide", "list")
  slide
}

# =============================================================================
# PLAN
# =============================================================================

#' @title Construir un plan de diapositivas
#'
#' @description
#' Une objetos `p_slide_*()` en un plan ordenado. La entrada puede ser una lista
#' o argumentos sueltos.
#'
#' @param ... Objetos `ppt_slide`.
#' @param slides Alternativa a `...`: lista de slides.
#'
#' @return Lista de slides (plan) con clase `"ppt_plan"`.
#'
#' @export
p_plan <- function(..., slides = NULL) {
  out <- if (!is.null(slides)) {
    if (!is.list(slides)) stop("`slides` debe ser lista.", call. = FALSE)
    slides
  } else {
    list(...)
  }

  if (!length(out)) {
    out <- structure(list(), class = c("ppt_plan", "list"))
    return(out)
  }

  bad <- vapply(out, function(x) !inherits(x, "ppt_slide"), logical(1))
  if (any(bad)) {
    stop("`p_plan()`: todos los elementos deben ser `ppt_slide`. Malos: ",
         paste(which(bad), collapse = ", "), call. = FALSE)
  }

  class(out) <- c("ppt_plan", "list")
  out
}

# =============================================================================
# SLIDES — TÍTULO / SECCIÓN
# =============================================================================

#' @title Slide de cambio de sección
#'
#' @param title Título de la sección.
#' @param subtitle Subtítulo opcional.
#' @param meta Lista libre para notas internas (no se imprime).
#'
#' @return Objeto con clase `"ppt_slide"`.
#'
#' @export
p_slide_section <- function(title, subtitle = NULL, meta = list()) {
  title <- .ppt_norm_text1(title)
  if (is.null(title)) stop("`title` debe ser un texto no vacío.", call. = FALSE)

  subtitle <- .ppt_norm_text1(subtitle, blank = NULL)
  .ppt_chk_meta(meta)

  .ppt_as_slide(list(
    .slide_type = "section",
    title       = title,
    slots       = list(
      title    = title,
      subtitle = subtitle
    ),
    meta        = meta
  ))
}

#' @title Slide de portada (Title Slide)
#'
#' @param title Título (requerido).
#' @param subtitle Subtítulo opcional.
#' @param date Fecha opcional (si NULL: no tocar el placeholder de la plantilla).
#' @param meta_left,meta_right,meta_line Textos opcionales para placeholders secundarios.
#' @param meta Lista libre para notas internas (no se imprime).
#'
#' @return Objeto con clase `"ppt_slide"`.
#' @export
p_slide_title <- function(
    title,
    subtitle   = NULL,
    date       = NULL,
    meta_left  = NULL,
    meta_right = NULL,
    meta_line  = NULL,
    meta       = list()
) {
  title <- .ppt_norm_text1(title)
  if (is.null(title)) stop("`title` debe ser texto no vacío.", call. = FALSE)

  .ppt_chk_meta(meta)

  .ppt_as_slide(list(
    .slide_type = "title_slide",
    title       = title,
    slots       = list(
      title      = title,
      subtitle   = .ppt_norm_text1(subtitle,   blank = NULL),
      date       = .ppt_norm_text1(date,       blank = NULL),
      meta_left  = .ppt_norm_text1(meta_left,  blank = NULL),
      meta_right = .ppt_norm_text1(meta_right, blank = NULL),
      meta_line  = .ppt_norm_text1(meta_line,  blank = NULL)
    ),
    meta = meta
  ))
}

# =============================================================================
# SLIDES — BÁSICAS (1 gráfico / 2 gráficos)
# =============================================================================

#' @title Slide con 1 gráfico a pantalla completa
#'
#' @param title Título del slide (opcional).
#' @param subtitle Subtítulo opcional alineado bajo el título.
#' @param plot Elemento `p_*()` principal (requerido).
#' @param base Elemento opcional (p.ej. `p_text()` o `character(1)`).
#' @param footer Elemento opcional (p.ej. `p_text()` o `character(1)`).
#' @param meta Lista libre para notas internas.
#'
#' @return Objeto con clase `"ppt_slide"`.
#' @export
p_slide_1 <- function(title = NULL, subtitle = NULL, plot, base = NULL, footer = NULL, meta = list()) {
  .ppt_chk_element(plot, "plot")
  .ppt_chk_meta(meta)

  title <- .ppt_norm_text1(title, blank = NULL)
  subtitle <- .ppt_norm_text1(subtitle, blank = NULL)

  if (!is.null(base)) {
    .ppt_chk_element_or_text(base, "base")
    if (is.character(base)) base <- .ppt_norm_text1(base, blank = NULL)
  }

  if (!is.null(footer)) {
    .ppt_chk_element_or_text(footer, "footer")
    if (is.character(footer)) footer <- .ppt_norm_text1(footer, blank = NULL)
  }

  .ppt_as_slide(list(
    .slide_type = "slide_1",
    title       = title,
    slots       = list(
      title    = title,
      subtitle = subtitle,
      plot     = plot,
      base     = base,
      footer   = footer
    ),
    meta        = meta
  ))
}

#' @title Slide con 2 gráficos lado a lado
#'
#' @param title Título del slide (opcional).
#' @param left Elemento `p_*()` izquierda (requerido).
#' @param right Elemento `p_*()` derecha (requerido).
#' @param base Elemento opcional (p.ej. `p_text()` / `character(1)`).
#' @param footer Texto o elemento opcional para la caja derecha del layout.
#' @param meta Lista libre.
#'
#' @return Objeto con clase `"ppt_slide"`.
#' @export
p_slide_2 <- function(title = NULL, left, right, base = NULL, footer = NULL, meta = list()) {
  .ppt_chk_element(left,  "left")
  .ppt_chk_element(right, "right")
  .ppt_chk_meta(meta)

  title <- .ppt_norm_text1(title, blank = NULL)

  if (!is.null(base)) {
    .ppt_chk_element_or_text(base, "base")
    if (is.character(base)) base <- .ppt_norm_text1(base, blank = NULL)
  }

  if (!is.null(footer)) {
    .ppt_chk_element_or_text(footer, "footer")
    if (is.character(footer)) footer <- .ppt_norm_text1(footer, blank = NULL)
  }

  .ppt_as_slide(list(
    .slide_type = "slide_2",
    title       = title,
    slots       = list(
      title      = title,
      left       = left,
      right      = right,
      base       = base,
      footer     = footer,
      right_text = if (inherits(footer, "ppt_element_text")) {
        footer$text %||% NULL
      } else if (is.character(footer)) {
        footer
      } else {
        NULL
      }
    ),
    meta = meta
  ))
}

# =============================================================================
# SLIDES — TEXTO + GRÁFICO(S) (contrato consistente)
# =============================================================================

#' @export
p_slide_text_r <- function(
    title = NULL,
    plot,
    text = "Lorem ipsum…",
    tag = NULL,
    base = NULL,
    footer = NULL,
    meta = list()
) {
  .ppt_chk_element(plot, "plot")
  .ppt_chk_meta(meta)

  title <- .ppt_norm_text1(title, blank = NULL)

  if (!is.null(base)) {
    .ppt_chk_element_or_text(base, "base")
    if (is.character(base)) base <- .ppt_norm_text1(base, blank = NULL)
  }
  if (!is.null(footer)) {
    .ppt_chk_element_or_text(footer, "footer")
    if (is.character(footer)) footer <- .ppt_norm_text1(footer, blank = NULL)
  }

  .ppt_as_slide(list(
    .slide_type = "text_r",
    title       = title,
    slots       = list(
      title  = title,
      plot   = plot,
      text   = .ppt_norm_text1(text, blank = " "),
      tag    = .ppt_norm_text1(tag,  blank = NULL),
      base   = base,
      footer = footer
    ),
    meta = meta
  ))
}

#' @export
p_slide_text_l <- function(
    title = NULL,
    plot,
    text = "Lorem ipsum…",
    tag = NULL,
    base = NULL,
    footer = NULL,
    meta = list()
) {
  .ppt_chk_element(plot, "plot")
  .ppt_chk_meta(meta)

  title <- .ppt_norm_text1(title, blank = NULL)

  if (!is.null(base)) {
    .ppt_chk_element_or_text(base, "base")
    if (is.character(base)) base <- .ppt_norm_text1(base, blank = NULL)
  }
  if (!is.null(footer)) {
    .ppt_chk_element_or_text(footer, "footer")
    if (is.character(footer)) footer <- .ppt_norm_text1(footer, blank = NULL)
  }

  .ppt_as_slide(list(
    .slide_type = "text_l",
    title       = title,
    slots       = list(
      title  = title,
      plot   = plot,
      text   = .ppt_norm_text1(text, blank = " "),
      tag    = .ppt_norm_text1(tag,  blank = NULL),
      base   = base,
      footer = footer
    ),
    meta = meta
  ))
}

#' @export
p_slide_text_r2 <- function(
    title = NULL,
    plot1,
    plot2,
    text = "Lorem ipsum…",
    tag = NULL,
    base = NULL,
    footer = NULL,
    meta = list()
) {
  .ppt_chk_element(plot1, "plot1")
  .ppt_chk_element(plot2, "plot2")
  .ppt_chk_meta(meta)

  title <- .ppt_norm_text1(title, blank = NULL)

  if (!is.null(base)) {
    .ppt_chk_element_or_text(base, "base")
    if (is.character(base)) base <- .ppt_norm_text1(base, blank = NULL)
  }
  if (!is.null(footer)) {
    .ppt_chk_element_or_text(footer, "footer")
    if (is.character(footer)) footer <- .ppt_norm_text1(footer, blank = NULL)
  }

  .ppt_as_slide(list(
    .slide_type = "text_r2",
    title       = title,
    slots       = list(
      title  = title,
      plot1  = plot1,
      plot2  = plot2,
      text   = .ppt_norm_text1(text, blank = " "),
      tag    = .ppt_norm_text1(tag,  blank = NULL),
      base   = base,
      footer = footer
    ),
    meta = meta
  ))
}

#' @export
p_slide_text_l2 <- function(
    title = NULL,
    plot1,
    plot2,
    text = "Lorem ipsum…",
    tag = NULL,
    base = NULL,
    footer = NULL,
    meta = list()
) {
  .ppt_chk_element(plot1, "plot1")
  .ppt_chk_element(plot2, "plot2")
  .ppt_chk_meta(meta)

  title <- .ppt_norm_text1(title, blank = NULL)

  if (!is.null(base)) {
    .ppt_chk_element_or_text(base, "base")
    if (is.character(base)) base <- .ppt_norm_text1(base, blank = NULL)
  }
  if (!is.null(footer)) {
    .ppt_chk_element_or_text(footer, "footer")
    if (is.character(footer)) footer <- .ppt_norm_text1(footer, blank = NULL)
  }

  .ppt_as_slide(list(
    .slide_type = "text_l2",
    title       = title,
    slots       = list(
      title  = title,
      plot1  = plot1,
      plot2  = plot2,
      text   = .ppt_norm_text1(text, blank = " "),
      tag    = .ppt_norm_text1(tag,  blank = NULL),
      base   = base,
      footer = footer
    ),
    meta = meta
  ))
}

# =============================================================================
# SLIDES — POBLACIÓN (CONTRATO CANÓNICO)
# - p_slide_poblacion_2 / _4 / _5 / _6 son el estándar público.
# =============================================================================

#' @export
p_slide_poblacion_2 <- function(
    title = NULL,
    left,
    right,
    tag = NULL,
    center_note = "",
    base = NULL,
    footer = NULL,
    meta = list()
) {
  .ppt_chk_element(left,  "left")
  .ppt_chk_element(right, "right")
  .ppt_chk_meta(meta)

  title <- .ppt_norm_text1(title, blank = NULL)

  if (!is.null(base)) {
    .ppt_chk_element_or_text(base, "base")
    if (is.character(base)) base <- .ppt_norm_text1(base, blank = NULL)
  }
  if (!is.null(footer)) {
    .ppt_chk_element_or_text(footer, "footer")
    if (is.character(footer)) footer <- .ppt_norm_text1(footer, blank = NULL)
  }

  .ppt_as_slide(list(
    .slide_type = "poblacion_2",
    title       = title,
    slots       = list(
      title       = title,
      left        = left,
      right       = right,
      tag         = .ppt_norm_text1(tag,         blank = NULL),
      center_note = .ppt_norm_text1(center_note, blank = " "),
      base        = base,
      footer      = footer
    ),
    meta = meta
  ))
}

#' @export
p_slide_poblacion_4 <- function(
    title = NULL,
    up_left, up_right, bottom_left, bottom_right,
    tag = NULL,
    center_note = "",
    base = NULL,
    footer = NULL,
    meta = list()
) {
  .ppt_chk_element(up_left,      "up_left")
  .ppt_chk_element(up_right,     "up_right")
  .ppt_chk_element(bottom_left,  "bottom_left")
  .ppt_chk_element(bottom_right, "bottom_right")
  .ppt_chk_meta(meta)

  title <- .ppt_norm_text1(title, blank = NULL)

  if (!is.null(base)) {
    .ppt_chk_element_or_text(base, "base")
    if (is.character(base)) base <- .ppt_norm_text1(base, blank = NULL)
  }
  if (!is.null(footer)) {
    .ppt_chk_element_or_text(footer, "footer")
    if (is.character(footer)) footer <- .ppt_norm_text1(footer, blank = NULL)
  }

  .ppt_as_slide(list(
    .slide_type = "poblacion_4",
    title       = title,
    slots       = list(
      title        = title,
      up_left      = up_left,
      up_right     = up_right,
      bottom_left  = bottom_left,
      bottom_right = bottom_right,
      tag          = .ppt_norm_text1(tag,         blank = NULL),
      center_note  = .ppt_norm_text1(center_note, blank = " "),
      base         = base,
      footer       = footer
    ),
    meta = meta
  ))
}

#' @export
p_slide_poblacion_5 <- function(
    title = NULL,
    pic1, pic2, pic3, pic4, pic5,
    tag = NULL,
    icon = NULL,
    footer = NULL,
    meta = list()
) {
  .ppt_chk_element(pic1, "pic1")
  .ppt_chk_element(pic2, "pic2")
  .ppt_chk_element(pic3, "pic3")
  .ppt_chk_element(pic4, "pic4")
  .ppt_chk_element(pic5, "pic5")
  .ppt_chk_meta(meta)

  title <- .ppt_norm_text1(title, blank = NULL)
  tag   <- .ppt_norm_text1(tag,   blank = NULL)

  footer_txt <- paste(
    c(.ppt_norm_text1(icon, blank = NULL), .ppt_norm_text1(footer, blank = NULL)),
    collapse = "\n"
  )
  if (!nzchar(trimws(footer_txt))) footer_txt <- NULL

  .ppt_as_slide(list(
    .slide_type = "poblacion_5",
    title       = title,
    slots       = list(
      title  = title,
      tag    = tag,
      icon   = NULL,        # si NO hay placeholder propio para icon, mantener NULL
      footer = footer_txt,

      pic1 = pic1,
      pic2 = pic2,
      pic3 = pic3,
      pic4 = pic4,
      pic5 = pic5
    ),
    meta = meta
  ))
}

#' @export
p_slide_poblacion_6 <- function(
    title = NULL,
    pic1, pic2, pic3, pic4, pic5, pic6,
    tag = NULL,
    icon = NULL,
    footer = NULL,
    meta = list()
) {
  .ppt_chk_element(pic1, "pic1")
  .ppt_chk_element(pic2, "pic2")
  .ppt_chk_element(pic3, "pic3")
  .ppt_chk_element(pic4, "pic4")
  .ppt_chk_element(pic5, "pic5")
  .ppt_chk_element(pic6, "pic6")
  .ppt_chk_meta(meta)

  title <- .ppt_norm_text1(title, blank = NULL)
  tag   <- .ppt_norm_text1(tag,   blank = NULL)

  footer_txt <- paste(
    c(.ppt_norm_text1(icon, blank = NULL), .ppt_norm_text1(footer, blank = NULL)),
    collapse = "\n"
  )
  if (!nzchar(trimws(footer_txt))) footer_txt <- NULL

  .ppt_as_slide(list(
    .slide_type = "poblacion_6",
    title       = title,
    slots       = list(
      title  = title,
      tag    = tag,
      icon   = NULL,
      footer = footer_txt,

      pic1 = pic1,
      pic2 = pic2,
      pic3 = pic3,
      pic4 = pic4,
      pic5 = pic5,
      pic6 = pic6
    ),
    meta = meta
  ))
}

# =============================================================================
# ELEMENTOS p_* (objetos declarativos)
# =============================================================================

#' @keywords internal
.ppt_norm_filters <- function(filtros) {
  if (is.null(filtros)) return(list())
  if (!is.list(filtros)) stop("`filtros` debe ser lista.", call. = FALSE)

  nms <- names(filtros)
  if (length(filtros) && is.null(nms)) {
    stop("`filtros` debe ser una lista nombrada por variable.", call. = FALSE)
  }
  if (!is.null(nms)) {
    if (any(!nzchar(trimws(nms)))) {
      stop("`filtros` debe ser una lista nombrada por variable.", call. = FALSE)
    }
    names(filtros) <- trimws(nms)
  }

  filtros
}

#' @title Barras agrupadas (1 variable)
#' @param filtros Lista nombrada de filtros por igualdad/inclusión,
#'   por ejemplo `list(region = "Lima", sexo = c("Mujer", "Otro"))`.
#' @examples
#' p_barras_agrupadas("p102", filtros = list(region = "Lima"))
#' @export
p_barras_agrupadas <- function(var, titulo = NULL, cruces = NULL, overrides = list(), base = list(), filtros = list()) {
  if (!is.character(var) || length(var) != 1L || !nzchar(trimws(var))) {
    stop("`var` debe ser character(1) no vacío.", call. = FALSE)
  }
  var <- trimws(var)

  titulo <- .ppt_norm_text1(titulo, blank = NULL)

  if (!is.null(cruces)) {
    if (!is.character(cruces) || length(cruces) != 1L || !nzchar(trimws(cruces))) {
      stop("`cruces` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    cruces <- trimws(cruces)
  }

  if (!is.list(overrides)) stop("`overrides` debe ser lista.", call. = FALSE)
  if (!is.list(base)) stop("`base` debe ser lista.", call. = FALSE)
  filtros <- .ppt_norm_filters(filtros)

  el <- list(
    .element_type = "barras_agrupadas",
    var           = var,
    title_slide   = titulo,
    cruces        = cruces,
    overrides     = overrides,
    base          = base,
    filtros       = filtros
  )
  class(el) <- c("ppt_element", "list")
  el
}

#' @title Barras apiladas (1 variable)
#' @param filtros Lista nombrada de filtros por igualdad/inclusión,
#'   por ejemplo `list(region = "Lima", sexo = c("Mujer", "Otro"))`.
#' @examples
#' p_barras_apiladas("p102", filtros = list(region = "Lima"))
#' @export
p_barras_apiladas <- function(var, titulo = NULL, cruces = NULL, overrides = list(), base = list(), filtros = list()) {
  if (!is.character(var) || length(var) != 1L || !nzchar(trimws(var))) {
    stop("`var` debe ser character(1) no vacío.", call. = FALSE)
  }
  var <- trimws(var)

  titulo <- .ppt_norm_text1(titulo, blank = NULL)

  if (!is.null(cruces)) {
    if (!is.character(cruces) || length(cruces) != 1L || !nzchar(trimws(cruces))) {
      stop("`cruces` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    cruces <- trimws(cruces)
  }

  if (!is.list(overrides)) stop("`overrides` debe ser lista.", call. = FALSE)
  if (!is.list(base)) stop("`base` debe ser lista.", call. = FALSE)
  filtros <- .ppt_norm_filters(filtros)

  el <- list(
    .element_type = "barras_apiladas",
    var           = var,
    title_slide   = titulo,
    cruces        = cruces,
    overrides     = overrides,
    base          = base,
    filtros       = filtros
  )
  class(el) <- c("ppt_element", "list")
  el
}

#' @title Barras multi-apiladas (varias variables o 1 variable cruzada)
#' @param filtros Lista nombrada de filtros por igualdad/inclusión.
#' @param bloques En `modo = "multilista"`, lista de bloques. Cada bloque debe
#'   ser una lista con al menos `modo` (`"var"`, `"cruce"` o `"var_cruce"`),
#'   y los argumentos necesarios para ese submodo (`vars`, `var`, `cruces`,
#'   `titulos_grupo`, etc.). Cada bloque puede incluir opcionalmente
#'   `altura_rel`, `overrides`, `base` y `filtros`.
#' @examples
#' p_barras_multiapiladas(
#'   modo = "cruce",
#'   var = "p102",
#'   cruces = "region",
#'   filtros = list(sexo = "Mujer")
#' )
#'
#' En `modo = "var_cruce"`, `vars` también puede ser una lista nombrada de
#' bloques. Cada bloque debe contener referencias `fuente$variable` cuando se
#' comparan varias bases en un mismo gráfico.
#'
#' En `modo = "multilista"`, se pueden apilar varios bloques con distintas
#' escalas dentro de una sola composición vertical.
#' @export
p_barras_multiapiladas <- function(
    modo = c("var", "cruce", "var_cruce", "multilista"),
    vars = NULL,
    bloques = NULL,
    var  = NULL,
    titulo = NULL,
    cruces = NULL,
    wrap_y = 50,
    top2box        = FALSE,
    top2box_codes  = NULL,
    top2box_labels = NULL,
    titulos_grupo  = NULL,
    overrides = list(),
    base = list(),
    filtros = list()
) {
  modo <- match.arg(modo)
  titulo <- .ppt_norm_text1(titulo, blank = NULL)

  if (!is.null(cruces)) {
    if (!is.character(cruces) || length(cruces) != 1L || !nzchar(trimws(cruces))) {
      stop("`cruces` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    cruces <- trimws(cruces)
  }

  if (!is.numeric(wrap_y) || length(wrap_y) != 1L || !is.finite(wrap_y) || wrap_y < 10) {
    stop("`wrap_y` debe ser numérico (>=10).", call. = FALSE)
  }

  # NOTE: top2box_codes se deja por compatibilidad futura; por ahora se usa top2box_labels
  if (!is.logical(top2box) || length(top2box) != 1L || is.na(top2box)) {
    stop("`top2box` debe ser logical(1).", call. = FALSE)
  }
  if (!is.null(top2box_labels)) {
    if (!is.character(top2box_labels) || !length(top2box_labels)) {
      stop("`top2box_labels` debe ser NULL o character() no vacío.", call. = FALSE)
    }
    top2box_labels <- trimws(top2box_labels)
    top2box_labels <- top2box_labels[nzchar(top2box_labels)]
    if (!length(top2box_labels)) top2box_labels <- NULL
  }
  if (!is.null(titulos_grupo)) {
    if (!is.character(titulos_grupo) || !length(titulos_grupo)) {
      stop("`titulos_grupo` debe ser NULL o character() no vacío.", call. = FALSE)
    }
    titulos_grupo <- trimws(titulos_grupo)
    titulos_grupo <- titulos_grupo[nzchar(titulos_grupo)]
    if (!length(titulos_grupo)) {
      titulos_grupo <- NULL
    } else if (is.null(names(titulos_grupo)) || any(!nzchar(trimws(names(titulos_grupo))))) {
      stop("`titulos_grupo` debe ser un vector nombrado por variable.", call. = FALSE)
    } else {
      names(titulos_grupo) <- trimws(names(titulos_grupo))
      titulos_grupo <- titulos_grupo[nzchar(names(titulos_grupo))]
      if (!length(titulos_grupo)) titulos_grupo <- NULL
    }
  }

  if (!is.list(overrides)) stop("`overrides` debe ser una lista.", call. = FALSE)
  if (!is.list(base)) stop("`base` debe ser una lista.", call. = FALSE)
  filtros <- .ppt_norm_filters(filtros)

  if (identical(modo, "multilista")) {
    if (!is.list(bloques) || !length(bloques)) {
      stop("modo='multilista': `bloques` debe ser una lista no vacía.", call. = FALSE)
    }

    bloques_norm <- lapply(seq_along(bloques), function(i) {
      block <- bloques[[i]]
      if (!is.list(block)) {
        stop("modo='multilista': cada bloque debe ser una lista.", call. = FALSE)
      }

      modo_block <- block[["modo", exact = TRUE]] %||% NULL
      if (!is.character(modo_block) || length(modo_block) != 1L || !nzchar(trimws(modo_block))) {
        stop("modo='multilista': cada bloque debe definir `modo`.", call. = FALSE)
      }
      modo_block <- trimws(modo_block)
      if (identical(modo_block, "multilista")) {
        stop("modo='multilista': no se permiten bloques anidados de tipo `multilista`.", call. = FALSE)
      }

      filtros_block <- utils::modifyList(filtros, .ppt_norm_filters(block[["filtros", exact = TRUE]] %||% list()))
      base_block <- utils::modifyList(base, block[["base", exact = TRUE]] %||% list())
      overrides_block <- utils::modifyList(overrides, block[["overrides", exact = TRUE]] %||% list())

      titulo_block <- .ppt_norm_text1(block[["titulo", exact = TRUE]] %||% NULL, blank = NULL)
      subtitulo_block <- .ppt_norm_text1(block[["subtitulo", exact = TRUE]] %||% NULL, blank = NULL)

      # En multilista, por defecto los subbloques NO deben heredar títulos
      # automáticos ni desde presets ni desde otros overrides. Solo se muestran
      # si el usuario los define explícitamente en el bloque.
      overrides_block$titulo <- titulo_block %||% ""
      overrides_block$subtitulo <- subtitulo_block %||% ""

      child <- p_barras_multiapiladas(
        modo = modo_block,
        vars = block[["vars", exact = TRUE]] %||% NULL,
        bloques = NULL,
        var = block[["var", exact = TRUE]] %||% NULL,
        titulo = titulo_block,
        cruces = block[["cruces", exact = TRUE]] %||% NULL,
        wrap_y = block[["wrap_y", exact = TRUE]] %||% wrap_y,
        top2box = block[["top2box", exact = TRUE]] %||% FALSE,
        top2box_codes = block[["top2box_codes", exact = TRUE]] %||% NULL,
        top2box_labels = block[["top2box_labels", exact = TRUE]] %||% NULL,
        titulos_grupo = block[["titulos_grupo", exact = TRUE]] %||% NULL,
        overrides = overrides_block,
        base = base_block,
        filtros = filtros_block
      )
      child$title_slide <- NULL
      child$.multilista_block_title <- titulo_block
      child$.multilista_block_subtitle <- subtitulo_block
      child$altura_rel <- block[["altura_rel", exact = TRUE]] %||% NULL
      child
    })

    el <- list(
      .element_type  = "barras_multiapiladas",
      modo           = "multilista",
      bloques        = bloques_norm,
      vars           = NULL,
      var            = NULL,
      cruce          = NULL,
      title_slide    = titulo,
      wrap_y         = wrap_y,
      top2box        = FALSE,
      top2box_codes  = NULL,
      top2box_labels = NULL,
      titulos_grupo  = NULL,
      overrides      = overrides,
      base           = base,
      filtros        = filtros
    )
    class(el) <- c("ppt_element", "list")
    return(el)
  }

  if (identical(modo, "var")) {
    if (is.null(vars)) stop("modo='var': `vars` no puede ser NULL.", call. = FALSE)
    if (!is.character(vars) || length(vars) < 1L) stop("modo='var': `vars` debe ser character() con >= 1 variable.", call. = FALSE)
    vars <- trimws(vars)
    vars <- vars[nzchar(vars)]
    if (!length(vars)) stop("modo='var': `vars` quedó vacío luego de limpiar.", call. = FALSE)

    el <- list(
      .element_type  = "barras_multiapiladas",
      modo           = "var",
      vars           = vars,
      var            = NULL,
      cruce          = cruces,
      title_slide    = titulo,
      wrap_y         = wrap_y,
      top2box        = isTRUE(top2box),
      top2box_codes  = top2box_codes,
      top2box_labels = top2box_labels,
      titulos_grupo  = NULL,
      overrides      = overrides,
      base           = base,
      filtros        = filtros
    )
    class(el) <- c("ppt_element", "list")
    return(el)
  }

  if (identical(modo, "var_cruce")) {
    if (is.null(vars)) stop("modo='var_cruce': `vars` no puede ser NULL.", call. = FALSE)
    if (is.character(vars)) {
      if (length(vars) < 1L) stop("modo='var_cruce': `vars` debe ser character() con >= 1 variable.", call. = FALSE)
      vars <- trimws(vars)
      vars <- vars[nzchar(vars)]
      if (!length(vars)) stop("modo='var_cruce': `vars` quedó vacío luego de limpiar.", call. = FALSE)

      if (is.null(cruces)) stop("modo='var_cruce': `cruces` es obligatorio (character(1)).", call. = FALSE)
    } else if (is.list(vars)) {
      if (!length(vars)) stop("modo='var_cruce': `vars` no puede ser una lista vacía.", call. = FALSE)
      if (is.null(names(vars)) || any(!nzchar(trimws(names(vars))))) {
        stop("modo='var_cruce': cuando `vars` es lista, debe ser una lista nombrada.", call. = FALSE)
      }
      names(vars) <- trimws(names(vars))
      vars <- vars[nzchar(names(vars))]
      if (!length(vars)) stop("modo='var_cruce': `vars` quedó vacío luego de limpiar.", call. = FALSE)

      vars <- lapply(vars, function(x) {
        if (!is.character(x) || !length(x)) {
          stop("modo='var_cruce': cada bloque de `vars` debe ser character() no vacío.", call. = FALSE)
        }
        x <- trimws(x)
        x <- x[nzchar(x)]
        if (!length(x)) {
          stop("modo='var_cruce': un bloque de `vars` quedó vacío luego de limpiar.", call. = FALSE)
        }
        x
      })
    } else {
      stop("modo='var_cruce': `vars` debe ser character() o lista nombrada.", call. = FALSE)
    }

    el <- list(
      .element_type  = "barras_multiapiladas",
      modo           = "var_cruce",
      vars           = vars,
      var            = NULL,
      cruce          = cruces,
      title_slide    = titulo,
      wrap_y         = wrap_y,
      top2box        = isTRUE(top2box),
      top2box_codes  = top2box_codes,
      top2box_labels = top2box_labels,
      titulos_grupo  = titulos_grupo,
      overrides      = overrides,
      base           = base,
      filtros        = filtros
    )
    class(el) <- c("ppt_element", "list")
    return(el)
  }

  # modo == "cruce"
  if (!is.character(var) || length(var) != 1L || !nzchar(trimws(var))) {
    stop("modo='cruce': `var` debe ser character(1) no vacío.", call. = FALSE)
  }
  var <- trimws(var)

  if (is.null(cruces)) stop("modo='cruce': `cruces` es obligatorio (character(1)).", call. = FALSE)

  el <- list(
    .element_type  = "barras_multiapiladas",
    modo           = "cruce",
    vars           = NULL,
    var            = var,
    cruce          = cruces,
    title_slide    = titulo,
    wrap_y         = wrap_y,
    top2box        = isTRUE(top2box),
    top2box_codes  = top2box_codes,
    top2box_labels = top2box_labels,
    titulos_grupo  = NULL,
    overrides      = overrides,
    base           = base,
    filtros        = filtros
  )
  class(el) <- c("ppt_element", "list")
  el
}



#' @title Pie (torta)
#' @param filtros Lista nombrada de filtros por igualdad/inclusión.
#' @examples
#' p_pie("p108", filtros = list(sexo = "Mujer", edad_grupo = c("60-69", "70+")))
#' @export
p_pie <- function(var, titulo = NULL, overrides = list(), base = list(), filtros = list()) {
  if (!is.character(var) || length(var) != 1L || !nzchar(trimws(var))) {
    stop("`var` debe ser character(1) no vacío.", call. = FALSE)
  }
  var <- trimws(var)

  titulo <- .ppt_norm_text1(titulo, blank = NULL)

  if (!is.list(overrides)) stop("`overrides` debe ser lista.", call. = FALSE)
  if (!is.list(base)) stop("`base` debe ser lista.", call. = FALSE)
  filtros <- .ppt_norm_filters(filtros)

  el <- list(
    .element_type = "pie",
    var           = var,
    title_slide   = titulo,
    overrides     = overrides,
    base          = base,
    filtros       = filtros
  )
  class(el) <- c("ppt_element", "list")
  el
}

#' @title Donut
#' @param filtros Lista nombrada de filtros por igualdad/inclusión.
#' @export
p_donut <- function(var, titulo = NULL, overrides = list(), base = list(), filtros = list()) {
  if (!is.character(var) || length(var) != 1L || !nzchar(trimws(var))) {
    stop("`var` debe ser character(1) no vacío.", call. = FALSE)
  }
  var <- trimws(var)

  titulo <- .ppt_norm_text1(titulo, blank = NULL)

  if (!is.list(overrides)) stop("`overrides` debe ser lista.", call. = FALSE)
  if (!is.list(base)) stop("`base` debe ser lista.", call. = FALSE)
  filtros <- .ppt_norm_filters(filtros)

  el <- list(
    .element_type = "donut",
    var           = var,
    title_slide   = titulo,
    overrides     = overrides,
    base          = base,
    filtros       = filtros
  )
  class(el) <- c("ppt_element", "list")
  el
}

#' @title KPI numérico
#'
#' @param var Variable base (opcional según métrica).
#' @param metrica "N", "pct", "mean", "median".
#' @param cruce Variable opcional de cruce (si el renderer lo soporta).
#' @param titulo Título opcional.
#' @param formato Formato de salida (p.ej. `"%.0f%%"`).
#' @param overrides Lista de overrides (p.ej. `fn`, `denom`, `na_rm`).
#' @param filtros Lista nombrada de filtros por igualdad/inclusión.
#' @examples
#' p_numerico("p118_tbc_a", cruce = "region", filtros = list(sexo = "Mujer"))
#'
#' @return Objeto `"ppt_element"`.
#' @export
p_numerico <- function(
    var = NULL,
    metrica = c("N", "pct", "mean", "median"),
    cruce = NULL,
    titulo = NULL,
    formato = NULL,
    overrides = list(),
    filtros = list()
) {
  metrica <- match.arg(metrica)

  if (!is.null(var)) {
    if (!is.character(var) || length(var) != 1L || !nzchar(trimws(var))) {
      stop("`var` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    var <- trimws(var)
  }

  if (!is.null(cruce)) {
    if (!is.character(cruce) || length(cruce) != 1L || !nzchar(trimws(cruce))) {
      stop("`cruce` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    cruce <- trimws(cruce)
  }

  titulo  <- .ppt_norm_text1(titulo,  blank = NULL)
  formato <- .ppt_norm_text1(formato, blank = NULL)

  if (!is.list(overrides)) stop("`overrides` debe ser lista.", call. = FALSE)
  filtros <- .ppt_norm_filters(filtros)

  el <- list(
    .element_type = "numerico",
    var           = var,
    metrica       = metrica,
    cruce         = cruce,
    title_slide   = titulo,
    formato       = formato,
    overrides     = overrides,
    filtros       = filtros
  )
  class(el) <- c("ppt_element", "list")
  el
}

#' @title Radar + tabla derecha (SM o Top/Bottom 2 Box)
#' @param filtros Lista nombrada de filtros por igualdad/inclusión.
#' @export
p_radar_tabla <- function(
    modo = c("sm", "box"),
    var  = NULL,
    vars = NULL,
    cruce = NULL,
    box_labels = NULL,
    titulo_tabla = NULL,
    colores_series = NULL,
    titulo = NULL,
    top_n = NULL,
    sm_omit_codes  = NULL,
    sm_omit_labels = NULL,
    sm_omit_na     = TRUE,
    overrides = list(),
    base = list(),
    filtros = list()
) {
  modo <- match.arg(modo)

  if (identical(modo, "sm")) {
    if (!is.character(var) || length(var) != 1L || !nzchar(trimws(var))) {
      stop("p_radar_tabla(modo='sm'): `var` debe ser character(1) no vacío.", call. = FALSE)
    }
    var <- trimws(var)
    if (!is.null(vars)) stop("p_radar_tabla(modo='sm'): no usar `vars`.", call. = FALSE)
  }

  if (identical(modo, "box")) {
    if (is.character(vars)) {
      if (length(vars) < 1L) {
        stop("p_radar_tabla(modo='box'): `vars` debe ser character() con >=1 variable.", call. = FALSE)
      }
      vars <- trimws(vars); vars <- vars[nzchar(vars)]
      if (!length(vars)) stop("p_radar_tabla(modo='box'): `vars` quedó vacío.", call. = FALSE)
    } else if (is.list(vars)) {
      if (!length(vars)) {
        stop("p_radar_tabla(modo='box'): `vars` no puede ser una lista vacía.", call. = FALSE)
      }
      if (is.null(names(vars)) || any(!nzchar(trimws(names(vars))))) {
        stop("p_radar_tabla(modo='box'): cuando `vars` es lista, debe ser una lista nombrada.", call. = FALSE)
      }
      names(vars) <- trimws(names(vars))
      vars <- vars[nzchar(names(vars))]
      if (!length(vars)) stop("p_radar_tabla(modo='box'): `vars` quedó vacío luego de limpiar.", call. = FALSE)

      vars <- lapply(vars, function(x) {
        if (!is.character(x) || !length(x)) {
          stop("p_radar_tabla(modo='box'): cada bloque de `vars` debe ser character() no vacío.", call. = FALSE)
        }
        x <- trimws(x)
        x <- x[nzchar(x)]
        if (!length(x)) {
          stop("p_radar_tabla(modo='box'): un bloque de `vars` quedó vacío luego de limpiar.", call. = FALSE)
        }
        x
      })
    } else {
      stop("p_radar_tabla(modo='box'): `vars` debe ser character() o lista nombrada.", call. = FALSE)
    }

    if (!is.null(var)) stop("p_radar_tabla(modo='box'): no usar `var`.", call. = FALSE)

    if (!is.character(box_labels) || length(box_labels) != 2L) {
      stop("p_radar_tabla(modo='box'): `box_labels` debe ser character(2).", call. = FALSE)
    }
    box_labels <- as.character(box_labels)
  }

  if (!is.null(cruce)) {
    if (!is.character(cruce) || length(cruce) != 1L || !nzchar(trimws(cruce))) {
      stop("`cruce` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    cruce <- trimws(cruce)
  }

  titulo <- .ppt_norm_text1(titulo, blank = NULL)

  if (!is.null(top_n)) {
    if (!is.numeric(top_n) || length(top_n) != 1L || !is.finite(top_n) || top_n < 3) {
      stop("`top_n` debe ser numérico >= 3 (o NULL).", call. = FALSE)
    }
    top_n <- as.integer(top_n)
  }

  if (!is.list(overrides)) stop("`overrides` debe ser lista.", call. = FALSE)
  if (!is.list(base)) stop("`base` debe ser lista.", call. = FALSE)
  if (!is.null(colores_series)) {
    if (!is.atomic(colores_series) || is.null(names(colores_series))) {
      stop("`colores_series` debe ser NULL o un vector nombrado.", call. = FALSE)
    }
  }
  filtros <- .ppt_norm_filters(filtros)

  if (is.null(titulo_tabla) || !nzchar(trimws(as.character(titulo_tabla)))) {
    titulo_tabla <- if (identical(modo, "sm")) "Opciones" else "Top 2 Box"
  }

  el <- list(
    .element_type   = "radar_tabla",
    modo            = modo,
    var             = var,
    vars            = vars,
    cruce           = cruce,
    box_labels      = box_labels,
    colores_series  = colores_series,
    sm_omit_codes   = sm_omit_codes,
    sm_omit_labels  = sm_omit_labels,
    sm_omit_na      = sm_omit_na,
    titulo_tabla    = as.character(titulo_tabla)[1],
    title_slide     = titulo,
    top_n           = top_n,
    overrides       = overrides,
    base            = base,
    filtros         = filtros
  )
  class(el) <- c("ppt_element", "list")
  el
}

#' @title Heatmap de dimensiones
#' @export
p_dim_heatmap <- function(
    modo = c("general", "indicadores"),
    objetivo,
    cruce = NULL,
    incluir_total = NULL,
    filtros = list(),
    iter_var = NULL,
    iter_level = NULL,
    titulo = NULL,
    overrides = list(),
    base = list()
) {
  modo <- match.arg(modo)

  if (!is.character(objetivo) || length(objetivo) != 1L || !nzchar(trimws(objetivo))) {
    stop("`objetivo` debe ser character(1) no vacío.", call. = FALSE)
  }
  objetivo <- trimws(objetivo)

  if (!is.null(cruce)) {
    if (!is.character(cruce) || length(cruce) != 1L || !nzchar(trimws(cruce))) {
      stop("`cruce` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    cruce <- trimws(cruce)
  }

  if (!is.null(iter_var)) {
    if (!is.character(iter_var) || length(iter_var) != 1L || !nzchar(trimws(iter_var))) {
      stop("`iter_var` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    iter_var <- trimws(iter_var)
  }

  if (!is.null(iter_level)) {
    if (!is.character(iter_level) || length(iter_level) != 1L || !nzchar(trimws(iter_level))) {
      stop("`iter_level` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    iter_level <- trimws(iter_level)
  }

  if (!is.null(incluir_total)) {
    if (!is.logical(incluir_total) || length(incluir_total) != 1L || is.na(incluir_total)) {
      stop("`incluir_total` debe ser NULL o logical(1).", call. = FALSE)
    }
  }

  filtros <- .ppt_norm_filters(filtros)
  if (!is.list(overrides)) stop("`overrides` debe ser lista.", call. = FALSE)
  if (!is.list(base)) stop("`base` debe ser lista.", call. = FALSE)

  el <- list(
    .element_type = "dim_heatmap",
    modo = modo,
    objetivo = objetivo,
    cruce = cruce,
    incluir_total = incluir_total,
    filtros = filtros,
    iter_var = iter_var,
    iter_level = iter_level,
    title_slide = .ppt_norm_text1(titulo, blank = NULL),
    overrides = overrides,
    base = base
  )
  class(el) <- c("ppt_element", "list")
  el
}

#' @title Radar de dimensiones con fallback automático a barras
#' @export
p_dim_radar <- function(
    modo = c("general", "indicadores"),
    objetivo,
    cruce = NULL,
    incluir_total = NULL,
    filtros = list(),
    iter_var = NULL,
    iter_level = NULL,
    titulo = NULL,
    overrides = list(),
    base = list()
) {
  modo <- match.arg(modo)

  if (!is.character(objetivo) || length(objetivo) != 1L || !nzchar(trimws(objetivo))) {
    stop("`objetivo` debe ser character(1) no vacío.", call. = FALSE)
  }
  objetivo <- trimws(objetivo)

  if (!is.null(cruce)) {
    if (!is.character(cruce) || length(cruce) != 1L || !nzchar(trimws(cruce))) {
      stop("`cruce` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    cruce <- trimws(cruce)
  }

  if (!is.null(iter_var)) {
    if (!is.character(iter_var) || length(iter_var) != 1L || !nzchar(trimws(iter_var))) {
      stop("`iter_var` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    iter_var <- trimws(iter_var)
  }

  if (!is.null(iter_level)) {
    if (!is.character(iter_level) || length(iter_level) != 1L || !nzchar(trimws(iter_level))) {
      stop("`iter_level` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    iter_level <- trimws(iter_level)
  }

  if (!is.null(incluir_total)) {
    if (!is.logical(incluir_total) || length(incluir_total) != 1L || is.na(incluir_total)) {
      stop("`incluir_total` debe ser NULL o logical(1).", call. = FALSE)
    }
  }

  filtros <- .ppt_norm_filters(filtros)
  if (!is.list(overrides)) stop("`overrides` debe ser lista.", call. = FALSE)
  if (!is.list(base)) stop("`base` debe ser lista.", call. = FALSE)

  el <- list(
    .element_type = "dim_radar",
    modo = modo,
    objetivo = objetivo,
    cruce = cruce,
    incluir_total = incluir_total,
    filtros = filtros,
    iter_var = iter_var,
    iter_level = iter_level,
    title_slide = .ppt_norm_text1(titulo, blank = NULL),
    overrides = overrides,
    base = base
  )
  class(el) <- c("ppt_element", "list")
  el
}

#' @title Radar + tabla de dimensiones
#' @export
p_dim_radar_tabla <- function(
    modo = c("general", "indicadores"),
    objetivo,
    cruce = NULL,
    incluir_total = NULL,
    filtros = list(),
    iter_var = NULL,
    iter_level = NULL,
    titulo = NULL,
    titulo_tabla = NULL,
    overrides = list(),
    base = list()
) {
  modo <- match.arg(modo)

  if (!is.character(objetivo) || length(objetivo) != 1L || !nzchar(trimws(objetivo))) {
    stop("`objetivo` debe ser character(1) no vacío.", call. = FALSE)
  }
  objetivo <- trimws(objetivo)

  if (!is.null(cruce)) {
    if (!is.character(cruce) || length(cruce) != 1L || !nzchar(trimws(cruce))) {
      stop("`cruce` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    cruce <- trimws(cruce)
  }

  if (!is.null(iter_var)) {
    if (!is.character(iter_var) || length(iter_var) != 1L || !nzchar(trimws(iter_var))) {
      stop("`iter_var` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    iter_var <- trimws(iter_var)
  }

  if (!is.null(iter_level)) {
    if (!is.character(iter_level) || length(iter_level) != 1L || !nzchar(trimws(iter_level))) {
      stop("`iter_level` debe ser NULL o character(1) no vacío.", call. = FALSE)
    }
    iter_level <- trimws(iter_level)
  }

  if (!is.null(incluir_total)) {
    if (!is.logical(incluir_total) || length(incluir_total) != 1L || is.na(incluir_total)) {
      stop("`incluir_total` debe ser NULL o logical(1).", call. = FALSE)
    }
  }

  if (!is.null(titulo_tabla)) {
    titulo_tabla <- .ppt_norm_text1(titulo_tabla, blank = NULL)
  }

  filtros <- .ppt_norm_filters(filtros)
  if (!is.list(overrides)) stop("`overrides` debe ser lista.", call. = FALSE)
  if (!is.list(base)) stop("`base` debe ser lista.", call. = FALSE)

  el <- list(
    .element_type = "dim_radar_tabla",
    modo = modo,
    objetivo = objetivo,
    cruce = cruce,
    incluir_total = incluir_total,
    filtros = filtros,
    iter_var = iter_var,
    iter_level = iter_level,
    titulo_tabla = titulo_tabla,
    title_slide = .ppt_norm_text1(titulo, blank = NULL),
    overrides = overrides,
    base = base
  )
  class(el) <- c("ppt_element", "list")
  el
}

#' @title Texto (para cajas libres en layouts)
#' @export
p_text <- function(text, overrides = list()) {
  if (missing(text) || is.null(text)) stop("`text` no puede ser NULL.", call. = FALSE)
  if (!is.character(text) || length(text) != 1L) stop("`text` debe ser character(1).", call. = FALSE)

  text <- .ppt_norm_text1(text, blank = " ")

  if (!is.list(overrides)) stop("`overrides` debe ser lista.", call. = FALSE)

  el <- list(
    .element_type = "text",
    text          = text,
    overrides     = overrides
  )
  class(el) <- c("ppt_element_text", "ppt_element", "list")
  el
}
