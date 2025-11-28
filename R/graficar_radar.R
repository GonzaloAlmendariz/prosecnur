# =============================================================================
# graficar_radar()
# -----------------------------------------------------------------------------
# Gráfico tipo radar (telaraña) para comparar 2+ series en varias dimensiones.
# =============================================================================

#' Graficar un radar (telaraña) para varias dimensiones e indicadores
#'
#' Genera un gráfico tipo radar (telaraña) donde cada dimensión
#' (por ejemplo, "Calidad de atención", "Accesibilidad") se representa como
#' un eje alrededor de una figura poligonal, y para cada una se comparan dos
#' o más series (por ejemplo, "Calidad percibida", "Acceso", "Satisfacción").
#'
#' Los datos se esperan en formato ancho: una columna categórica que define
#' las dimensiones y varias columnas numéricas con porcentajes o puntajes
#' a comparar.
#'
#' @param data Data frame en formato ancho que contiene la columna de
#'   dimensiones y las columnas de valores.
#' @param var_categoria Nombre (string) de la columna que define la dimensión
#'   (por ejemplo, `"dominio"` o `"dimension"`).
#' @param cols_valor Vector de nombres de columnas numéricas con los valores
#'   a comparar (por ejemplo, `c("pct_calidad","pct_acceso","pct_satisf")`).
#' @param etiquetas_series Vector de caracteres con nombre (named) que asigna
#'   etiquetas legibles a las columnas de \code{cols_valor}. Los nombres deben
#'   coincidir exactamente con \code{cols_valor}; los valores se usan en la
#'   leyenda.
#' @param escala_valor Escala de los valores: \code{"proporcion_1"} si están
#'   en 0–1, o \code{"proporcion_100"} si están en 0–100.
#' @param colores_series Vector de colores HEX con nombre (opcional) para las
#'   series, usando como nombres las etiquetas de \code{etiquetas_series}.
#'   Si se omite, \pkg{ggplot2} asignará colores por defecto.
#'
#' @param breaks_radiales Vector numérico con los valores radiales (0–1) para
#'   las circunferencias de referencia (rejilla).
#' @param mostrar_radios Lógico; si \code{TRUE}, muestra las líneas rectas que
#'   van del centro hacia las dimensiones (ejes radiales).
#' @param mostrar_rejilla Lógico; si \code{TRUE}, muestra las circunferencias
#'   (poligonales) de referencia en los \code{breaks_radiales}.
#' @param mostrar_labels_radiales Lógico; si \code{TRUE}, muestra las etiquetas
#'   de texto de los niveles radiales (por ejemplo, "20%", "40%", etc.).
#' @param dist_etiquetas_rel Factor multiplicativo (>1) que controla qué tan
#'   lejos del polígono se ubican las etiquetas de las dimensiones.
#'
#' @param alpha_relleno Nivel de transparencia (0–1) para el relleno de los
#'   polígonos de cada serie.
#' @param mostrar_puntos Lógico; si \code{TRUE}, dibuja puntos en los vértices.
#' @param mostrar_valores Lógico; si \code{TRUE}, muestra el valor (porcentaje)
#'   cerca de cada vértice.
#' @param decimales Número de decimales para los valores mostrados cuando
#'   \code{mostrar_valores = TRUE}.
#' @param color_texto_valores Color del texto de los valores.
#' @param size_texto_valores Tamaño del texto de los valores.
#'
#' @param titulo Título del gráfico.
#' @param subtitulo Subtítulo opcional.
#' @param nota_pie Nota o fuente para el pie de página (alineado a la derecha).
#'
#' @param color_titulo Color del título.
#' @param size_titulo Tamaño del título.
#' @param color_subtitulo Color del subtítulo.
#' @param size_subtitulo Tamaño del subtítulo.
#' @param color_nota_pie Color del texto del pie de página.
#' @param size_nota_pie Tamaño del texto del pie de página.
#' @param color_leyenda Color del texto de la leyenda.
#' @param size_leyenda Tamaño del texto de la leyenda.
#' @param color_texto_breaks Color del texto de los breaks radiales.
#' @param size_texto_breaks Tamaño del texto de los breaks radiales.
#'
#' @param size_dimensiones Tamaño del texto de cada dimensión alrededor
#'   del radar.
#' @param color_dimensiones Color del texto de cada dimensión.
#'
#' @param mostrar_leyenda Lógico; si \code{FALSE}, oculta la leyenda.
#' @param invertir_leyenda Lógico; si \code{TRUE}, invierte el orden de los
#'   ítems de la leyenda.
#'
#' @param textos_negrita Vector de caracteres que indica qué elementos deben
#'   mostrarse en negrita. Puede incluir cualquiera de:
#'   \code{"titulo"}, \code{"leyenda"}, \code{"valores"}, \code{"dimensiones"}.
#'
#' @param exportar Método de salida: \code{"rplot"} (devuelve un objeto
#'   \code{ggplot}), \code{"png"} (exporta un archivo PNG) o \code{"ppt"}
#'   (exporta una diapositiva PPTX con el gráfico incrustado).
#' @param path_salida Ruta del archivo a crear cuando \code{exportar} no es
#'   \code{"rplot"}.
#' @param ancho Ancho del gráfico (cuando se exporta a archivo).
#' @param alto Alto del gráfico (cuando se exporta a archivo).
#' @param dpi Resolución en puntos por pulgada (solo para PNG).
#'
#' @return Un objeto \code{ggplot} si \code{exportar = "rplot"}. De forma
#'   invisible, el gráfico exportado si se utiliza \code{"png"} o \code{"ppt"}.
#' @export
graficar_radar <- function(
    data,
    var_categoria,
    cols_valor,
    etiquetas_series,
    escala_valor            = c("proporcion_1", "proporcion_100"),
    colores_series          = NULL,
    breaks_radiales         = c(0.2, 0.4, 0.6, 0.8, 1),
    mostrar_radios          = TRUE,
    mostrar_rejilla         = TRUE,
    mostrar_labels_radiales = TRUE,
    dist_etiquetas_rel      = 1.12,
    alpha_relleno           = 0.20,
    mostrar_puntos          = TRUE,
    mostrar_valores         = FALSE,
    decimales               = 1,
    color_texto_valores     = "#000000",
    size_texto_valores      = 3,
    titulo                  = NULL,
    subtitulo               = NULL,
    nota_pie                = NULL,
    color_titulo            = "#000000",
    size_titulo             = 11,
    color_subtitulo         = "#000000",
    size_subtitulo          = 9,
    color_nota_pie          = "#000000",
    size_nota_pie           = 8,
    color_leyenda           = "#000000",
    size_leyenda            = 8,
    color_texto_breaks      = "#666666",
    size_texto_breaks       = 3,
    size_dimensiones        = 4,
    color_dimensiones       = "#000000",
    mostrar_leyenda         = TRUE,
    invertir_leyenda        = FALSE,
    textos_negrita          = NULL,
    exportar                = c("rplot", "png", "ppt"),
    path_salida             = NULL,
    ancho                   = 8,
    alto                    = 8,
    dpi                     = 300
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y

  escala_valor <- match.arg(escala_valor)
  exportar     <- match.arg(exportar)

  # ---------------------------------------------------------------------------
  # 0. Validaciones
  # ---------------------------------------------------------------------------
  if (!var_categoria %in% names(data)) {
    stop("`var_categoria` no existe en `data`.", call. = FALSE)
  }
  if (!all(cols_valor %in% names(data))) {
    faltan <- cols_valor[!cols_valor %in% names(data)]
    stop(
      "Las siguientes columnas de `cols_valor` no existen en `data`: ",
      paste(faltan, collapse = ", "),
      call. = FALSE
    )
  }
  if (!all(names(etiquetas_series) %in% cols_valor)) {
    stop(
      "Los nombres de `etiquetas_series` deben coincidir con columnas de `cols_valor`.",
      call. = FALSE
    )
  }

  textos_negrita <- textos_negrita %||% character(0)
  df <- data

  # ---------------------------------------------------------------------------
  # 1. Ancho → largo
  # ---------------------------------------------------------------------------
  df_long <- df |>
    dplyr::select(
      dplyr::all_of(c(var_categoria, cols_valor))
    ) |>
    tidyr::pivot_longer(
      cols      = dplyr::all_of(cols_valor),
      names_to  = ".col_valor",
      values_to = ".valor"
    )

  if (!is.numeric(df_long$.valor)) {
    stop("Las columnas de `cols_valor` deben ser numéricas.", call. = FALSE)
  }

  # Etiquetas legibles de series
  df_long$.serie <- dplyr::recode(df_long$.col_valor, !!!etiquetas_series)

  # Escala interna 0–1
  if (escala_valor == "proporcion_100") {
    df_long$.valor_plot <- df_long$.valor / 100
  } else {
    df_long$.valor_plot <- df_long$.valor
  }

  # Orden de series según etiquetas_series
  df_long$.serie <- factor(df_long$.serie, levels = unname(etiquetas_series))

  # Orden de dimensiones según aparición
  dim_vec  <- df_long[[var_categoria]]
  dim_lvls <- unique(dim_vec)
  df_long[[var_categoria]] <- factor(dim_vec, levels = dim_lvls)

  k <- length(dim_lvls)
  if (k < 3) {
    stop("Se requieren al menos 3 dimensiones para un radar.", call. = FALSE)
  }

  # Ángulos igualmente espaciados (comenzando arriba)
  theta0 <- -pi / 2
  angle_tbl <- tibble::tibble(
    !!var_categoria := factor(dim_lvls, levels = dim_lvls),
    angle           = theta0 + 2 * pi * (seq_len(k) - 1) / k
  )

  # Máximos radiales
  max_break <- if (length(breaks_radiales)) max(breaks_radiales, na.rm = TRUE) else 1
  max_plot  <- max_break * dist_etiquetas_rel

  # ---------------------------------------------------------------------------
  # 2. Convertir a coordenadas (x, y) para cada serie y dimensión
  # ---------------------------------------------------------------------------
  df_xy <- df_long |>
    dplyr::left_join(angle_tbl, by = var_categoria) |>
    dplyr::mutate(
      x = .data$.valor_plot * cos(.data$angle),
      y = .data$.valor_plot * sin(.data$angle)
    )

  # Polígonos cerrados por serie
  df_poly <- df_xy |>
    dplyr::arrange(.data$.serie, .data[[var_categoria]]) |>
    dplyr::group_by(.data$.serie) |>
    dplyr::group_modify(~ {
      d <- .x
      rbind(d, d[1, ])
    }) |>
    dplyr::ungroup()

  # ---------------------------------------------------------------------------
  # 3. Rejilla poligonal (breaks radiales) y radios
  # ---------------------------------------------------------------------------
  grid_df <- NULL
  if (length(breaks_radiales) && any(breaks_radiales > 0) && mostrar_rejilla) {
    grid_df <- lapply(breaks_radiales, function(r) {
      tibble::tibble(
        !!var_categoria := factor(dim_lvls, levels = dim_lvls)
      ) |>
        dplyr::left_join(angle_tbl, by = var_categoria) |>
        dplyr::mutate(
          nivel = r,
          x     = r * cos(.data$angle),
          y     = r * sin(.data$angle)
        ) |>
        dplyr::arrange(.data[[var_categoria]]) |>
        (\(d) rbind(d, d[1, ]))()
    }) |>
      dplyr::bind_rows()
  }

  axes_df <- NULL
  if (mostrar_radios) {
    axes_df <- angle_tbl |>
      dplyr::mutate(
        x0 = 0,
        y0 = 0,
        x1 = max_break * cos(.data$angle),
        y1 = max_break * sin(.data$angle)
      )
  }

  # ---------------------------------------------------------------------------
  # 4. Etiquetas de dimensiones y breaks
  # ---------------------------------------------------------------------------
  df_lab_dim <- angle_tbl |>
    dplyr::mutate(
      x_lab = max_plot * cos(.data$angle),
      y_lab = max_plot * sin(.data$angle),
      label = as.character(.data[[var_categoria]])
    )

  df_breaks <- NULL
  if (length(breaks_radiales) && any(breaks_radiales > 0) && mostrar_labels_radiales) {
    df_breaks <- tibble::tibble(
      r   = breaks_radiales,
      x   = breaks_radiales,
      y   = 0,
      lab = scales::percent(breaks_radiales, accuracy = 1)
    )
  }

  # ---------------------------------------------------------------------------
  # 5. Construir el gráfico
  # ---------------------------------------------------------------------------
  p <- ggplot2::ggplot()

  # Rejilla
  if (!is.null(grid_df)) {
    p <- p +
      ggplot2::geom_path(
        data = grid_df,
        ggplot2::aes(x = .data$x, y = .data$y, group = .data$nivel),
        color = "grey88",
        linewidth = 0.4
      )
  }

  # Radios
  if (!is.null(axes_df)) {
    p <- p +
      ggplot2::geom_segment(
        data = axes_df,
        ggplot2::aes(x = .data$x0, y = .data$y0, xend = .data$x1, yend = .data$y1),
        color = "grey85",
        linewidth = 0.4
      )
  }

  # Polígonos de series
  p <- p +
    ggplot2::geom_polygon(
      data = df_poly,
      ggplot2::aes(
        x     = .data$x,
        y     = .data$y,
        group = .data$.serie,
        fill  = .data$.serie,
        color = .data$.serie
      ),
      alpha     = alpha_relleno,
      linewidth = 0.9
    )

  # Puntos
  if (isTRUE(mostrar_puntos)) {
    p <- p +
      ggplot2::geom_point(
        data = df_xy,
        ggplot2::aes(
          x     = .data$x,
          y     = .data$y,
          color = .data$.serie
        ),
        size = 2
      )
  }

  # Etiquetas de dimensiones
  p <- p +
    ggplot2::geom_text(
      data = df_lab_dim,
      ggplot2::aes(
        x     = .data$x_lab,
        y     = .data$y_lab,
        label = .data$label
      ),
      color      = color_dimensiones,
      size       = size_dimensiones,
      fontface   = if ("dimensiones" %in% textos_negrita) "bold" else "plain",
      lineheight = 0.95
    )

  # Etiquetas de breaks
  if (!is.null(df_breaks)) {
    p <- p +
      ggplot2::geom_text(
        data = df_breaks,
        ggplot2::aes(
          x     = .data$x,
          y     = .data$y,
          label = .data$lab
        ),
        color = color_texto_breaks,
        size  = size_texto_breaks,
        vjust = -0.2
      )
  }

  # Valores en cada vértice
  if (isTRUE(mostrar_valores)) {
    df_vals <- df_xy
    df_vals$lab <- scales::percent(df_vals$.valor_plot, accuracy = 10^(-decimales))

    offset_r <- max_break * 0.05
    df_vals$r_lab <- pmin(df_vals$.valor_plot + offset_r, max_plot * 0.95)
    df_vals$x_lab <- df_vals$r_lab * cos(df_vals$angle)
    df_vals$y_lab <- df_vals$r_lab * sin(df_vals$angle)

    p <- p +
      ggplot2::geom_text(
        data = df_vals,
        ggplot2::aes(
          x     = .data$x_lab,
          y     = .data$y_lab,
          label = .data$lab
        ),
        color      = color_texto_valores,
        size       = size_texto_valores,
        fontface   = if ("valores" %in% textos_negrita) "bold" else "plain",
        lineheight = 0.95
      )
  }

  # Colores de series
  if (!is.null(colores_series)) {
    p <- p +
      ggplot2::scale_fill_manual(values = colores_series) +
      ggplot2::scale_color_manual(values = colores_series)
  }

  # Escala y tema
  lim <- max_plot * 1.05

  p <- p +
    ggplot2::coord_equal(xlim = c(-lim, lim), ylim = c(-lim, lim), expand = FALSE) +
    ggplot2::theme_minimal(base_size = 10) +
    ggplot2::theme(
      panel.grid      = ggplot2::element_blank(),
      axis.text       = ggplot2::element_blank(),
      axis.title      = ggplot2::element_blank(),
      axis.ticks      = ggplot2::element_blank(),
      legend.position = if (mostrar_leyenda) "bottom" else "none",
      legend.title    = ggplot2::element_blank(),
      legend.text     = ggplot2::element_text(
        color = color_leyenda,
        size  = size_leyenda,
        face  = if ("leyenda" %in% textos_negrita) "bold" else "plain"
      ),
      plot.title      = ggplot2::element_text(
        hjust = 0.5,
        color = color_titulo,
        size  = size_titulo,
        face  = if ("titulo" %in% textos_negrita) "bold" else "plain"
      ),
      plot.subtitle   = ggplot2::element_text(
        hjust = 0.5,
        color = color_subtitulo,
        size  = size_subtitulo
      ),
      plot.caption    = ggplot2::element_text(
        hjust = 1,
        color = color_nota_pie,
        size  = size_nota_pie
      ),
      plot.margin     = ggplot2::margin(t = 15, r = 15, b = 20, l = 15)
    ) +
    ggplot2::labs(
      title    = titulo,
      subtitle = subtitulo,
      caption  = nota_pie
    )

  # Leyenda invertida
  if (invertir_leyenda && mostrar_leyenda) {
    p <- p +
      ggplot2::guides(
        color = ggplot2::guide_legend(reverse = TRUE),
        fill  = ggplot2::guide_legend(reverse = TRUE)
      )
  }

  # ---------------------------------------------------------------------------
  # 7. Exportación
  # ---------------------------------------------------------------------------
  if (exportar == "rplot") {
    return(p)
  }

  if (is.null(path_salida) || !nzchar(path_salida)) {
    stop("Debe especificar `path_salida` cuando `exportar` no es 'rplot'.", call. = FALSE)
  }

  if (exportar == "png") {
    ggplot2::ggsave(
      filename = path_salida,
      plot     = p,
      width    = ancho,
      height   = alto,
      dpi      = dpi
    )
    return(invisible(p))
  }

  if (exportar == "ppt") {
    if (!requireNamespace("officer", quietly = TRUE) ||
        !requireNamespace("rvg", quietly = TRUE)) {
      stop(
        "Para exportar a PPT se requieren los paquetes 'officer' y 'rvg'.",
        call. = FALSE
      )
    }

    doc <- officer::read_pptx()
    doc <- officer::add_slide(doc, layout = "Blank", master = "Office Theme")
    doc <- officer::ph_with(
      doc,
      rvg::dml(ggobj = p),
      location = officer::ph_location_fullsize()
    )
    print(doc, target = path_salida)

    return(invisible(p))
  }

  p
}
