#' Graficar barras horizontales agrupadas a partir de una tabla ancha
#'
#' Genera un gráfico de barras horizontales agrupadas donde cada categoría
#' (por ejemplo, servicios, enfermedades, regiones) se muestra en el eje
#' vertical, y para cada una se comparan dos o más porcentajes en el eje
#' horizontal (por ejemplo, "% capacitado en el área" y "% capacitado en M&R").
#'
#' La función está pensada para usarse directamente con tablas de indicadores
#' ya calculadas (formato ancho), sin necesidad de reestructurar los datos
#' manualmente.
#'
#' @param data Data frame en formato ancho que contiene la categoría, la base N
#'   y las columnas de porcentajes.
#' @param var_categoria Nombre (string) de la columna que define cada categoría
#'   en el eje vertical (por ejemplo, `"servicio"` o `"condicion"`).
#' @param var_n Nombre (string) de la columna con el N total por categoría
#'   (se usa para mostrar una etiqueta adicional, típicamente `"N="`, al
#'   extremo derecho de las barras).
#' @param cols_porcentaje Vector de nombres de columnas con los porcentajes
#'   que se desean comparar como series agrupadas (por ejemplo,
#'   `c("pct_area", "pct_mr")`).
#' @param etiquetas_series Vector de caracteres con nombre (named) que asigna
#'   etiquetas legibles a las columnas de \code{cols_porcentaje}. Los nombres
#'   deben coincidir exactamente con los nombres de \code{cols_porcentaje}.
#'   Los valores (etiquetas) se usan en la leyenda.
#' @param escala_valor Escala de los porcentajes: \code{"proporcion_1"} si los
#'   valores están en 0–1, o \code{"proporcion_100"} si están en 0–100.
#' @param colores_series Vector de colores HEX con nombre (opcional) para las
#'   series, usando como nombres las etiquetas de \code{etiquetas_series}. Si
#'   se omite, \pkg{ggplot2} asignará colores por defecto.
#' @param mostrar_valores Lógico; si \code{TRUE}, muestra los porcentajes
#'   asociados a cada barra.
#' @param decimales Número de decimales para los porcentajes.
#' @param umbral_etiqueta Mínima proporción (en escala 0–1) para mostrar texto
#'   (por ejemplo, 0.03 equivale a 3\%). Por debajo de este valor, no se
#'   muestra la etiqueta.
#' @param umbral_posicion Proporción (0–1) que define desde qué tamaño de barra
#'   el porcentaje se dibuja centrado dentro de la barra. Si la barra es
#'   menor a este umbral, el porcentaje se dibuja ligeramente a la derecha del
#'   final de la barra.
#' @param sufijo_etiqueta String opcional que se añade al final de cada
#'   etiqueta de porcentaje (por ejemplo, `" pts"`). Se aplica solo a las
#'   etiquetas visibles.
#' @param mostrar_barra_extra Lógico; si \code{TRUE}, agrega una "barra extra"
#'   de texto al extremo derecho de cada categoría (típicamente el N, por
#'   ejemplo `"N=305"`).
#' @param prefijo_barra_extra Texto añadido antes del valor mostrado en la
#'   barra extra (por defecto \code{"N="}).
#' @param titulo_barra_extra Texto que se coloca como encabezado encima de la
#'   columna de barra extra (por ejemplo, `"Total"`). Si es \code{NULL}, no se
#'   agrega encabezado.
#' @param titulo Título del gráfico.
#' @param subtitulo Subtítulo opcional del gráfico.
#' @param nota_pie Nota o fuente para el pie de página.
#' @param nota_pie_derecha Texto adicional para el pie de página que se
#'   concatenará a la derecha de \code{nota_pie}.
#' @param pos_titulo Alineación horizontal del título: `"centro"`, `"izquierda"`
#'   o `"derecha"`.
#' @param pos_nota_pie Alineación horizontal de la nota al pie (caption):
#'   `"derecha"`, `"izquierda"` o `"centro"`.
#' @param color_titulo,size_titulo,color_subtitulo,size_subtitulo,
#'   color_nota_pie,size_nota_pie,color_leyenda,size_leyenda,
#'   color_texto_barras,color_texto_barras_fuera,size_texto_barras,
#'   color_barra_extra,size_barra_extra,color_ejes,size_ejes,
#'   color_fondo Parámetros estéticos.
#' @param grosor_barras Grosor relativo de cada barra dentro de su grupo.
#' @param extra_derecha_rel Porción adicional a la derecha, en unidades de
#'   proporción, para la barra extra.
#' @param espacio_izquierda_rel Proporción del rango del eje de porcentajes
#'   que se deja como espacio en blanco en el extremo izquierdo. Para
#'   \code{escala_valor} proporcional, el eje se fija en 0 sin expansión extra.
#' @param ancho_max_eje_y Número aproximado de caracteres por línea para las
#'   etiquetas de las categorías (eje Y tras el \code{coord_flip}). Si no es
#'   \code{NULL}, se insertan saltos de línea automáticos entre palabras usando
#'   \pkg{stringr}.
#' @param mostrar_leyenda,invertir_leyenda,invertir_barras,invertir_series
#'   Controles de leyenda y orden.
#' @param orientacion Dirección de las barras: `"horizontal"` (por defecto,
#'   categorías en el eje vertical y porcentajes en el eje horizontal mediante
#'   \code{coord_flip()}) o `"vertical"` (categorías en el eje horizontal y
#'   porcentajes en el eje vertical). No afecta el resto de argumentos ni el
#'   cálculo de porcentajes, solo la orientación del gráfico y la lógica de
#'   altura sugerida.
#' @param textos_negrita Vector que puede incluir `"titulo"`, `"porcentajes"`,
#'   `"leyenda"` y/o `"barra_extra"`.
#' @param exportar `"rplot"`, `"png"`, `"ppt"` o `"word"`.
#' @param path_salida Ruta del archivo cuando `exportar != "rplot"`.
#' @param ancho,alto Parámetros de tamaño del gráfico.
#' @param alto_por_categoria Alto (en pulgadas) a asignar por categoría en el
#'   área de barras cuando no se especifica \code{alto}. Sobre ese alto se suma
#'   automáticamente un bloque fijo para la leyenda (si \code{mostrar_leyenda})
#'   y otro más pequeño para el caption (si existe), y luego se acota entre un
#'   mínimo y un máximo razonables para mantener la consistencia visual.
#' @param dpi Resolución para PNG.
#'
#' @return Un objeto \code{ggplot} si \code{exportar = "rplot"}.
#' @export
graficar_barras_agrupadas <- function(
    data,
    var_categoria,
    var_n,
    cols_porcentaje,
    etiquetas_series,
    escala_valor              = c("proporcion_1", "proporcion_100"),
    orientacion               = c("horizontal", "vertical"),
    colores_series            = NULL,
    mostrar_valores           = TRUE,
    decimales                 = 1,
    umbral_etiqueta           = 0.03,
    umbral_posicion           = 0.15,
    sufijo_etiqueta           = "",
    mostrar_barra_extra       = TRUE,
    prefijo_barra_extra       = "N=",
    titulo_barra_extra        = NULL,
    titulo                    = NULL,
    subtitulo                 = NULL,
    nota_pie                  = NULL,
    nota_pie_derecha          = NULL,
    pos_titulo                = c("centro", "izquierda", "derecha"),
    pos_nota_pie              = c("derecha", "izquierda", "centro"),
    # Estilo de texto y layout
    color_titulo              = "#004B8D",
    size_titulo               = 11,
    color_subtitulo           = "#004B8D",
    size_subtitulo            = 9,
    color_nota_pie            = "#004B8D",
    size_nota_pie             = 8,
    color_leyenda             = "#004B8D",
    size_leyenda              = 8,
    color_texto_barras        = "white",
    color_texto_barras_fuera  = "#004B8D",
    size_texto_barras         = 3,
    color_barra_extra         = "#004B8D",
    size_barra_extra          = 3,
    color_ejes                = "#004B8D",
    size_ejes                 = 9,
    color_fondo               = NA,
    grosor_barras             = 0.6,
    extra_derecha_rel         = 0.25,
    espacio_izquierda_rel     = 0.05,
    ancho_max_eje_y           = NULL,
    mostrar_leyenda           = TRUE,
    invertir_leyenda          = FALSE,
    invertir_barras           = FALSE,
    invertir_series           = FALSE,
    textos_negrita            = NULL,
    exportar                  = c("rplot", "png", "ppt", "word"),
    path_salida               = NULL,
    ancho                     = 10,
    alto                      = 6,
    alto_por_categoria        = NULL,
    dpi                       = 300
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y

  escala_valor <- match.arg(escala_valor)
  orientacion  <- match.arg(orientacion)
  exportar     <- match.arg(exportar)
  pos_titulo   <- match.arg(pos_titulo)
  pos_nota_pie <- match.arg(pos_nota_pie)

  decimales <- suppressWarnings(as.numeric(decimales))
  if (length(decimales) < 1L || !is.finite(decimales[1])) {
    decimales <- 1
  } else {
    decimales <- decimales[1]
  }

  hjust_from_pos <- function(x) {
    switch(
      x,
      "izquierda" = 0,
      "centro"    = 0.5,
      "derecha"   = 1,
      0.5
    )
  }

  hjust_titulo  <- hjust_from_pos(pos_titulo)
  hjust_caption <- hjust_from_pos(pos_nota_pie)

  # ---------------------------------------------------------------------------
  # Validaciones básicas
  # ---------------------------------------------------------------------------
  if (!var_categoria %in% names(data)) {
    stop("`var_categoria` no existe en `data`.", call. = FALSE)
  }
  if (!var_n %in% names(data)) {
    stop("`var_n` no existe en `data`.", call. = FALSE)
  }
  if (!all(cols_porcentaje %in% names(data))) {
    faltan <- cols_porcentaje[!cols_porcentaje %in% names(data)]
    stop(
      "Las siguientes columnas de `cols_porcentaje` no existen en `data`: ",
      paste(faltan, collapse = ", "),
      call. = FALSE
    )
  }
  if (!all(names(etiquetas_series) %in% cols_porcentaje)) {
    stop(
      "Los nombres de `etiquetas_series` deben coincidir con columnas de `cols_porcentaje`.",
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
      dplyr::all_of(c(var_categoria, var_n, cols_porcentaje))
    ) |>
    tidyr::pivot_longer(
      cols      = dplyr::all_of(cols_porcentaje),
      names_to  = ".col_pct",
      values_to = ".valor"
    )

  if (!is.numeric(df_long$.valor)) {
    stop("Las columnas de porcentaje deben ser numéricas.", call. = FALSE)
  }

  df_long$.serie <- dplyr::recode(df_long$.col_pct, !!!etiquetas_series)

  # Escala interna 0–1
  if (escala_valor == "proporcion_100") {
    df_long$.valor_plot <- df_long$.valor / 100
  } else {
    df_long$.valor_plot <- df_long$.valor
  }

  # Mantener ceros; SOLO NA → 0
  df_long$.valor_plot[is.na(df_long$.valor_plot)] <- 0

  # Orden de series
  niveles_series <- unname(etiquetas_series)
  if (invertir_series) {
    niveles_series <- rev(niveles_series)
  }
  df_long$.serie <- factor(df_long$.serie, levels = niveles_series)

  # Orden de categorías
  cat_vec  <- df_long[[var_categoria]]
  cat_lvls <- unique(cat_vec)
  if (invertir_barras) {
    cat_lvls <- rev(cat_lvls)
  }
  df_long[[var_categoria]] <- factor(cat_vec, levels = cat_lvls)

  # Tamaño de texto según nº de series
  n_series <- length(unique(df_long$.serie))
  size_texto_barras_eff <- dplyr::case_when(
    n_series <= 2 ~ size_texto_barras * 1.00,
    n_series == 3 ~ size_texto_barras * 0.85,
    n_series == 4 ~ size_texto_barras * 0.70,
    TRUE          ~ size_texto_barras * 0.55
  )

  max_valor <- max(df_long$.valor_plot, na.rm = TRUE)

  # ---------------------------------------------------------------------------
  # 2. Gráfico base
  # ---------------------------------------------------------------------------
  width_dodge <- 0.7

  p <- ggplot2::ggplot(
    df_long,
    ggplot2::aes_string(
      x    = var_categoria,
      y    = ".valor_plot",
      fill = ".serie"
    )
  ) +
    ggplot2::geom_col(
      position = ggplot2::position_dodge(width = width_dodge),
      width    = grosor_barras
    )

  # ---------------------------------------------------------------------------
  # 3. Etiquetas de porcentaje
  # ---------------------------------------------------------------------------
  if (mostrar_valores) {

    df_lab <- df_long

    # Construir etiqueta numérica: enteros sin decimales, resto con `decimales`
    pct_num <- df_lab$.valor_plot * 100
    tol <- 10^(-(decimales + 1))
    es_entero <- is.finite(pct_num) & (abs(pct_num - round(pct_num)) < tol)

    lab_base <- character(nrow(df_lab))
    fmt_no_entero <- paste0("%.", decimales, "f%%")

    lab_base[es_entero]  <- sprintf("%d%%", round(pct_num[es_entero]))
    lab_base[!es_entero] <- sprintf(fmt_no_entero, pct_num[!es_entero])

    # Ocultar 0% y valores bajo umbral
    lab_base[df_lab$.valor_plot <= 0] <- NA_character_
    lab_base[df_lab$.valor_plot < umbral_etiqueta] <- NA_character_

    # Añadir sufijo solo a etiquetas visibles
    df_lab$lab <- ifelse(
      !is.na(lab_base),
      paste0(lab_base, sufijo_etiqueta),
      ""
    )

    # Umbral para decidir dentro / fuera
    umbral_posicion_eff <- umbral_posicion
    if (!is.finite(umbral_posicion_eff) || umbral_posicion_eff <= 0) {
      umbral_posicion_eff <- 0.15
    }

    # Offset para etiquetas fuera de la barra:
    # en vertical conviene un poco más de espacio
    offset_lab <- if (orientacion == "vertical") {
      max_valor * 0.03
    } else {
      max_valor * 0.015
    }

    df_lab$inside <- df_lab$.valor_plot >= umbral_posicion_eff & df_lab$lab != ""

    df_lab$valor_label <- df_lab$.valor_plot
    df_lab$valor_label[df_lab$inside] <-
      df_lab$.valor_plot[df_lab$inside] / 2
    df_lab$valor_label[!df_lab$inside & df_lab$.valor_plot > 0] <-
      df_lab$.valor_plot[!df_lab$inside & df_lab$.valor_plot > 0] + offset_lab

    # En horizontal: dentro centrado, fuera a la derecha.
    # En vertical: SIEMPRE centrado sobre la barra.
    df_lab$hjust_label <- ifelse(df_lab$inside, 0.5, 0)
    if (orientacion == "vertical") {
      df_lab$hjust_label <- 0.5
    }
    df_lab$col_label <- ifelse(
      df_lab$inside,
      color_texto_barras,
      color_texto_barras_fuera
    )

    p <- p +
      ggplot2::geom_text(
        data        = df_lab[df_lab$lab != "", , drop = FALSE],
        mapping     = ggplot2::aes(
          x       = .data[[var_categoria]],
          y       = valor_label,
          label   = lab,
          group   = .serie,
          colour  = col_label,
          hjust   = hjust_label
        ),
        inherit.aes = FALSE,
        position    = ggplot2::position_dodge(width = width_dodge),
        vjust       = 0.5,
        size        = size_texto_barras_eff,
        fontface    = if ("porcentajes" %in% textos_negrita) "bold" else "plain",
        show.legend = FALSE
      ) +
      ggplot2::scale_colour_identity(guide = "none")
  }

  # ---------------------------------------------------------------------------
  # 4. Escala Y (proporción + espacio para barra extra)
  # ---------------------------------------------------------------------------
  es_proporcion <- escala_valor %in% c("proporcion_1", "proporcion_100")

  if (es_proporcion) {

    if (orientacion == "horizontal") {
      # Mantener la lógica 0–100% fija para horizontales
      if (mostrar_barra_extra) {
        y_lim   <- 1 + extra_derecha_rel
        y_extra <- 1 + extra_derecha_rel * 0.5
      } else {
        y_lim   <- 1
        y_extra <- NA_real_
      }

      p <- p +
        ggplot2::scale_y_continuous(
          limits = c(0, y_lim),
          breaks = seq(0.25, 1, by = 0.25),
          labels = scales::percent_format(accuracy = 1),
          expand = ggplot2::expansion(mult = c(0, 0.02))
        )

    } else {
      # VERTICAL: escala libre según los datos (pero manteniendo %)
      if (mostrar_barra_extra) {
        y_lim   <- max_valor * (1 + extra_derecha_rel)
        y_extra <- max_valor * (1 + extra_derecha_rel * 0.95)
      } else {
        y_lim   <- max_valor * 1.05
        y_extra <- NA_real_
      }

      p <- p +
        ggplot2::scale_y_continuous(
          limits = c(0, y_lim),
          labels = scales::percent_format(accuracy = 1),
          expand = ggplot2::expansion(mult = c(0, 0.02))
        )
    }

  } else {
    y_lim   <- max_valor * (1 + extra_derecha_rel)
    y_extra <- max_valor * (1 + extra_derecha_rel * 0.95)

    p <- p +
      ggplot2::scale_y_continuous(
        limits = c(0, y_lim),
        expand = ggplot2::expansion(mult = c(espacio_izquierda_rel, 0.05))
      )
  }

  # ---------------------------------------------------------------------------
  # 5. Barra extra con N a la derecha
  # ---------------------------------------------------------------------------
  if (mostrar_barra_extra) {
    df_extra <- df |>
      dplyr::select(dplyr::all_of(c(var_categoria, var_n))) |>
      dplyr::distinct() |>
      dplyr::filter(.data[[var_categoria]] %in% levels(df_long[[var_categoria]])) |>
      dplyr::mutate(
        ypos      = if (!is.na(y_extra)) y_extra else max_valor * (1 + extra_derecha_rel * 0.95),
        lab_extra = paste0(prefijo_barra_extra, .data[[var_n]])
      )

    p <- p +
      ggplot2::geom_text(
        data        = df_extra,
        mapping     = ggplot2::aes_string(
          x     = var_categoria,
          y     = "ypos",
          label = "lab_extra"
        ),
        inherit.aes = FALSE,
        hjust       = 0,
        vjust       = 0.5,
        size        = size_barra_extra,
        color       = color_barra_extra,
        fontface    = if ("barra_extra" %in% textos_negrita) "bold" else "plain"
      )

    if (!is.null(titulo_barra_extra) && nzchar(titulo_barra_extra)) {
      lvls <- levels(df_long[[var_categoria]])
      cat_superior <- if (invertir_barras) tail(lvls, 1) else head(lvls, 1)

      df_header <- df_extra[df_extra[[var_categoria]] == cat_superior, , drop = FALSE]

      if (nrow(df_header) == 1L) {
        p <- p +
          ggplot2::geom_text(
            data        = df_header,
            mapping     = ggplot2::aes_string(
              x = var_categoria,
              y = "ypos"
            ),
            label       = titulo_barra_extra,
            inherit.aes = FALSE,
            hjust       = 0,
            vjust       = -1.2,
            size        = size_barra_extra,
            color       = color_barra_extra,
            fontface    = if ("barra_extra" %in% textos_negrita) "bold" else "plain"
          )
      }
    }
  }

  # ---------------------------------------------------------------------------
  # 6. Colores, orientación, tema y wrap de eje de categorías
  # ---------------------------------------------------------------------------
  if (!is.null(colores_series)) {
    p <- p +
      ggplot2::scale_fill_manual(values = colores_series)
  }

  if (!is.null(ancho_max_eje_y)) {
    if (!requireNamespace("stringr", quietly = TRUE)) {
      stop("Para usar `ancho_max_eje_y` se requiere el paquete 'stringr'.",
           call. = FALSE)
    }
    # Las categorías están en el eje X en los datos; en horizontal se hace flip,
    # pero el wrap se aplica igual sobre los niveles de la categoría.
    p <- p +
      ggplot2::scale_x_discrete(
        labels = function(x) stringr::str_wrap(x, width = ancho_max_eje_y)
      )
  }

  caption_text <- NULL
  if (!is.null(nota_pie) && nzchar(nota_pie) &&
      !is.null(nota_pie_derecha) && nzchar(nota_pie_derecha)) {
    caption_text <- paste0(nota_pie, "   ", nota_pie_derecha)
  } else if (!is.null(nota_pie) && nzchar(nota_pie)) {
    caption_text <- nota_pie
  } else if (!is.null(nota_pie_derecha) && nzchar(nota_pie_derecha)) {
    caption_text <- nota_pie_derecha
  }

  base_theme <- ggplot2::theme_minimal(base_size = 9) +
    ggplot2::theme(
      panel.grid.major.y = ggplot2::element_blank(),
      panel.grid.minor   = ggplot2::element_blank(),
      panel.grid.major.x = ggplot2::element_blank(),
      axis.title.x       = ggplot2::element_blank(),
      axis.title.y       = ggplot2::element_blank(),
      legend.title       = ggplot2::element_blank(),
      legend.position    = if (mostrar_leyenda) "bottom" else "none",
      legend.text        = ggplot2::element_text(
        color = color_leyenda,
        size  = size_leyenda,
        face  = if ("leyenda" %in% textos_negrita) "bold" else "plain"
      ),
      # margen condicionado a la orientación
      plot.margin        = if (orientacion == "horizontal") {
        ggplot2::margin(t = 15, r = 80, b = 15, l = 5)
      } else {
        ggplot2::margin(t = 15, r = 10, b = 15, l = 10)
      },
      plot.title         = ggplot2::element_text(
        hjust = hjust_titulo,
        color = color_titulo,
        size  = size_titulo,
        face  = if ("titulo" %in% textos_negrita) "bold" else "plain"
      ),
      plot.subtitle      = ggplot2::element_text(
        hjust = hjust_titulo,
        color = color_subtitulo,
        size  = size_subtitulo
      ),
      plot.caption       = ggplot2::element_text(
        hjust = hjust_caption,
        color = color_nota_pie,
        size  = size_nota_pie
      ),
      plot.background    = ggplot2::element_rect(fill = color_fondo, color = NA),
      panel.background   = ggplot2::element_rect(fill = color_fondo, color = NA)
    )

  # Eje de porcentajes (gris): depende de la orientación
  if (escala_valor %in% c("proporcion_1", "proporcion_100")) {

    if (orientacion == "horizontal") {
      # En horizontal, tras el coord_flip, los porcentajes terminan en el eje X.
      eje_theme <- ggplot2::theme(
        axis.text.x  = ggplot2::element_text(
          color = "#7F7F7F",
          size  = size_ejes
        ),
        axis.ticks.x = ggplot2::element_line(
          color     = "#7F7F7F",
          linewidth = 0.3
        ),
        axis.line.x  = ggplot2::element_line(
          color     = "#7F7F7F",
          linewidth = 0.4
        )
      )
    } else {
      # En vertical, los porcentajes están en el eje Y.
      eje_theme <- ggplot2::theme(
        axis.text.y  = ggplot2::element_text(
          color = "#7F7F7F",
          size  = size_ejes
        ),
        axis.ticks.y = ggplot2::element_line(
          color     = "#7F7F7F",
          linewidth = 0.3
        ),
        axis.line.y  = ggplot2::element_line(
          color     = "#7F7F7F",
          linewidth = 0.4
        )
      )
    }

  } else {
    eje_theme <- ggplot2::theme()
  }

  # Eje de categorías (color_ejes): también depende de la orientación
  theme_orient <- if (orientacion == "horizontal") {
    ggplot2::theme(
      axis.text.y = ggplot2::element_text(
        color = color_ejes,
        size  = size_ejes,
        hjust = 1,
        vjust = 0.5
      ),
      axis.text.x  = ggplot2::element_blank(),
      axis.ticks.x = ggplot2::element_blank(),
      axis.line.x  = ggplot2::element_blank()
    )
  } else {
    ggplot2::theme(
      axis.text.x = ggplot2::element_text(
        color = color_ejes,
        size  = size_ejes,
        hjust = 0.5,
        vjust = 1
      )
    )
  }

  # Número de ítems en la leyenda y filas necesarias (máx. 5 por fila)
  n_items_leyenda <- length(levels(df_long$.serie))
  n_por_fila      <- 5L
  n_filas_leyenda <- max(1L, ceiling(n_items_leyenda / n_por_fila))

  if (mostrar_leyenda) {
    p <- p +
      ggplot2::guides(
        fill = ggplot2::guide_legend(
          nrow    = n_filas_leyenda,
          reverse = invertir_leyenda
        )
      )
  }

  # Aplicar orientación (coord_flip solo en horizontal)
  if (orientacion == "horizontal") {
    p <- p + ggplot2::coord_flip()
  }

  p <- p +
    base_theme +
    eje_theme +
    theme_orient +
    ggplot2::labs(
      title    = titulo,
      subtitle = subtitulo,
      caption  = caption_text
    )

  # ---------------------------------------------------------------------------
  # 7. Exportación (altura total sugerida)
  # ---------------------------------------------------------------------------

  n_categorias <- length(unique(df_long[[var_categoria]]))

  # Parámetros de descomposición de altura
  alto_max_total    <- 9.0   # tope máximo en pulgadas
  alto_leyenda_row  <- 0.35  # alto aprox. por fila de leyenda
  alto_caption      <- 0.25  # bloque adicional si hay caption
  alto_por_cat_eff  <- alto_por_categoria %||% 0.35

  # Leyenda: cuántos ítems y cuántas filas
  n_items_leyenda <- length(levels(df_long$.serie))
  if (!exists("n_filas_leyenda")) {
    # Por defecto: máximo 5 ítems por fila
    n_filas_leyenda <- if (n_items_leyenda <= 0) 0L else ceiling(n_items_leyenda / 5)
  }

  tiene_caption <- !is.null(caption_text) && nzchar(caption_text)

  if (orientacion == "horizontal") {

    # =================== BARRAS HORIZONTALES ==================

    # 7.1 Alto del PANEL de barras (una barra por categoría)
    alto_panel <- max(n_categorias, 1L) * alto_por_cat_eff

    # 7.2 Alto de la LEYENDA
    alto_leyenda <- if (mostrar_leyenda && n_items_leyenda > 0) {
      n_filas_leyenda * alto_leyenda_row
    } else {
      0
    }

    # 7.3 Alto del CAPTION interno (nota_pie)
    alto_cap <- if (tiene_caption) alto_caption else 0

    # 7.4 Alto extra por EJE de categorías (cuando las etiquetas tienen varias líneas)
    max_lineas_eje <- 1L
    if (!is.null(ancho_max_eje_y)) {
      if (requireNamespace("stringr", quietly = TRUE)) {
        # Tomamos las etiquetas tal como se mostrarán en el eje de categorías
        etiq_orig <- levels(df_long[[var_categoria]])
        if (length(etiq_orig) == 0L) {
          etiq_orig <- unique(as.character(df_long[[var_categoria]]))
        }
        etiq_wrap <- stringr::str_wrap(etiq_orig, width = ancho_max_eje_y)
        lineas    <- stringr::str_count(etiq_wrap, "\n") + 1L
        max_lineas_eje <- max(1L, lineas, na.rm = TRUE)
      }
    }

    # Cada línea adicional del eje Y suma un pequeño bloque de altura
    alto_por_linea_eje <- 0.12
    alto_extra_eje_y <- if (max_lineas_eje > 1L) {
      (max_lineas_eje - 1L) * alto_por_linea_eje
    } else {
      0
    }

    # 7.5 Altura total sugerida (panel + eje Y + leyenda + caption)
    alto_total_sugerido <- alto_panel + alto_leyenda + alto_cap + alto_extra_eje_y

    # Tope máximo
    alto_total_sugerido <- min(alto_max_total, alto_total_sugerido)

    # Mínimo "suave" dependiente del alto por categoría
    alto_min_suave <- (alto_por_cat_eff * 1.2) + alto_leyenda + alto_cap
    alto_total_sugerido <- max(alto_min_suave, alto_total_sugerido)

  } else {

    # ==================== BARRAS VERTICALES ======================

    # Aquí el alto ya no depende de cuántas categorías haya en el eje X.
    # Se usa un panel base fijo, ajustado por leyenda y caption.
    alto_panel_base <- if (!is.null(alto_por_categoria)) {
      max(3.5, alto_por_categoria * 5)
    } else {
      4.5
    }

    alto_leyenda <- if (mostrar_leyenda && n_items_leyenda > 0) {
      n_filas_leyenda * alto_leyenda_row
    } else {
      0
    }

    alto_cap <- if (tiene_caption) alto_caption else 0

    alto_total_sugerido <- alto_panel_base + alto_leyenda + alto_cap

    alto_total_sugerido <- min(alto_max_total, alto_total_sugerido)

    # Mínimo razonable para que no quede enano
    alto_min_vertical <- 3.5
    alto_total_sugerido <- max(alto_min_vertical, alto_total_sugerido)
  }

  # Si solo queremos el ggplot, devolvemos p con el alto sugerido como atributo
  if (exportar == "rplot") {
    attr(p, "alto_word_sugerido") <- alto_total_sugerido
    return(p)
  }

  if (is.null(path_salida) || !nzchar(path_salida)) {
    stop("Debe especificar `path_salida` cuando `exportar` no es 'rplot'.", call. = FALSE)
  }

  # Altura efectiva para PNG / PPT / WORD
  height_plot <- if (!missing(alto) && !is.null(alto)) {
    alto
  } else {
    alto_total_sugerido
  }

  # ---------------------- WORD ----------------------
  if (exportar == "word") {
    if (!requireNamespace("officer", quietly = TRUE)) {
      stop(
        "Para exportar a Word se requiere el paquete 'officer'.",
        call. = FALSE
      )
    }

    width_word <- if (!missing(ancho) && !is.null(ancho)) ancho else 6.5

    doc <- officer::read_docx()
    doc <- officer::body_add_gg(
      doc,
      value  = p,
      width  = width_word,
      height = height_plot,
      style  = "centered"
    )
    print(doc, target = path_salida)

    return(invisible(p))
  }

  # ---------------------- PNG ----------------------
  if (exportar == "png") {
    ggplot2::ggsave(
      filename = path_salida,
      plot     = p,
      width    = ancho,
      height   = height_plot,
      dpi      = dpi,
      bg       = if (is.na(color_fondo)) "transparent" else color_fondo
    )
    return(invisible(p))
  }

  # ---------------------- PPT ----------------------
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
