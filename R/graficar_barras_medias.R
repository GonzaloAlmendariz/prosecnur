# =============================================================================
# graficar_barras_media_agrupadas()
# -----------------------------------------------------------------------------
# Barras (horizontales o verticales) agrupadas para comparar 2+ medias
# por categoría (por ejemplo, "Promedio total" vs "Promedio consulta").
# =============================================================================

#' Graficar barras agrupadas de medias por categoría
#'
#' Genera un gráfico de barras agrupadas donde cada categoría (por ejemplo,
#' servicios, enfermedades, regiones) se muestra en el eje de categorías y,
#' para cada una, se comparan dos o más medias (por ejemplo, "Promedio total"
#' y "Promedio por consulta"), diferenciadas por color y con leyenda inferior.
#'
#' La función está pensada para usarse directamente con tablas de indicadores
#' ya calculadas (formato ancho), sin necesidad de reestructurar los datos
#' manualmente.
#'
#' @param data Data frame en formato ancho que contiene la categoría, la base N
#'   (opcional) y las columnas de medias.
#' @param var_categoria Nombre (string) de la columna que define cada categoría
#'   en el eje (por ejemplo, `"servicio"` o `"condicion"`).
#' @param var_n Nombre (string) de la columna con el N total por categoría.
#'   Se usa para mostrar una etiqueta adicional al extremo derecho de las barras
#'   (por ejemplo `"N=50"`). Si es \code{NULL}, no se muestra la barra extra.
#' @param vars_media Vector de nombres de columnas con las medias que se desean
#'   comparar como series agrupadas (por ejemplo,
#'   \code{c("promedio_total", "promedio_consulta")}).
#' @param etiquetas_series Vector de caracteres con nombre (named) que asigna
#'   etiquetas legibles a las columnas de \code{vars_media}. Los nombres deben
#'   coincidir exactamente con \code{vars_media}; los valores se usan en la
#'   leyenda.
#'
#' @param orientacion Orientación de las barras: \code{"horizontal"} (por
#'   defecto) o \code{"vertical"}. En horizontal las categorías se muestran en
#'   el eje Y; en vertical, en el eje X.
#' @param formato_valor Formato de las etiquetas de valor:
#'   \code{"numero"} o \code{"moneda"} (por ejemplo, S/).
#' @param decimales Número de decimales a mostrar en las etiquetas.
#' @param simbolo_moneda Símbolo de moneda a anteponer cuando
#'   \code{formato_valor = "moneda"}.
#' @param separador_miles Separador de miles para el formato de moneda.
#' @param separador_decimales Separador de decimales para el formato de moneda.
#'
#' @param colores_series Vector de colores HEX con nombre (opcional) para las
#'   series, usando como nombres las etiquetas de \code{etiquetas_series}. Si
#'   se omite, \pkg{ggplot2} asignará colores por defecto.
#'
#' @param mostrar_valores Lógico; si \code{TRUE}, muestra las etiquetas de
#'   valor para cada barra.
#' @param umbral_etiqueta Mínimo valor (en escala original) para mostrar
#'   etiquetas. Debajo de este umbral, el texto se oculta.
#' @param umbral_interno Umbral para decidir si la etiqueta se muestra dentro
#'   de la barra (centrada) o justo al final hacia la derecha. Valores mayores
#'   o iguales a \code{umbral_interno} se dibujan dentro de la barra; valores
#'   entre \code{umbral_etiqueta} y \code{umbral_interno} se dibujan fuera.
#'
#' @param mostrar_barra_extra Lógico; si \code{TRUE} y \code{var_n} no es
#'   \code{NULL}, agrega una "barra extra" de texto al extremo derecho de cada
#'   categoría (típicamente el N, por ejemplo `"N=50"`).
#' @param prefijo_barra_extra Texto añadido antes del valor mostrado en la
#'   barra extra (por defecto \code{"N="}).
#' @param titulo_barra_extra Texto que se coloca como encabezado encima de la
#'   columna de barra extra (por ejemplo, `"Total"`). Si es \code{NULL}, no se
#'   agrega encabezado.
#'
#' @param titulo Título del gráfico.
#' @param subtitulo Subtítulo opcional del gráfico.
#' @param nota_pie Nota o fuente para colocar en el pie de página del gráfico.
#'
#' @param color_titulo Color del título.
#' @param size_titulo Tamaño del título.
#' @param color_subtitulo Color del subtítulo.
#' @param size_subtitulo Tamaño del subtítulo.
#' @param color_nota_pie Color del texto del pie de página.
#' @param size_nota_pie Tamaño del texto del pie de página.
#' @param color_leyenda Color del texto de la leyenda.
#' @param size_leyenda Tamaño del texto de la leyenda.
#' @param color_texto_barras Color del texto de las etiquetas de valor.
#' @param size_texto_barras Tamaño del texto de las etiquetas de valor.
#' @param color_barra_extra Color del texto de la barra extra (N).
#' @param size_barra_extra Tamaño del texto de la barra extra.
#' @param color_ejes Color del texto de las categorías en el eje de categorías.
#' @param size_ejes Tamaño del texto de las categorías en el eje de categorías.
#'
#' @param extra_derecha_rel Porcentaje adicional de espacio a la derecha,
#'   relativo a la barra más larga, para ubicar la barra extra (N).
#' @param mostrar_leyenda Lógico; si \code{FALSE}, oculta la leyenda (útil
#'   si solo hay una serie).
#' @param invertir_leyenda Lógico; si \code{TRUE}, invierte el orden de los
#'   ítems de la leyenda.
#' @param invertir_barras Lógico; si \code{TRUE}, invierte el orden en que las
#'   categorías aparecen (arriba/abajo en horizontal o izquierda/derecha en
#'   vertical).
#' @param textos_negrita Vector de caracteres que indica qué elementos deben
#'   mostrarse en negrita. Puede incluir cualquiera de:
#'   \code{"titulo"}, \code{"valores"}, \code{"leyenda"}, \code{"barra_extra"}.
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
graficar_barras_media_agrupadas <- function(
    data,
    var_categoria,
    var_n                = NULL,
    vars_media,
    etiquetas_series,
    orientacion          = c("horizontal", "vertical"),
    formato_valor        = c("numero", "moneda"),
    decimales            = 1,
    simbolo_moneda       = "S/",
    separador_miles      = ".",
    separador_decimales  = ",",
    colores_series       = NULL,
    mostrar_valores      = TRUE,
    umbral_etiqueta      = 0.03,
    umbral_interno       = 0.15,
    mostrar_barra_extra  = FALSE,
    prefijo_barra_extra  = "N=",
    titulo_barra_extra   = NULL,
    titulo               = NULL,
    subtitulo            = NULL,
    nota_pie             = NULL,
    # Estilo de texto y layout
    color_titulo         = "#000000",
    size_titulo          = 11,
    color_subtitulo      = "#000000",
    size_subtitulo       = 9,
    color_nota_pie       = "#000000",
    size_nota_pie        = 8,
    color_leyenda        = "#000000",
    size_leyenda         = 8,
    color_texto_barras   = "#000000",
    size_texto_barras    = 3,
    color_barra_extra    = "#000000",
    size_barra_extra     = 3,
    color_ejes           = "#000000",
    size_ejes            = 9,
    extra_derecha_rel    = 0.10,
    mostrar_leyenda      = TRUE,
    invertir_leyenda     = FALSE,
    invertir_barras      = FALSE,
    textos_negrita       = NULL,
    exportar             = c("rplot", "png", "ppt"),
    path_salida          = NULL,
    ancho                = 10,
    alto                 = 6,
    dpi                  = 300
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y

  orientacion   <- match.arg(orientacion)
  formato_valor <- match.arg(formato_valor)
  exportar      <- match.arg(exportar)

  # ---------------------------------------------------------------------------
  # 0. Validaciones básicas
  # ---------------------------------------------------------------------------
  if (!var_categoria %in% names(data)) {
    stop("`var_categoria` no existe en `data`.", call. = FALSE)
  }
  if (!is.null(var_n) && !var_n %in% names(data)) {
    stop("`var_n` no existe en `data`.", call. = FALSE)
  }
  if (!all(vars_media %in% names(data))) {
    faltan <- vars_media[!vars_media %in% names(data)]
    stop(
      "Las siguientes columnas de `vars_media` no existen en `data`: ",
      paste(faltan, collapse = ", "),
      call. = FALSE
    )
  }
  if (!all(names(etiquetas_series) %in% vars_media)) {
    stop(
      "Los nombres de `etiquetas_series` deben coincidir con columnas de `vars_media`.",
      call. = FALSE
    )
  }

  textos_negrita <- textos_negrita %||% character(0)
  df <- data

  # ---------------------------------------------------------------------------
  # 1. Ancho → largo
  # ---------------------------------------------------------------------------
  cols_sel <- c(var_categoria, vars_media)
  if (!is.null(var_n)) {
    cols_sel <- c(cols_sel, var_n)
  }

  df_long <- df |>
    dplyr::select(dplyr::all_of(cols_sel)) |>
    tidyr::pivot_longer(
      cols      = dplyr::all_of(vars_media),
      names_to  = ".col_media",
      values_to = ".valor"
    )

  if (!is.numeric(df_long$.valor)) {
    stop("Las columnas de `vars_media` deben ser numéricas.", call. = FALSE)
  }

  # Etiquetas legibles de series
  df_long$.serie <- dplyr::recode(df_long$.col_media, !!!etiquetas_series)

  # Orden de series según etiquetas_series
  df_long$.serie <- factor(df_long$.serie, levels = unname(etiquetas_series))

  # Orden de categorías según aparición (con opción de invertir)
  cat_vec  <- df_long[[var_categoria]]
  cat_lvls <- unique(cat_vec)
  if (invertir_barras) {
    cat_lvls <- rev(cat_lvls)
  }
  df_long[[var_categoria]] <- factor(cat_vec, levels = cat_lvls)

  # Máximo valor para escala y barra extra
  max_valor <- max(df_long$.valor, na.rm = TRUE)
  y_max     <- max_valor * (1 + extra_derecha_rel)

  # ---------------------------------------------------------------------------
  # 2. Gráfico base (antes de coord_flip)
  # ---------------------------------------------------------------------------
  p <- ggplot2::ggplot(
    df_long,
    ggplot2::aes_string(
      x    = var_categoria,
      y    = ".valor",
      fill = ".serie"
    )
  ) +
    ggplot2::geom_col(
      position = ggplot2::position_dodge(width = 0.7),
      width    = 0.6
    )

  # ---------------------------------------------------------------------------
  # 3. Etiquetas de valor (dentro o al final de la barra)
  # ---------------------------------------------------------------------------
  if (mostrar_valores) {

    df_lab <- df_long

    # Construir etiqueta numérica o monetaria
    if (formato_valor == "numero") {
      df_lab$lab <- scales::number(
        df_lab$.valor,
        accuracy = 10^(-decimales),
        big.mark = separador_miles,
        decimal.mark = separador_decimales
      )
    } else {
      df_lab$lab <- paste0(
        simbolo_moneda, " ",
        scales::number(
          df_lab$.valor,
          accuracy = 10^(-decimales),
          big.mark = separador_miles,
          decimal.mark = separador_decimales
        )
      )
    }

    # Ocultar etiquetas por debajo del umbral mínimo
    df_lab$mostrar <- df_lab$.valor >= umbral_etiqueta

    df_lab_in  <- df_lab[df_lab$mostrar & df_lab$.valor >= umbral_interno, , drop = FALSE]
    df_lab_out <- df_lab[df_lab$mostrar & df_lab$.valor <  umbral_interno, , drop = FALSE]

    # Etiquetas internas (centradas en la barra)
    if (nrow(df_lab_in) > 0) {
      p <- p +
        ggplot2::geom_text(
          data        = df_lab_in,
          mapping     = ggplot2::aes_string(
            x     = var_categoria,
            y     = ".valor / 2",
            label = "lab",
            group = ".serie"
          ),
          inherit.aes = FALSE,
          position    = ggplot2::position_dodge(width = 0.7),
          hjust       = 0.5,
          vjust       = 0.5,
          color       = color_texto_barras,
          size        = size_texto_barras,
          fontface    = if ("valores" %in% textos_negrita) "bold" else "plain",
          show.legend = FALSE
        )
    }

    # Etiquetas externas (ligeramente a la derecha del extremo de la barra)
    if (nrow(df_lab_out) > 0) {
      offset_lab <- max_valor * 0.02
      df_lab_out$valor_label <- df_lab_out$.valor + offset_lab

      p <- p +
        ggplot2::geom_text(
          data        = df_lab_out,
          mapping     = ggplot2::aes_string(
            x     = var_categoria,
            y     = "valor_label",
            label = "lab",
            group = ".serie"
          ),
          inherit.aes = FALSE,
          position    = ggplot2::position_dodge(width = 0.7),
          hjust       = 0,
          vjust       = 0.5,
          color       = color_texto_barras,
          size        = size_texto_barras,
          fontface    = if ("valores" %in% textos_negrita) "bold" else "plain",
          show.legend = FALSE
        )
    }
  }

  # ---------------------------------------------------------------------------
  # 4. Escala del eje de valores + espacio extra
  # ---------------------------------------------------------------------------
  p <- p +
    ggplot2::scale_y_continuous(
      limits = c(0, y_max),
      expand = ggplot2::expansion(mult = c(0, 0.05))
    )

  # ---------------------------------------------------------------------------
  # 5. Barra extra con N a la derecha (opcional)
  # ---------------------------------------------------------------------------
  if (mostrar_barra_extra && !is.null(var_n)) {
    df_extra <- df |>
      dplyr::select(dplyr::all_of(c(var_categoria, var_n))) |>
      dplyr::distinct() |>
      dplyr::mutate(
        ypos      = max_valor * (1 + extra_derecha_rel * 0.7),
        ypos_tit  = max_valor * (1 + extra_derecha_rel * 0.95),
        lab_extra = paste0(prefijo_barra_extra, .data[[var_n]])
      )

    # Texto principal N=...
    p <- p +
      ggplot2::geom_text(
        data        = df_extra,
        mapping     = ggplot2::aes_string(
          x     = var_categoria,
          y     = "ypos",
          label = "lab_extra"
        ),
        inherit.aes = FALSE,
        hjust       = 0.5,
        vjust       = 0.5,
        size        = size_barra_extra,
        color       = color_barra_extra,
        fontface    = if ("barra_extra" %in% textos_negrita) "bold" else "plain"
      )

    # Encabezado de la columna de N (solo una vez)
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
              y = "ypos_tit"
            ),
            label       = titulo_barra_extra,
            inherit.aes = FALSE,
            hjust       = 0.5,
            vjust       = 0.5,
            size        = size_barra_extra,
            color       = color_barra_extra,
            fontface    = if ("barra_extra" %in% textos_negrita) "bold" else "plain"
          )
      }
    }
  }

  # ---------------------------------------------------------------------------
  # 6. Colores, orientación y tema
  # ---------------------------------------------------------------------------
  if (!is.null(colores_series)) {
    p <- p +
      ggplot2::scale_fill_manual(values = colores_series)
  }

  p <- p + ggplot2::theme_minimal(base_size = 9)

  # Ajustes según orientación
  if (orientacion == "horizontal") {
    p <- p +
      ggplot2::coord_flip() +
      ggplot2::theme(
        panel.grid.major.y = ggplot2::element_blank(),
        panel.grid.minor   = ggplot2::element_blank(),
        panel.grid.major.x = ggplot2::element_blank(),
        axis.title.x       = ggplot2::element_blank(),
        axis.text.x        = ggplot2::element_blank(),
        axis.ticks.x       = ggplot2::element_blank(),
        axis.title.y       = ggplot2::element_blank(),
        axis.text.y        = ggplot2::element_text(
          color = color_ejes,
          size  = size_ejes,
          hjust = 0.5,
          vjust = 0.5
        ),
        axis.line.y        = ggplot2::element_line(color = color_ejes, linewidth = 0.3)
      )
  } else {
    # Vertical: categorías en X, eje Y numérico oculto
    p <- p +
      ggplot2::theme(
        panel.grid.major.y = ggplot2::element_blank(),
        panel.grid.minor   = ggplot2::element_blank(),
        panel.grid.major.x = ggplot2::element_blank(),
        axis.title.y       = ggplot2::element_blank(),
        axis.text.y        = ggplot2::element_blank(),
        axis.ticks.y       = ggplot2::element_blank(),
        axis.title.x       = ggplot2::element_blank(),
        axis.text.x        = ggplot2::element_text(
          color = color_ejes,
          size  = size_ejes,
          hjust = 0.5,
          vjust = 0.5
        ),
        axis.line.x        = ggplot2::element_line(color = color_ejes, linewidth = 0.3)
      )
  }

  # Leyenda
  p <- p +
    ggplot2::theme(
      legend.title    = ggplot2::element_blank(),
      legend.position = if (mostrar_leyenda) "bottom" else "none",
      legend.text     = ggplot2::element_text(
        color = color_leyenda,
        size  = size_leyenda,
        face  = if ("leyenda" %in% textos_negrita) "bold" else "plain"
      ),
      plot.margin     = ggplot2::margin(t = 15, r = 80, b = 15, l = 5),
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
        hjust = 0,
        color = color_nota_pie,
        size  = size_nota_pie
      )
    )

  if (invertir_leyenda && mostrar_leyenda) {
    p <- p + ggplot2::guides(fill = ggplot2::guide_legend(reverse = TRUE))
  }

  p <- p + ggplot2::labs(
    title    = titulo,
    subtitle = subtitulo,
    caption  = nota_pie
  )

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
