# =============================================================================
# graficar_barras_numericas()
# -----------------------------------------------------------------------------
# Barras (horizontales o verticales) agrupadas para comparar 1+ valores
# numéricos (medias, montos, etc.) por categoría.
# =============================================================================

#' Graficar barras agrupadas de valores numéricos por categoría
#'
#' Genera un gráfico de barras agrupadas donde cada categoría (por ejemplo,
#' servicios, enfermedades, regiones) se muestra en el eje de categorías y,
#' para cada una, se comparan una o más medidas numéricas (por ejemplo,
#' "Tarifa promedio", "Tarifa mínima"), diferenciadas por color y con
#' leyenda inferior.
#'
#' La función está pensada para usarse directamente con tablas de indicadores
#' ya calculadas (formato ancho), sin necesidad de reestructurar los datos
#' manualmente.
#'
#' @param data Data frame en formato ancho que contiene la categoría, la base N
#'   (opcional) y las columnas de valores numéricos.
#' @param var_categoria Nombre (string) de la columna que define cada categoría
#'   en el eje (por ejemplo, `"servicio"` o `"condicion"`).
#' @param var_n Nombre (string) de la columna con el N total por categoría.
#'   Se usa para mostrar una etiqueta adicional al extremo derecho de las barras
#'   (por ejemplo `"N=50"`). Si es \code{NULL}, no se muestra la barra extra.
#' @param vars_valor Vector de nombres de columnas con los valores numéricos
#'   que se desean comparar como series agrupadas (por ejemplo,
#'   \code{c("promedio_total", "promedio_consulta")}).
#' @param etiquetas_series Vector de caracteres con nombre (named) que asigna
#'   etiquetas legibles a las columnas de \code{vars_valor}. Los nombres deben
#'   coincidir exactamente con \code{vars_valor}; los valores se usan en la
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
#' @param separador_miles Separador de miles para el formato numérico/moneda.
#' @param separador_decimales Separador de decimales para el formato
#'   numérico/moneda.
#'
#' @param colores_series Vector de colores HEX con nombre (opcional) para las
#'   series, usando como nombres las etiquetas de \code{etiquetas_series}. Si
#'   se omite, \pkg{ggplot2} asignará colores por defecto.
#'
#' @param mostrar_valores Lógico; si \code{TRUE}, muestra las etiquetas de
#'   valor para cada barra.
#' @param umbral_etiqueta Mínimo valor (en escala original) para mostrar
#'   etiquetas. Debajo de este umbral, el texto se oculta.
#' @param umbral_interno Umbral (en escala original) para decidir si la
#'   etiqueta se muestra dentro de la barra (centrada) o justo al final hacia
#'   la derecha. Valores mayores o iguales a \code{umbral_interno} se dibujan
#'   dentro de la barra; valores entre \code{umbral_etiqueta} y
#'   \code{umbral_interno} se dibujan fuera.
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
#' @param pos_titulo Alineación horizontal del título: `"centro"`,
#'   `"izquierda"` o `"derecha"`.
#' @param pos_nota_pie Alineación horizontal de la nota al pie (caption):
#'   `"derecha"`, `"izquierda"` o `"centro"`.
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
#' @param color_fondo Color de fondo del panel y del plot (por defecto
#'   \code{NA}, transparente).
#'
#' @param extra_derecha_rel Porcentaje adicional de espacio a la derecha,
#'   relativo a la barra más larga, para ubicar la barra extra (N).
#' @param ancho_max_eje_cat Número aproximado de caracteres por línea para las
#'   etiquetas de las categorías. Si no es \code{NULL}, se insertan saltos de
#'   línea automáticos entre palabras usando \pkg{stringr}.
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
#'   \code{ggplot}), \code{"png"} (exporta un archivo PNG),
#'   \code{"ppt"} (exporta una diapositiva PPTX) o \code{"word"} (exporta un
#'   documento Word con el gráfico incrustado).
#' @param path_salida Ruta del archivo a crear cuando \code{exportar} no es
#'   \code{"rplot"}.
#' @param ancho Ancho del gráfico (cuando se exporta a archivo).
#' @param alto Alto del gráfico (cuando se exporta a archivo).
#' @param alto_por_categoria Alto (en pulgadas) a asignar por categoría en el
#'   área de barras cuando no se especifica \code{alto}. Sobre ese alto se suma
#'   automáticamente un bloque fijo para la leyenda (si \code{mostrar_leyenda})
#'   y otro más pequeño para el caption (si existe), y luego se acota entre un
#'   mínimo y un máximo razonables para mantener la consistencia visual.
#' @param dpi Resolución en puntos por pulgada (solo para PNG).
#'
#' @return Un objeto \code{ggplot} si \code{exportar = "rplot"}. De forma
#'   invisible, el gráfico exportado si se utiliza \code{"png"}, \code{"ppt"}
#'   o \code{"word"}.
#' @export
graficar_barras_numericas <- function(
    data,
    var_categoria,
    var_n                = NULL,
    vars_valor,
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
    pos_titulo           = c("centro", "izquierda", "derecha"),
    pos_nota_pie         = c("derecha", "izquierda", "centro"),
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
    color_fondo          = NA,
    extra_derecha_rel    = 0.10,
    ancho_max_eje_cat    = NULL,
    mostrar_leyenda      = TRUE,
    invertir_leyenda     = FALSE,
    invertir_barras      = FALSE,
    textos_negrita       = NULL,
    exportar             = c("rplot", "png", "ppt", "word"),
    path_salida          = NULL,
    ancho                = 10,
    alto                 = 6,
    alto_por_categoria   = NULL,
    dpi                  = 300
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y

  orientacion   <- match.arg(orientacion)
  formato_valor <- match.arg(formato_valor)
  exportar      <- match.arg(exportar)
  pos_titulo    <- match.arg(pos_titulo)
  pos_nota_pie  <- match.arg(pos_nota_pie)

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
  # 0. Validaciones básicas
  # ---------------------------------------------------------------------------
  if (!var_categoria %in% names(data)) {
    stop("`var_categoria` no existe en `data`.", call. = FALSE)
  }
  if (!is.null(var_n) && !var_n %in% names(data)) {
    stop("`var_n` no existe en `data`.", call. = FALSE)
  }
  if (!all(vars_valor %in% names(data))) {
    faltan <- vars_valor[!vars_valor %in% names(data)]
    stop(
      "Las siguientes columnas de `vars_valor` no existen en `data`: ",
      paste(faltan, collapse = ", "),
      call. = FALSE
    )
  }
  if (!all(names(etiquetas_series) %in% vars_valor)) {
    stop(
      "Los nombres de `etiquetas_series` deben coincidir con columnas de `vars_valor`.",
      call. = FALSE
    )
  }

  textos_negrita <- textos_negrita %||% character(0)
  df <- data

  # ---------------------------------------------------------------------------
  # 1. Ancho → largo
  # ---------------------------------------------------------------------------
  cols_sel <- c(var_categoria, vars_valor)
  if (!is.null(var_n)) {
    cols_sel <- c(cols_sel, var_n)
  }

  df_long <- df |>
    dplyr::select(dplyr::all_of(cols_sel)) |>
    tidyr::pivot_longer(
      cols      = dplyr::all_of(vars_valor),
      names_to  = ".col_val",
      values_to = ".valor"
    )

  if (!is.numeric(df_long$.valor)) {
    stop("Las columnas de `vars_valor` deben ser numéricas.", call. = FALSE)
  }

  # Etiquetas legibles de series
  df_long$.serie <- dplyr::recode(df_long$.col_val, !!!etiquetas_series)

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
        accuracy     = 10^(-decimales),
        big.mark     = separador_miles,
        decimal.mark = separador_decimales
      )
    } else {
      df_lab$lab <- paste0(
        simbolo_moneda, " ",
        scales::number(
          df_lab$.valor,
          accuracy     = 10^(-decimales),
          big.mark     = separador_miles,
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
  # 6. Colores, wrap de categorías, orientación y tema
  # ---------------------------------------------------------------------------
  if (!is.null(colores_series)) {
    p <- p +
      ggplot2::scale_fill_manual(values = colores_series)
  }

  # Wrap de etiquetas de categoría (antes de coord_flip)
  if (!is.null(ancho_max_eje_cat)) {
    if (!requireNamespace("stringr", quietly = TRUE)) {
      stop("Para usar `ancho_max_eje_cat` se requiere el paquete 'stringr'.",
           call. = FALSE)
    }
    p <- p +
      ggplot2::scale_x_discrete(
        labels = function(x) stringr::str_wrap(x, width = ancho_max_eje_cat)
      )
  }

  base_theme <- ggplot2::theme_minimal(base_size = 9) +
    ggplot2::theme(
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
      plot.margin        = ggplot2::margin(t = 15, r = 80, b = 15, l = 5),
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

  # Eje de valores visible (ticks + línea)
  eje_valores_theme <- ggplot2::theme(
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

  if (orientacion == "horizontal") {
    # Horizontal → coord_flip
    p <- p +
      ggplot2::coord_flip() +
      base_theme +
      eje_valores_theme +
      ggplot2::theme(
        panel.grid.major.y = ggplot2::element_blank(),
        axis.text.y        = ggplot2::element_text(
          color = color_ejes,
          size  = size_ejes,
          hjust = 1,
          vjust = 0.5
        ),
        axis.line.y        = ggplot2::element_blank()
      )
  } else {
    # Vertical
    p <- p +
      base_theme +
      eje_valores_theme +
      ggplot2::theme(
        panel.grid.major.y = ggplot2::element_blank(),
        axis.text.x        = ggplot2::element_text(
          color = color_ejes,
          size  = size_ejes,
          hjust = 0.5,
          vjust = 0.5
        ),
        axis.line.y        = ggplot2::element_blank()
      )
  }

  # Leyenda en filas (máx. 5 por fila), como en las otras funciones
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

  p <- p + ggplot2::labs(
    title    = titulo,
    subtitle = subtitulo,
    caption  = nota_pie
  )

  # ---------------------------------------------------------------------------
  # 7. Exportación (altura total = panel + leyenda + caption)
  # ---------------------------------------------------------------------------
  n_categorias <- length(unique(df_long[[var_categoria]]))

  # Parámetros de descomposición de altura (en pulgadas)
  alto_min_total   <- 2.5
  alto_max_total   <- 9.0
  alto_leyenda_row <- 0.35
  alto_caption     <- 0.25

  # Leyenda: cuántos ítems y cuántas filas
  n_items_leyenda <- length(levels(df_long$.serie))
  if (!exists("n_filas_leyenda")) {
    n_filas_leyenda <- if (n_items_leyenda <= 0) {
      0L
    } else {
      ceiling(n_items_leyenda / 5)
    }
  }

  tiene_caption <- !is.null(nota_pie) && nzchar(nota_pie)

  # alto_por_categoria efectivo
  alto_por_cat_eff <- alto_por_categoria %||% 0.35

  alto_panel <- max(n_categorias, 1L) * alto_por_cat_eff

  alto_leyenda <- if (mostrar_leyenda && n_items_leyenda > 0) {
    n_filas_leyenda * alto_leyenda_row
  } else {
    0
  }

  alto_cap <- if (tiene_caption) alto_caption else 0

  alto_total_sugerido <- alto_panel + alto_leyenda + alto_cap
  alto_total_sugerido <- max(alto_min_total, min(alto_max_total, alto_total_sugerido))

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
  } else if (!is.null(alto_por_categoria)) {
    alto_total_sugerido
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
