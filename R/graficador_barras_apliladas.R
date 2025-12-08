# =============================================================================
# graficar_barras_apiladas()
# -----------------------------------------------------------------------------
# Graficador genérico de barras horizontales apiladas a partir de una tabla
# de indicadores en formato ancho, pensado para reportes tipo ACNUR / OPS.
# =============================================================================

#' Graficar barras horizontales apiladas a partir de una tabla ancha
#'
#' Genera un gráfico de barras horizontales apiladas donde cada categoría
#' (por ejemplo, enfermedades, regiones o grupos etarios) se muestra como una
#' barra compuesta por proporciones asociadas a diferentes estados o respuestas.
#' La función está pensada para usarse directamente con tablas de indicadores
#' ya calculadas (formato ancho), sin necesidad de reestructurar los datos
#' manualmente.
#'
#' @param data Data frame en formato ancho que contiene la categoría, la base N
#'   y las columnas de porcentajes.
#' @param var_categoria Nombre (string) de la columna que define cada barra
#'   (por ejemplo, `"condicion"`).
#' @param var_n Nombre (string) de la columna con el N total por categoría
#'   (se usa como base para la barra extra; típicamente el N).
#' @param cols_porcentaje Vector de nombres de columnas con los porcentajes
#'   que deben formar los segmentos apilados.
#' @param etiquetas_grupos Vector de caracteres con nombre (named) que asigna
#'   etiquetas legibles a las columnas de \code{cols_porcentaje}. Los nombres
#'   deben coincidir exactamente con los nombres de \code{cols_porcentaje}.
#' @param escala_valor Escala de los porcentajes: \code{"proporcion_1"} si los
#'   valores están en 0–1, o \code{"proporcion_100"} si están en 0–100.
#' @param colores_grupos Vector de colores HEX con nombre (opcional) para los
#'   segmentos, usando como nombres las etiquetas de \code{etiquetas_grupos}.
#'   Si se omite, \pkg{ggplot2} asignará colores por defecto.
#' @param mostrar_valores Lógico; si \code{TRUE}, muestra porcentajes dentro
#'   de cada segmento.
#' @param decimales Número de decimales para los porcentajes internos.
#' @param umbral_etiqueta Mínima proporción para mostrar texto en tamaño
#'   normal dentro de un segmento (por ejemplo, 0.03 equivale a 3\%).
#' @param umbral_etiqueta_peq Mínima proporción para mostrar texto en tamaño
#'   reducido. Los segmentos con proporción entre \code{umbral_etiqueta_peq}
#'   y \code{umbral_etiqueta} se muestran con letra más pequeña. Si es
#'   \code{NULL}, se asume igual a \code{umbral_etiqueta}.
#'
#' @param mostrar_barra_extra Lógico; si \code{TRUE}, agrega una “barra extra”
#'   de texto al extremo derecho de cada barra (típicamente para mostrar el N
#'   o algún indicador agregado como “Top 2 Box”).
#' @param barra_extra_preset Preset para la barra extra. Puede ser:
#'   \itemize{
#'     \item \code{"ninguno"}: comportamiento manual, usa \code{prefijo_barra_extra}
#'       y \code{titulo_barra_extra} con el valor de \code{var_n}.
#'     \item \code{"totales"}: muestra N total; por defecto
#'       \code{prefijo_barra_extra = "N="}, \code{titulo_barra_extra = "Total"}.
#'     \item \code{"top2box"}: suma los 2 valores más altos de cada fila.
#'     \item \code{"top3box"}: suma los 3 valores más altos de cada fila.
#'     \item \code{"bottom2box"}: suma los 2 valores más bajos de cada fila.
#'   }
#' @param prefijo_barra_extra Texto añadido antes del valor mostrado en la
#'   barra extra (por defecto \code{"N="}). Puede ser sobreescrito por el preset.
#' @param titulo_barra_extra Texto que se coloca como encabezado encima de la
#'   columna de barra extra (por ejemplo, "Total", "Top 2 Box"). Si es
#'   \code{NULL}, no se agrega encabezado. Puede ser sobreescrito por el preset.
#'
#' @param titulo Título del gráfico.
#' @param subtitulo Subtítulo opcional del gráfico.
#' @param nota_pie Nota o fuente para colocar en el pie de página del gráfico.
#' @param nota_pie_derecha Texto adicional para el pie de página que se
#'   concatenará a la derecha de \code{nota_pie} (por ejemplo, "N total = 52").
#'   Ambos se muestran en una única línea de caption.
#' @param pos_titulo Alineación horizontal del título. Puede ser
#'   \code{"centro"}, \code{"izquierda"} o \code{"derecha"}.
#' @param pos_nota_pie Alineación horizontal de la nota al pie (caption). Puede ser
#'   \code{"derecha"}, \code{"izquierda"} o \code{"centro"}.
#' @param centro_cowplot Posición horizontal de la leyenda cuando se use un
#'   lienzo compuesto con \pkg{cowplot}. Debe ser un número entre 0 y 1 que
#'   se interpreta como la coordenada \code{x} de la leyenda en
#'   \code{draw_grob()} (\code{0} = totalmente a la izquierda,
#'   \code{1} = totalmente a la derecha). Si se deja como \code{NA},
#'   la función puede calcular una posición automática en función del número
#'   de categorías de la leyenda.
#' @param color_titulo Color del título.
#' @param size_titulo Tamaño del título.
#' @param color_subtitulo Color del subtítulo.
#' @param size_subtitulo Tamaño del subtítulo.
#' @param color_nota_pie Color del texto del pie de página.
#' @param size_nota_pie Tamaño del texto del pie de página.
#' @param color_leyenda Color del texto de la leyenda.
#' @param size_leyenda Tamaño del texto de la leyenda.
#' @param color_texto_barras Color del texto dentro de los segmentos.
#' @param size_texto_barras Tamaño del texto dentro de los segmentos.
#' @param size_texto_barras_peq Tamaño del texto para etiquetas de segmentos
#'   pequeños (ver \code{umbral_etiqueta_peq}). Si es \code{NULL}, se usa el
#'   mismo valor de \code{size_texto_barras}.
#' @param color_barra_extra Color base del texto de la barra extra (se puede
#'   sobreescribir internamente si se usa un preset).
#' @param size_barra_extra Tamaño del texto de la barra extra.
#' @param color_ejes Color del texto de las categorías en el eje Y.
#' @param size_ejes Tamaño del texto de las categorías en el eje Y.
#' @param color_fondo Color de fondo del gráfico. Si es \code{NA} (por defecto),
#'   el fondo será transparente (útil para insertar en PPT/Word).
#' @param grosor_barras Ancho relativo de las barras (argumento `width` de
#'   \code{geom_col()}). Toma valores típicos entre 0 y 1. Por defecto 0.7.
#'
#' @param extra_derecha_rel Porcentaje adicional a la derecha para ubicar la
#'   barra extra, relativo a la longitud de la barra más larga.
#' @param espacio_izquierda_rel Proporción del rango del eje X que se deja
#'   como espacio en blanco a la izquierda (entre el borde del panel y el
#'   comienzo de las barras). Por defecto 0.05 (5\%).
#' @param ancho_max_eje_y Número aproximado de caracteres por línea para las
#'   etiquetas del eje Y. Si no es \code{NULL}, se insertan saltos de línea
#'   automáticos entre palabras usando \pkg{stringr}.
#'
#' @param mostrar_leyenda Lógico; si \code{FALSE}, oculta la leyenda.
#' @param usar_leyenda_cowplot Lógico; si \code{TRUE}, recompone el gráfico
#'   usando \pkg{cowplot} para colocar la leyenda en un bloque inferior que
#'   ocupa ~5\% del alto total, manteniendo las barras en ~95\%. Requiere
#'   tener instalado el paquete \pkg{cowplot}.
#' @param invertir_leyenda Lógico; si \code{TRUE}, invierte el orden de los
#'   ítems de la leyenda sin alterar el orden de las barras ni de los segmentos.
#' @param invertir_barras Lógico; si \code{TRUE}, invierte el orden en que las
#'   categorías aparecen en el eje Y (orden vertical de las barras).
#' @param invertir_segmentos Lógico; si \code{TRUE}, invierte el orden en que se
#'   apilan los segmentos dentro de cada barra (orden horizontal de los colores).
#' @param textos_negrita Vector de caracteres que indica qué elementos deben
#'   mostrarse en negrita. Puede incluir cualquiera de:
#'   \code{"titulo"}, \code{"porcentajes"}, \code{"leyenda"},
#'   \code{"barra_extra"}, \code{"eje_y"}.
#'
#' @param exportar Método de salida:
#'   \itemize{
#'     \item \code{"rplot"}: devuelve un objeto \code{ggplot}.
#'     \item \code{"png"}: exporta un archivo PNG.
#'     \item \code{"ppt"}: exporta un PPTX con el gráfico vectorial editable.
#'     \item \code{"word"}: exporta un DOCX con el gráfico vectorial editable.
#'   }
#' @param path_salida Ruta del archivo a crear cuando \code{exportar} no es
#'   \code{"rplot"}.
#' @param ancho Ancho del gráfico (cuando se exporta a archivo).
#' @param alto Alto del gráfico (cuando se exporta a archivo).
#' @param alto_por_categoria Alto (en pulgadas) a asignar por categoría en el
#'   área de barras cuando no se especifica \code{alto}. Sobre ese alto se suma
#'   automáticamente un bloque fijo para la leyenda (si \code{mostrar_leyenda})
#'   y otro más pequeño para el caption (si existe), para evitar que la barra
#'   desaparezca cuando hay pocas categorías.
#' @param dpi Resolución en puntos por pulgada (solo para PNG).
#'
#' @return Un objeto \code{ggplot} si \code{exportar = "rplot"}. De forma
#'   invisible, el gráfico exportado si se utiliza \code{"png"}, \code{"ppt"}
#'   o \code{"word"}.
#'
#' @export
graficar_barras_apiladas <- function(
    data,
    var_categoria,
    var_n,
    cols_porcentaje,
    etiquetas_grupos,
    escala_valor          = c("proporcion_1", "proporcion_100"),
    colores_grupos        = NULL,
    mostrar_valores       = TRUE,
    decimales             = 1,
    umbral_etiqueta       = 0.03,
    umbral_etiqueta_peq   = NULL,
    mostrar_barra_extra   = TRUE,
    barra_extra_preset    = c("ninguno", "totales", "top2box", "top3box", "bottom2box"),
    prefijo_barra_extra   = NULL,
    titulo_barra_extra    = NULL,
    barra_extra_vjust     = NULL,
    titulo                = NULL,
    subtitulo             = NULL,
    nota_pie              = NULL,
    nota_pie_derecha      = NULL,
    pos_titulo            = c("centro", "izquierda", "derecha"),
    pos_nota_pie          = c("derecha", "izquierda", "centro"),
    centro_cowplot        = NA_real_,
    # Estilo de texto y layout
    color_titulo          = "#000000",
    size_titulo           = 11,
    color_subtitulo       = "#000000",
    size_subtitulo        = 9,
    color_nota_pie        = "#000000",
    size_nota_pie         = 8,
    color_leyenda         = "#000000",
    size_leyenda          = 8,
    color_texto_barras    = "white",
    size_texto_barras     = 3,
    size_texto_barras_peq = NULL,
    color_barra_extra     = "#000000",
    size_barra_extra      = 3,
    color_ejes            = "#000000",
    size_ejes             = 9,
    color_fondo           = NA,
    grosor_barras         = 0.7,
    extra_derecha_rel     = 0.10,
    espacio_izquierda_rel = 0,
    ancho_max_eje_y       = NULL,
    mostrar_leyenda       = TRUE,
    usar_leyenda_cowplot  = FALSE,
    invertir_leyenda      = FALSE,
    invertir_barras       = FALSE,
    invertir_segmentos    = FALSE,
    textos_negrita        = NULL,
    exportar              = c("rplot", "png", "ppt", "word"),
    path_salida           = NULL,
    ancho                 = 10,
    alto                  = 6,
    alto_por_categoria    = NULL,
    dpi                   = 300
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y

  hjust_from_pos <- function(x) {
    switch(
      x,
      "izquierda" = 0,
      "centro"    = 0.5,
      "derecha"   = 1,
      0.5
    )
  }

  escala_valor       <- match.arg(escala_valor)
  exportar           <- match.arg(exportar)
  barra_extra_preset <- match.arg(barra_extra_preset)
  pos_titulo         <- match.arg(pos_titulo)
  pos_nota_pie       <- match.arg(pos_nota_pie)

  # -----------------------------------------------------------------
  # Normalizar `decimales` y tamaños de texto seguros
  # -----------------------------------------------------------------
  decimales <- suppressWarnings(as.numeric(decimales))
  if (length(decimales) < 1L || !is.finite(decimales[1])) {
    decimales <- 1
  } else {
    decimales <- decimales[1]
  }

  size_texto_barras_peq <- size_texto_barras_peq %||% size_texto_barras

  if (is.null(umbral_etiqueta_peq)) {
    umbral_etiqueta_peq <- umbral_etiqueta
  }

  hjust_titulo  <- hjust_from_pos(pos_titulo)
  hjust_caption <- hjust_from_pos(pos_nota_pie)

  pulso_azul  <- "#002768"
  pulso_verde <- "#5BAF31"

  textos_negrita <- textos_negrita %||% character(0)

  alto_extra_header <- 0

  # ---------------------------------------------------------------------------
  # 0. Validaciones
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
  if (!all(names(etiquetas_grupos) %in% cols_porcentaje)) {
    stop(
      "Los nombres de `etiquetas_grupos` deben coincidir con columnas de `cols_porcentaje`.",
      call. = FALSE
    )
  }

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
    ) |>
    dplyr::mutate(
      .grupo = dplyr::recode(.data$.col_pct, !!!etiquetas_grupos)
    )

  if (!is.numeric(df_long$.valor)) {
    stop("Las columnas de porcentaje deben ser numéricas.", call. = FALSE)
  }

  # Escala interna: proporción (si viene en 0–100 la llevamos a 0–1)
  if (escala_valor == "proporcion_100") {
    df_long$.valor_plot <- df_long$.valor / 100
  } else {
    df_long$.valor_plot <- df_long$.valor
  }

  # ---------------------------------------------------------------------------
  # 1.1 Orden de grupos (segmentos), con opción de invertir
  # ---------------------------------------------------------------------------
  niveles_grupos <- unname(etiquetas_grupos)
  if (invertir_segmentos) {
    niveles_grupos <- rev(niveles_grupos)
  }
  df_long$.grupo <- factor(df_long$.grupo, levels = niveles_grupos)

  # ---------------------------------------------------------------------------
  # 1.2 Orden de categorías (barras Y), con opción de invertir
  # ---------------------------------------------------------------------------
  cat_vec  <- df_long[[var_categoria]]
  cat_lvls <- unique(cat_vec)
  if (invertir_barras) {
    cat_lvls <- rev(cat_lvls)
  }
  df_long[[var_categoria]] <- factor(cat_vec, levels = cat_lvls)

  # Suma por categoría (longitud total de cada barra)
  df_sum <- df_long |>
    dplyr::group_by(.data[[var_categoria]]) |>
    dplyr::summarise(suma = sum(.valor_plot, na.rm = TRUE), .groups = "drop")

  max_suma <- max(df_sum$suma, na.rm = TRUE)

  # Para gráficos tipo 100% (escala_valor = "proporcion_100"),
  # forzar que el máximo del eje sea siempre 1 (100%)
  if (escala_valor == "proporcion_100") {
    max_suma <- 1
  }

  if (mostrar_barra_extra) {
    x_max <- max_suma * (1 + extra_derecha_rel)
  } else {
    x_max <- max_suma
  }

  # ---------------------------------------------------------------------------
  # 2. Gráfico base (una barra apilada por categoría)
  # ---------------------------------------------------------------------------
  p <- ggplot2::ggplot(
    df_long,
    ggplot2::aes_string(
      x    = ".valor_plot",
      y    = var_categoria,
      fill = ".grupo"
    )
  ) +
    ggplot2::geom_col(width = grosor_barras) +
    ggplot2::geom_vline(
      xintercept = 0,
      color      = NA,
      linewidth  = 0.3
    )

  # ---------------------------------------------------------------------------
  # 3. Etiquetas dentro de las barras (cálculo manual con cumsum)
  # ---------------------------------------------------------------------------
  if (mostrar_valores) {

    # Orden de apilado real (tal como lo usa ggplot2: del ÚLTIMO nivel al primero)
    niveles_fill  <- levels(df_long$.grupo)
    niveles_stack <- rev(niveles_fill)

    df_lab <- df_long |>
      dplyr::group_by(!!rlang::sym(var_categoria)) |>
      dplyr::arrange(
        factor(.grupo, levels = niveles_stack),
        .by_group = TRUE
      ) |>
      dplyr::mutate(
        x_center = cumsum(.valor_plot) - .valor_plot / 2
      ) |>
      dplyr::ungroup()

    # Formato de porcentajes: siempre enteros (redondeados)
    pct_num <- df_lab$.valor_plot * 100
    lab <- sprintf("%d%%", round(pct_num))

    df_lab$lab <- lab

    # Clasificar etiquetas según umbrales (grande / pequeña / oculta)
    df_lab <- df_lab |>
      dplyr::mutate(
        .tamano_etq = dplyr::case_when(
          .valor_plot >= umbral_etiqueta      ~ "grande",
          .valor_plot >= umbral_etiqueta_peq  ~ "peq",
          TRUE                                ~ "ninguna"
        )
      ) |>
      dplyr::filter(
        !is.na(.valor_plot),
        .tamano_etq != "ninguna"
      )

    df_lab_grande <- df_lab[df_lab$.tamano_etq == "grande", , drop = FALSE]
    df_lab_peq    <- df_lab[df_lab$.tamano_etq == "peq",    , drop = FALSE]

    if (nrow(df_lab_grande) > 0) {
      p <- p +
        ggplot2::geom_text(
          data        = df_lab_grande,
          mapping     = ggplot2::aes_string(
            x     = "x_center",
            y     = var_categoria,
            label = "lab"
          ),
          inherit.aes = FALSE,
          color       = color_texto_barras,
          size        = size_texto_barras,
          fontface    = if ("porcentajes" %in% textos_negrita) "bold" else "plain",
          show.legend = FALSE
        )
    }

    if (nrow(df_lab_peq) > 0) {
      p <- p +
        ggplot2::geom_text(
          data        = df_lab_peq,
          mapping     = ggplot2::aes_string(
            x     = "x_center",
            y     = var_categoria,
            label = "lab"
          ),
          inherit.aes = FALSE,
          color       = color_texto_barras,
          size        = size_texto_barras_peq,
          fontface    = if ("porcentajes" %in% textos_negrita) "bold" else "plain",
          show.legend = FALSE
        )
    }
  }

  # ---------------------------------------------------------------------------
  # 4. Escala X con margen proporcional e Y con wrap opcional
  # ---------------------------------------------------------------------------
  p <- p +
    ggplot2::scale_x_continuous(
      limits = c(0, x_max),
      expand = ggplot2::expansion(mult = c(espacio_izquierda_rel, 0.05))
    ) +
    ggplot2::coord_cartesian(clip = "off")

  if (!is.null(ancho_max_eje_y)) {
    if (!requireNamespace("stringr", quietly = TRUE)) {
      stop("Para usar `ancho_max_eje_y` se requiere el paquete 'stringr'.",
           call. = FALSE)
    }
    p <- p +
      ggplot2::scale_y_discrete(
        labels = function(x) stringr::str_wrap(x, width = ancho_max_eje_y)
      )
  }

  # ---------------------------------------------------------------------------
  # 5.1. Preparar tabla base para barra extra (N o Top/Bottom Box)
  # ---------------------------------------------------------------------------
  df_wide_extra <- df |>
    dplyr::select(dplyr::all_of(c(var_categoria, var_n, cols_porcentaje))) |>
    dplyr::mutate(
      valor_extra = .data[[var_n]]
    )

  # Si no se pasa nada, el prefijo por defecto es vacío ("")
  prefijo_extra_int     <- prefijo_barra_extra %||% ""
  titulo_extra_int      <- titulo_barra_extra
  color_barra_extra_int <- color_barra_extra
  fontface_barra_extra  <- if ("barra_extra" %in% textos_negrita) "bold" else "plain"

  if (barra_extra_preset != "ninguno") {

    if (barra_extra_preset == "totales") {

      if (is.null(titulo_barra_extra) || !nzchar(titulo_barra_extra)) {
        titulo_extra_int <- "Total"
      }

      # Solo en el preset "totales" se impone por defecto "N="
      if (is.null(prefijo_barra_extra)) {
        prefijo_extra_int <- "N = "
      }

      # Para totales: respetar el color que venga del estilo, salvo que no se haya
      # pasado nada explícito.
      if (is.null(color_barra_extra)) {
        color_barra_extra_int <- pulso_azul
      }

      fontface_barra_extra <- "bold"

    } else {

      base_mat <- df_wide_extra[, cols_porcentaje, drop = FALSE]
      if (escala_valor == "proporcion_100") {
        base_mat <- base_mat / 100
      }

      ordenado <- t(apply(as.matrix(base_mat), 1, sort, decreasing = TRUE))

      if (barra_extra_preset == "top2box") {
        df_wide_extra$valor_extra <- ordenado[, 1] + ordenado[, 2]

        if (is.null(titulo_barra_extra) || !nzchar(titulo_barra_extra)) {
          titulo_extra_int <- "TOP TWO BOX"
        }

      } else if (barra_extra_preset == "top3box") {
        df_wide_extra$valor_extra <- ordenado[, 1] + ordenado[, 2] + ordenado[, 3]

        if (is.null(titulo_barra_extra) || !nzchar(titulo_barra_extra)) {
          titulo_extra_int <- "TOP THREE BOX"
        }

      } else if (barra_extra_preset == "bottom2box") {
        ncol_mat <- ncol(ordenado)
        df_wide_extra$valor_extra <- ordenado[, ncol_mat] +
          ordenado[, ncol_mat - 1]

        if (is.null(titulo_barra_extra) || !nzchar(titulo_barra_extra)) {
          titulo_extra_int <- "BOTTOM TWO BOX"
        }
      }

      # Pasar a porcentaje 0–100 para impresión
      df_wide_extra$valor_extra <- df_wide_extra$valor_extra * 100

      # Para Top/Bottom Box: forzar verde, sin respetar el color del estilo.
      color_barra_extra_int <- pulso_verde
      fontface_barra_extra  <- "bold"
    }
  }

  # ---------------------------------------------------------------------------
  # 5.2. Barra extra al costado derecho (N, Top Box, etc.)
  # ---------------------------------------------------------------------------
  if (mostrar_barra_extra) {

    # Función auxiliar: 1 decimal como máximo, pero sin ".0"
    .format_pct_clean <- function(x) {
      x_round <- round(x, 1)  # seguimos con precisión 1 decimal
      txt     <- format(
        x_round,
        nsmall    = 1,
        trim      = TRUE,
        scientific = FALSE
      )
      # quitar ".0" al final si corresponde
      txt <- sub("\\.0$", "", txt)
      txt
    }

    df_extra <- df_sum |>
      dplyr::left_join(
        df_wide_extra |>
          dplyr::select(dplyr::all_of(c(var_categoria, "valor_extra"))),
        by = var_categoria
      ) |>
      dplyr::mutate(
        # Centro "teórico" de la columna de barra extra para cada barra
        xpos_col = suma * (1 + extra_derecha_rel * 0.5),
        # Mover un poquito a la izquierda todos los textos (N o %)
        xpos_text = xpos_col - max_suma * 0.015,
        lab_valor = dplyr::case_when(
          barra_extra_preset %in% c("top2box", "top3box", "bottom2box") ~
            paste0(.format_pct_clean(valor_extra), "%"),
          TRUE ~ format(
            valor_extra,
            big.mark   = ",",
            scientific = FALSE,
            trim       = TRUE
          )
        ),
        lab_extra = paste0(prefijo_extra_int, lab_valor)
      )

    # Textos de la barra extra (N, Top2Box, etc.), ligeramente a la izquierda
    p <- p +
      ggplot2::geom_text(
        data        = df_extra,
        mapping     = ggplot2::aes_string(
          x     = "xpos_text",
          y     = var_categoria,
          label = "lab_extra"
        ),
        inherit.aes = FALSE,
        hjust       = 0,
        size        = size_barra_extra,
        color       = color_barra_extra_int,
        fontface    = fontface_barra_extra
      )

    # Encabezado de la barra extra (Total / Top 2 Box / etc.)
    if (!is.null(titulo_extra_int) && nzchar(titulo_extra_int)) {

      lvls         <- levels(df_long[[var_categoria]])
      cat_superior <- tail(lvls, 1)

      # nº de barras en el gráfico
      n_categorias_header <- length(lvls)


      # vjust definido por el usuario → tiene prioridad absoluta
      if (!is.null(barra_extra_vjust)) {
        vjust_header <- barra_extra_vjust
      } else {

        # vjust dinámico:
        # - 1–2 barras  → más arriba (más negativo)
        # - 3 barras    → intermedio
        # - 4+ barras   → como ahora (-6)
        vjust_header <- dplyr::case_when(
          n_categorias_header <= 2 ~ -8,
          n_categorias_header == 3 ~ -7,
          TRUE                     ~ -6
        )
      }

      # Altura extra aproximada necesaria para que el título no se corte
      alto_extra_header <- dplyr::case_when(
        n_categorias_header <= 2 ~ 1.1,
        n_categorias_header == 3 ~ 0.8,
        TRUE                     ~ 0.6
      )

      df_header <- data.frame(
        var_cat     = cat_superior,
        xpos_header = max_suma * (1 + extra_derecha_rel * 0.5)
      )
      names(df_header)[1] <- var_categoria

      p <- p +
        ggplot2::geom_text(
          data        = df_header,
          mapping     = ggplot2::aes_string(
            x = "xpos_header",
            y = var_categoria
          ),
          label       = titulo_extra_int,
          inherit.aes = FALSE,
          hjust       = 0.5,
          vjust       = vjust_header,
          size        = size_barra_extra,
          color       = color_barra_extra_int,
          fontface    = fontface_barra_extra
        )
    }
  }

  # ---------------------------------------------------------------------------
  # 6. Colores, caption, leyenda y wrap automático de etiquetas largas
  # ---------------------------------------------------------------------------

  # LEYENDA — Wrap automático de etiquetas + tamaño fijo del rectángulo
  wrap_fun <- NULL
  if (requireNamespace("stringr", quietly = TRUE)) {
    # si quieres, aquí puedes cambiar a width = 60
    wrap_fun <- function(x) stringr::str_wrap(x, width = 40)
  }

  # Aplicación del wrap a la escala de colores
  if (!is.null(colores_grupos)) {

    # Paleta manual
    p <- p +
      ggplot2::scale_fill_manual(
        values = colores_grupos,
        labels = if (!is.null(wrap_fun)) wrap_fun else waiver()
      )

  } else if (!is.null(wrap_fun)) {

    # Paleta por defecto de ggplot con wrap
    p <- p +
      ggplot2::scale_fill_discrete(labels = wrap_fun)

  }

  # Mantener el mismo tamaño de “swatch” aunque el texto tenga 1 o 2 líneas
  p <- p +
    ggplot2::theme(
      legend.key.width  = grid::unit(0.3, "cm"),
      legend.key.height = grid::unit(0.3, "cm")
    )

  caption_text <- NULL
  if (!is.null(nota_pie) && nzchar(nota_pie) &&
      !is.null(nota_pie_derecha) && nzchar(nota_pie_derecha)) {
    caption_text <- paste0(nota_pie, "   ", nota_pie_derecha)
  } else if (!is.null(nota_pie) && nzchar(nota_pie)) {
    caption_text <- nota_pie
  } else if (!is.null(nota_pie_derecha) && nzchar(nota_pie_derecha)) {
    caption_text <- nota_pie_derecha
  }

  # Número de ítems en la leyenda y filas necesarias (máx. 5 por fila)
  n_items_leyenda <- length(levels(df_long$.grupo))
  n_por_fila      <- 6L
  n_filas_leyenda <- max(1L, ceiling(n_items_leyenda / n_por_fila))

  # Leyenda centrada con filas dinámicas
  if (mostrar_leyenda) {
    p <- p +
      ggplot2::guides(
        fill = ggplot2::guide_legend(
          nrow    = n_filas_leyenda,
          reverse = invertir_leyenda
        )
      )
  }

  p <- p +
    ggplot2::theme_minimal(base_size = 9) +
    ggplot2::theme(
      panel.grid.major.y  = ggplot2::element_blank(),
      panel.grid.minor    = ggplot2::element_blank(),
      panel.grid.major.x  = ggplot2::element_blank(),
      axis.title.x        = ggplot2::element_blank(),
      axis.text.x         = ggplot2::element_blank(),
      axis.ticks.x        = ggplot2::element_blank(),
      axis.title.y        = ggplot2::element_blank(),
      axis.text.y         = ggplot2::element_text(
        color = color_ejes,
        size  = size_ejes,
        hjust = 1,
        vjust = 0.5,
        face  = if ("eje_y" %in% textos_negrita) "bold" else "plain"
      ),
      axis.line.y         = ggplot2::element_blank(),
      legend.title        = ggplot2::element_blank(),
      legend.position     = if (mostrar_leyenda) "bottom" else "none",
      legend.justification = "center",
      legend.box          = "horizontal",
      legend.box.just     = "center",
      legend.text         = ggplot2::element_text(
        color = color_leyenda,
        size  = size_leyenda,
        face  = if ("leyenda" %in% textos_negrita) "bold" else "plain"
      ),
      plot.margin         = ggplot2::margin(t = 15, r = 10, b = 5, l = 10),
      plot.title.position = "plot",
      plot.title          = ggplot2::element_text(
        hjust = hjust_titulo,
        color = color_titulo,
        size  = size_titulo,
        face  = if ("titulo" %in% textos_negrita) "bold" else "plain"
      ),
      plot.subtitle       = ggplot2::element_text(
        hjust = hjust_titulo,
        color = color_subtitulo,
        size  = size_subtitulo
      ),
      plot.caption        = ggplot2::element_text(
        hjust = hjust_caption,
        color = color_nota_pie,
        size  = size_nota_pie
      ),
      plot.background     = ggplot2::element_rect(fill = color_fondo, color = NA),
      panel.background    = ggplot2::element_rect(fill = color_fondo, color = NA)
    ) +
    ggplot2::labs(
      title    = titulo,
      subtitle = subtitulo,
      caption  = caption_text
    )

  # ---------------------------------------------------------------------------
  # 6.bis. Recomposición con cowplot (95% barras / 5% leyenda)
  # ---------------------------------------------------------------------------
  if (mostrar_leyenda && usar_leyenda_cowplot) {

    if (!requireNamespace("cowplot", quietly = TRUE)) {
      stop(
        "Para usar `usar_leyenda_cowplot = TRUE` se requiere el paquete 'cowplot'.",
        call. = FALSE
      )
    }

    # Plot base sin leyenda y con margen inferior casi nulo
    p_base_sin_leyenda <- p +
      ggplot2::theme(
        legend.position = "none",
        plot.margin     = ggplot2::margin(t = 10, r = 10, b = 5, l = 10)
      )

    # Leyenda extraída tal cual se ve en `p`
    leg <- cowplot::get_legend(
      p +
        ggplot2::theme(
          legend.position      = "bottom",
          legend.direction     = "horizontal",
          legend.box           = "horizontal",
          legend.justification = "left",
          legend.box.just      = "left",
          legend.margin        = ggplot2::margin(),
          legend.box.margin    = ggplot2::margin()
        )
    )

    # n_items_leyenda ya lo tienes calculado
    n_items_leyenda <- length(levels(df_long$.grupo))

    # Ancho efectivo por fila (máx 6 por fila)
    ancho_fila <- if (n_items_leyenda <= 6) {
      n_items_leyenda
    } else {
      ceiling(n_items_leyenda / 2)
    }

    # POSICIÓN HORIZONTAL DINÁMICA DE LA LEYENDA
    pos_leyenda_x <- dplyr::case_when(
      ancho_fila <= 2 ~ 0.45,  # 1–2 ítems: bastante centrado
      ancho_fila == 3 ~ 0.38,  # 3 ítems
      ancho_fila == 4 ~ 0.32,  # 4 ítems
      ancho_fila == 5 ~ 0.28,  # 5 ítems
      TRUE           ~ 0.18    # 6 ítems (fila llena)
    )
    # OVERRIDE MANUAL (centro_cowplot)
    if (!is.na(centro_cowplot) && is.finite(centro_cowplot)) {
      pos_leyenda_x <- centro_cowplot
    }

    # Composición 95% / 5%
    p <- cowplot::ggdraw() +
      cowplot::draw_plot(
        p_base_sin_leyenda,
        x      = 0,
        y      = 0.05,  # empieza justo encima de la franja de leyenda
        width  = 1,
        height = 0.95   # 95% del alto para barras + título + caption
      ) +
      cowplot::draw_grob(
        leg,
        x      = pos_leyenda_x,  # mueve la leyenda hacia centro/izquierda
        y      = 0.035,  # pegada a la parte baja, pero tocando el plot
        width  = 1,
        height = 0.05   # 5% del alto total
      )
  }

  # ---------------------------------------------------------------------------
  # 7. Exportación (altura total = panel + leyenda + caption)
  # ---------------------------------------------------------------------------
  n_categorias <- length(unique(df_long[[var_categoria]]))

  # Parámetros de descomposición de altura en pulgadas
  alto_max_total <- 9.0     # tope máximo global
  alto_caption   <- 0.25    # bloque adicional si hay caption

  # Cálculo de alto sugerido para Word/PNG/PPT
  alto_por_cat_eff <- alto_por_categoria %||% 0.35

  alto_panel <- max(n_categorias, 1L) * alto_por_cat_eff

  n_items_leyenda <- length(levels(df_long$.grupo))
  alto_leyenda <- if (mostrar_leyenda && n_items_leyenda > 0) {
    if (n_items_leyenda <= 5)       0.5
    else if (n_items_leyenda <=10)  0.8
    else                            1.1
  } else {
    0
  }

  tiene_caption <- !is.null(caption_text) && nzchar(caption_text)
  alto_cap      <- if (tiene_caption) alto_caption else 0

  alto_total_sugerido <- alto_panel + alto_leyenda + alto_cap + alto_extra_header
  alto_total_sugerido <- min(alto_max_total, alto_total_sugerido)

  alto_min_suave <- (alto_por_cat_eff * 1.2) + alto_leyenda + alto_cap + alto_extra_header
  alto_total_sugerido <- max(alto_min_suave, alto_total_sugerido)

  if (exportar == "rplot") {
    attr(p, "alto_word_sugerido") <- alto_total_sugerido
    return(p)
  }

  if (is.null(path_salida) || !nzchar(path_salida)) {
    stop("Debe especificar `path_salida` cuando `exportar` no es 'rplot'.", call. = FALSE)
  }

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
