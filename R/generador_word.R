# =============================================================================
# reporte_word()
# -----------------------------------------------------------------------------
# Generar Word con gráficos a partir de tablas de frecuencias (SO / SM)
# usando los graficadores ya existentes del paquete.
# =============================================================================

#' Generar un reporte Word a partir de frecuencias (SO/SM) con graficadores propios
#'
#' `reporte_word()` toma una base de reporte (idealmente la misma que se usa en
#' `reporte_frecuencias()`), el instrumento y una definición de secciones, y
#' genera:
#' \itemize{
#'   \item Una lista de gráficos \code{ggplot} (uno por variable válida o bloque).
#'   \item Opcionalmente, un archivo Word donde cada gráfico ocupa un bloque:
#'         título externo + gráfico centrado + pie azul con fuente.
#' }
#'
#' El tipo de gráfico para cada variable se decide igual que en `reporte_ppt()`:
#' \enumerate{
#'   \item Sobrescritura explícita por variable:
#'     \itemize{
#'       \item \code{vars_dico}: variables que deben usarse como dicotómicas
#'             con \code{graficar_dico()}.
#'       \item \code{vars_barras_apiladas}: variables para
#'             \code{graficar_barras_apiladas()}.
#'       \item \code{vars_barras_agrupadas}: variables para
#'             \code{graficar_barras_agrupadas()}.
#'       \item \code{vars_radar}: reservado para futuros usos (no se construye
#'             tabla radar automáticamente).
#'     }
#'   \item Decisión por \code{list_name} de la pregunta en el instrumento:
#'     \itemize{
#'       \item \code{listnames_dico}: todas las variables cuyo \code{list_name}
#'             esté aquí tienden a graficarse como dicotómicas.
#'       \item \code{listnames_apiladas}: todas las variables cuyo \code{list_name}
#'             esté aquí tienden a graficarse como barras apiladas.
#'     }
#'   \item Defaults por tipo de pregunta:
#'     \itemize{
#'       \item \code{default_so}: tipo de gráfico por defecto para \code{select_one}.
#'       \item \code{default_sm}: tipo de gráfico por defecto para
#'             \code{select_multiple}.
#'     }
#' }
#'
#' Para las variables dicotómicas, no se asume ninguna pareja de categorías por
#' defecto: se requiere definirla en \code{dico_labels_por_var} o
#' \code{dico_labels_por_listname}. Si no se encuentra pareja válida, se usa
#' barras agrupadas por defecto.
#'
#' En el documento Word:
#' \itemize{
#'   \item El título del gráfico se escribe \emph{fuera} del gráfico, como
#'         \code{"Gráfico Nº X: "} (en negrita y cursiva) seguido del label de la
#'         variable, todo en color azul Pulso.
#'   \item El gráfico se inserta centrado, sin título interno.
#'   \item Debajo del gráfico se escribe un pie centrado en azul con el texto
#'         \code{"N = ..."} seguido de \code{fuente} cuando se indique.
#'   \item Dentro del gráfico, en el caption derecho, se escribe el \code{N}
#'         usando los argumentos \code{nota_pie_derecha} y \code{pos_nota_pie}
#'         de los graficadores.
#' }
#'
#' @inheritParams reporte_ppt
#' @param path_word Ruta del archivo DOCX a generar cuando \code{solo_lista = FALSE}.
#' @param fuente Texto de fuente que se mostrará debajo de cada gráfico
#'   en el documento Word (por ejemplo, `" estudiantes. Fuente: Pulso PUCP 2025."`,
#'   que se concatenará luego de `"N = <N> "`).
#' @param template_docx Ruta a una plantilla DOCX. Si es \code{NULL}, se intenta
#'   usar una plantilla interna \code{"plantillas/plantilla_pulso.docx"} y, si
#'   no existe, se usa la plantilla por defecto de Word.
#' @param bloques_multi_apiladas Lista nombrada de bloques especiales para
#'   multi-apiladas. Cada elemento debe ser una lista con al menos
#'   \code{vars = c("p101","p102",...)} y opcionalmente:
#'   \code{titulo}, \code{wrap_y}, \code{grosor_barras}, \code{invertir_barras}.
#'
#' @return Una lista con:
#' \describe{
#'   \item{plots}{Lista de objetos \code{ggplot} generados (uno por variable o bloque).}
#'   \item{log_decisiones}{Tibble con información sobre cada variable procesada.}
#' }
#'
#' @export
reporte_word <- function(
    data,
    instrumento      = NULL,
    secciones        = NULL,
    path_word        = "reporte_word.docx",
    fuente           = NULL,
    sm_vars_force    = NULL,
    mostrar_todo     = FALSE,
    solo_lista       = FALSE,
    incluir_titulo_var = TRUE,
    mensajes_progreso  = TRUE,
    template_docx      = NULL,

    # Bloques de varias vars apiladas en una misma barra
    bloques_multi_apiladas = NULL,

    # Sobrescritura por variable
    vars_dico             = NULL,
    vars_barras_apiladas  = NULL,
    vars_barras_agrupadas = NULL,
    vars_radar            = NULL,

    # Overrides por list_name
    listnames_dico        = NULL,
    listnames_apiladas    = NULL,

    # Etiquetas explícitas para dicotómicas
    dico_labels_por_var      = list(),
    dico_labels_por_listname = list(),

    # Colores para barras apiladas por list_name
    colores_apiladas_por_listname = list(),

    # Defaults por tipo de pregunta
    # SO → apiladas por defecto, SM → agrupadas
    default_so = c("barras_apiladas", "barras_agrupadas"),
    default_sm = c("barras_agrupadas", "barras_apiladas"),

    # Barra extra en barras apiladas/agrupadas
    barra_extra = c("ninguna", "total_n"),

    # Estilos por tipo de gráfico
    estilos_barras_agrupadas = list(),
    estilos_barras_apiladas  = list(),
    estilos_dico             = list()
) {

  `%||%` <- function(x, y) if (!is.null(x)) x else y

  default_so  <- match.arg(default_so)
  default_sm  <- match.arg(default_sm)
  barra_extra <- match.arg(barra_extra)

  vars_dico             <- vars_dico             %||% character(0)
  vars_barras_apiladas  <- vars_barras_apiladas  %||% character(0)
  vars_barras_agrupadas <- vars_barras_agrupadas %||% character(0)
  vars_radar            <- vars_radar            %||% character(0)

  listnames_dico     <- listnames_dico     %||% character(0)
  listnames_apiladas <- listnames_apiladas %||% character(0)

  if (!is.data.frame(data)) {
    stop("`data` debe ser un data.frame o tibble.", call. = FALSE)
  }

  # ---------------------------------------------------------------------------
  # 1. Instrumento y survey / choices / orders_list
  # ---------------------------------------------------------------------------
  if (is.null(instrumento)) {
    instrumento <- attr(data, "instrumento_reporte", exact = TRUE)
    if (is.null(instrumento)) {
      stop("No se proporcionó `instrumento` y `data` no tiene atributo ",
           "`instrumento_reporte`.", call. = FALSE)
    }
  }

  survey  <- instrumento$survey
  choices <- instrumento$choices
  if (is.null(survey) || !all(c("name", "label") %in% names(survey))) {
    stop("El `instrumento` no contiene un `survey` con columnas `name` y `label`.",
         call. = FALSE)
  }

  orders_list <- instrumento$orders_list %||% NULL

  dic_vars <- survey |>
    dplyr::filter(!is.na(.data$name), .data$name != "") |>
    dplyr::select(name, label) |>
    dplyr::mutate(label = trimws(as.character(.data$label))) |>
    dplyr::distinct(name, .keep_all = TRUE)

  # ---------------------------------------------------------------------------
  # 2. Inferir secciones si no se pasan
  # ---------------------------------------------------------------------------
  if (is.null(secciones)) {
    seccion_col <- NULL
    if ("section" %in% names(survey)) {
      seccion_col <- "section"
    } else if ("seccion" %in% names(survey)) {
      seccion_col <- "seccion"
    }

    if (is.null(seccion_col)) {
      stop("No se especificaron `secciones` y el `survey` no tiene columna ",
           "`section` ni `seccion`.", call. = FALSE)
    }

    secciones_df <- survey |>
      dplyr::filter(
        !is.na(.data[[seccion_col]]),
        !is.na(.data$name),
        .data$name %in% names(data)
      ) |>
      dplyr::select(seccion = !!rlang::sym(seccion_col), name)

    if (nrow(secciones_df) == 0) {
      stop("No se pudieron inferir secciones desde `survey$",
           seccion_col, "`.", call. = FALSE)
    }

    secciones <- split(secciones_df$name, secciones_df$seccion)
  }

  SECCIONES <- lapply(secciones, function(vars) {
    vars[vapply(
      vars,
      function(v) .has_var_or_dummies(data, v),
      logical(1)
    )]
  })
  SECCIONES <- SECCIONES[vapply(SECCIONES, length, integer(1)) > 0L]

  if (length(SECCIONES) == 0L) {
    stop("Después de filtrar por presencia en `data` (variable o dummies), ",
         "ninguna sección tiene variables válidas.", call. = FALSE)
  }

  # ---------------------------------------------------------------------------
  # (NUEVO) Precomputar variables incluidas en bloques multi-apilados
  # ---------------------------------------------------------------------------
  if (!is.null(bloques_multi_apiladas) && length(bloques_multi_apiladas) > 0) {
    vars_multi_all <- unique(unlist(lapply(bloques_multi_apiladas, `[[`, "vars")))
  } else {
    vars_multi_all <- character(0)
  }

  # ---------------------------------------------------------------------------
  # 3. Helpers internos
  # ---------------------------------------------------------------------------

  .titulo_var_safe <- function(var) {
    titulo_var(
      var,
      dic_vars        = dic_vars,
      labels_override = NULL,
      orders_list     = orders_list,
      df              = data
    )
  }

  .tab_freq_var <- function(var) {
    tab <- freq_table_spss(
      data,
      var,
      survey        = survey,
      sm_vars_force = sm_vars_force,
      orders_list   = orders_list,
      mostrar_todo  = mostrar_todo
    )

    if (!nrow(tab)) return(tab)

    tab |>
      dplyr::filter(.data$Opciones != "Total") |>
      dplyr::filter(is.na(.data$n) == FALSE & .data$n > 0)
  }

  # Helper: dado un vector de frecuencias n, devuelve porcentajes ENTEROS
  # que suman exactamente 100 (pensado para barras apiladas / dico).
  .pct_enteros_100 <- function(n) {
    n <- as.numeric(n)
    if (!length(n) || all(is.na(n))) {
      return(numeric(0))
    }

    n[is.na(n)] <- 0
    total <- sum(n)

    if (!is.finite(total) || total <= 0) {
      return(rep(0L, length(n)))
    }

    # Porcentajes crudos en 0–100
    raw_pct   <- n / total * 100
    floor_pct <- floor(raw_pct)

    # Cuánto falta o sobra para llegar a 100
    resid <- as.integer(round(100 - sum(floor_pct)))

    frac <- raw_pct - floor_pct

    if (resid > 0) {
      # Asignar +1 a los mayores restos decimales
      ord <- order(frac, decreasing = TRUE, na.last = TRUE)
      idx <- head(ord, resid)
      floor_pct[idx] <- floor_pct[idx] + 1L
    } else if (resid < 0) {
      # Quitar 1 a los menores restos decimales
      resid_neg <- abs(resid)
      ord <- order(frac, decreasing = FALSE, na.last = TRUE)
      idx <- head(ord, resid_neg)
      floor_pct[idx] <- pmax(0L, floor_pct[idx] - 1L)
    }

    floor_pct
  }

  .extraer_N_total <- function(tab_freq) {
    if ("N" %in% names(tab_freq)) {
      N_vals <- suppressWarnings(as.numeric(unique(tab_freq$N)))
      N_vals <- N_vals[is.finite(N_vals)]
      if (length(N_vals)) {
        return(max(N_vals))
      }
    }
    sum(tab_freq$n, na.rm = TRUE)
  }

  .build_tab_barras_agrupadas <- function(tab_freq, var_label) {
    if (!nrow(tab_freq)) return(NULL)

    n_total <- .extraer_N_total(tab_freq)

    pct_raw <- tab_freq$pct
    if (all(is.na(pct_raw))) return(NULL)

    # Detectar escala de pct (0–1 o 0–100)
    max_pct <- max(pct_raw, na.rm = TRUE)
    if (is.finite(max_pct) && max_pct <= 1 + 1e-8) {
      pct_0_100 <- pct_raw * 100
    } else {
      pct_0_100 <- pct_raw
    }

    # Convertir a enteros (no imponemos suma 100; en SM puede ser > 100)
    pct_int   <- round(pct_0_100)
    pct_prop  <- pct_int / 100

    tibble::tibble(
      categoria = tab_freq$Opciones,
      n_base    = n_total,
      pct       = pct_prop
    )
  }

  .build_tab_barras_apiladas <- function(tab_freq, var_label) {
    if (!nrow(tab_freq)) return(NULL)

    n_total <- .extraer_N_total(tab_freq)
    if (!is.finite(n_total) || n_total <= 0) return(NULL)

    n_cat <- nrow(tab_freq)

    # 1) Porcentajes enteros que suman 100 a partir de los conteos n
    pct_int  <- .pct_enteros_100(tab_freq$n)
    # 2) Proporciones 0–1 para graficar con escala_valor = "proporcion_1"
    pct_prop <- pct_int / 100

    cols_pct <- paste0("pct_", seq_len(n_cat))
    df_wide  <- tibble::tibble(
      categoria = var_label %||% "",
      n_base    = n_total
    )

    for (i in seq_len(n_cat)) {
      df_wide[[cols_pct[i]]] <- pct_prop[i]
    }

    etiquetas_grupos <- stats::setNames(as.character(tab_freq$Opciones), cols_pct)

    list(
      data             = df_wide,
      cols_porcentaje  = cols_pct,
      etiquetas_grupos = etiquetas_grupos
    )
  }

  .build_tab_barras_apiladas_multi_vars <- function(
    vars,
    data,
    survey,
    choices = NULL,
    orders_list,
    sm_vars_force,
    mostrar_todo,
    wrap_y = 50
  ) {

    listas   <- list()
    all_opts <- character(0)

    for (v in vars) {

      tab <- freq_table_spss(
        data,
        v,
        survey        = survey,
        sm_vars_force = sm_vars_force,
        orders_list   = orders_list,
        mostrar_todo  = mostrar_todo
      )

      tab <- tab |>
        dplyr::filter(Opciones != "Total") |>
        dplyr::filter(!is.na(n) & n > 0)

      if (!nrow(tab)) next

      # wrap del eje Y
      label_v <- .titulo_var_safe(v)
      if (requireNamespace("stringr", quietly = TRUE)) {
        label_v <- stringr::str_wrap(label_v, width = wrap_y)
      }

      n_total <- .extraer_N_total(tab)
      pct_int <- .pct_enteros_100(tab$n)

      listas[[v]] <- list(
        label     = label_v,
        n_total   = n_total,
        opciones  = tab$Opciones,
        pct_int   = pct_int
      )

      # union() conserva el orden de primera aparición
      all_opts <- union(all_opts, tab$Opciones)
    }

    if (!length(listas)) return(NULL)

    # ------------------------------------------------------------------
    # Ordenar las opciones según el orden formal de la lista en CHOICES
    # ------------------------------------------------------------------
    list_name_block <- NA_character_

    if ("list_name" %in% names(survey)) {
      tmp <- survey$list_name[survey$name %in% vars]
      tmp <- tmp[!is.na(tmp) & tmp != ""]
      if (length(tmp)) list_name_block <- tmp[1]
    } else if ("list_norm" %in% names(survey)) {
      tmp <- survey$list_norm[survey$name %in% vars]
      tmp <- tmp[!is.na(tmp) & tmp != ""]
      if (length(tmp)) list_name_block <- tmp[1]
    }

    if (!is.na(list_name_block)) {

      niveles_formales <- character(0)

      ## 1) PRIORIDAD: orden de la PALETA por list_name
      if (!is.null(colores_apiladas_por_listname[[list_name_block]])) {
        pal <- colores_apiladas_por_listname[[list_name_block]]

        # por si el objeto es lista con $colores
        if (is.list(pal) && !is.null(pal$colores)) {
          pal <- pal$colores
        }

        niveles_formales <- names(pal)
      }

      ## 2) Si no hay paleta o no tiene nombres, usar CHOICES como respaldo
      if (!length(niveles_formales) &&
          !is.null(choices) &&
          "list_name" %in% names(choices) &&
          "label" %in% names(choices)) {

        niveles_formales <- as.character(
          choices$label[choices$list_name == list_name_block]
        )
      }

      niveles_formales <- niveles_formales[
        !is.na(niveles_formales) & niveles_formales != ""
      ]

      if (length(niveles_formales)) {
        all_opts <- intersect(niveles_formales, all_opts)
      }
    }

    df_wide <- tibble::tibble(
      pregunta = vapply(listas, function(x) x$label, character(1)),
      n_base   = vapply(listas, function(x) x$n_total, numeric(1))
    )

    cols_pct <- paste0("pct_", seq_along(all_opts))

    for (i in seq_along(all_opts)) {
      opt_i <- all_opts[i]
      df_wide[[cols_pct[i]]] <- vapply(
        listas,
        function(x) {
          idx <- which(x$opciones == opt_i)
          if (length(idx)) x$pct_int[idx] / 100 else 0
        },
        numeric(1)
      )
    }

    etiquetas_grupos <- stats::setNames(all_opts, cols_pct)

    list(
      data             = df_wide,
      cols_porcentaje  = cols_pct,
      etiquetas_grupos = etiquetas_grupos
    )
  }

  .build_tab_dico <- function(tab_freq, var, var_label, labels_dico) {
    if (length(labels_dico) < 2) return(NULL)

    pos_lab <- labels_dico[1]
    neg_lab <- labels_dico[2]

    sub <- tab_freq |>
      dplyr::filter(.data$Opciones %in% c(pos_lab, neg_lab))

    if (nrow(sub) < 2) {
      warning(
        "En la variable '", var, "' no se encontraron ambas categorías ",
        "indicadas en `labels_dico`. Se omitirá este tratamiento dicotómico.",
        call. = FALSE
      )
      return(NULL)
    }

    n_pos <- sub$n[sub$Opciones == pos_lab][1]
    n_neg <- sub$n[sub$Opciones == neg_lab][1]
    n_vec <- c(n_pos, n_neg)

    if (!all(is.finite(n_vec)) || sum(n_vec) <= 0) return(NULL)

    # ---- Porcentajes enteros que suman 100 ----
    pct_int <- .pct_enteros_100(n_vec)

    # Valor que usará el graficador (solo el % "sí")
    pct_si  <- pct_int[1]   # entero 0–100

    indicador_val <- if (incluir_titulo_var) "" else (var_label %||% var)

    tibble::tibble(
      indicador = indicador_val,
      pct_si    = pct_si,
      n_total   = sum(n_vec)
    )
  }

  # ---------------------------------------------------------------------------
  # 4. Recorrido por secciones y variables
  # ---------------------------------------------------------------------------
  plots_list       <- list()
  titulos_list     <- list()
  log_list         <- list()
  seccion_por_plot <- character(0)
  N_texto_list     <- list()

  total_casos <- nrow(data)

  for (sec in names(SECCIONES)) {
    vars_sec <- SECCIONES[[sec]]

    if (mensajes_progreso) {
      message("Procesando sección: ", sec)
    }

    for (v in vars_sec) {

      # -----------------------------------------------------------------------
      # (NUEVO) Multi-apiladas: si la variable pertenece a un bloque especial
      # -----------------------------------------------------------------------
      if (v %in% vars_multi_all) {

        bloque_id <- names(
          Filter(function(x) v %in% x$vars, bloques_multi_apiladas)
        )[1]

        bloque_info   <- bloques_multi_apiladas[[bloque_id]]
        vars_bloque   <- bloque_info$vars
        titulo_bloque <- bloque_info$titulo %||% .titulo_var_safe(v)
        wrap_y        <- bloque_info$wrap_y %||% 50

        grosor_barras_bloque <- bloque_info$grosor_barras %||%
          estilos_barras_apiladas$grosor_barras %||% 0.7

        invertir_barras_bloque <- bloque_info$invertir_barras %||%
          estilos_barras_apiladas$invertir_barras %||% FALSE

        # Ejecutar SOLO en la primera variable del bloque
        if (v != vars_bloque[1]) {
          next
        }

        if (mensajes_progreso) {
          message(
            "   - [multi_apiladas] ",
            paste(vars_bloque, collapse = ", "),
            " → barras_apiladas_multi (bloque = ", bloque_id, ")"
          )
        }

        # 2. Construir tabla multi-var
        tab_multi <- .build_tab_barras_apiladas_multi_vars(
          vars          = vars_bloque,
          data          = data,
          survey        = survey,
          choices       = choices,
          orders_list   = orders_list,
          sm_vars_force = sm_vars_force,
          mostrar_todo  = mostrar_todo,
          wrap_y        = wrap_y
        )

        if (is.null(tab_multi)) next

        # 3. Detectar list_name del bloque (usando la primera variable)
        list_name_bloque <- NA_character_
        if ("list_name" %in% names(survey)) {
          tmp <- survey$list_name[survey$name %in% vars_bloque]
          tmp <- tmp[!is.na(tmp) & tmp != ""]
          if (length(tmp)) list_name_bloque <- tmp[1]
        } else if ("list_norm" %in% names(survey)) {
          tmp <- survey$list_norm[survey$name %in% vars_bloque]
          tmp <- tmp[!is.na(tmp) & tmp != ""]
          if (length(tmp)) list_name_bloque <- tmp[1]
        }

        # 4. Paleta y preset extra (TOP2, etc.) igual que en barras_apiladas simple
        colores_grupos <- NULL
        preset_extra   <- NULL

        if (!is.na(list_name_bloque) &&
            !is.null(colores_apiladas_por_listname[[list_name_bloque]])) {

          obj_col <- colores_apiladas_por_listname[[list_name_bloque]]

          if (is.list(obj_col)) {
            if (!is.null(obj_col$colores)) {
              colores_grupos <- obj_col$colores
            }
            if (!is.null(obj_col$preset_barra_extra)) {
              preset_extra <- obj_col$preset_barra_extra
            }
          } else {
            colores_grupos <- obj_col
          }
        }

        # 5. Flags de inversión por list_name / variable (misma lógica que apiladas simple)
        ln_inv_seg <- estilos_barras_apiladas$listnames_invertir_segmentos
        ln_inv_seg <- if (is.null(ln_inv_seg)) character(0) else ln_inv_seg

        ln_inv_ley <- estilos_barras_apiladas$listnames_invertir_leyenda
        ln_inv_ley <- if (is.null(ln_inv_ley)) character(0) else ln_inv_ley

        vars_inv_seg <- estilos_barras_apiladas$vars_invertir_segmentos
        vars_inv_seg <- if (is.null(vars_inv_seg)) character(0) else vars_inv_seg

        vars_inv_ley <- estilos_barras_apiladas$vars_invertir_leyenda
        vars_inv_ley <- if (is.null(vars_inv_ley)) character(0) else vars_inv_ley

        v_ref <- vars_bloque[1]

        invertir_segmentos_var <- (
          v_ref %in% vars_inv_seg ||
            (!is.na(list_name_bloque) && list_name_bloque %in% ln_inv_seg)
        )

        invertir_leyenda_var <- (
          v_ref %in% vars_inv_ley ||
            (!is.na(list_name_bloque) && list_name_bloque %in% ln_inv_ley)
        )

        # 6. Asegurar ORDEN consistente de segmentos y leyenda
        niveles_originales <- unname(tab_multi$etiquetas_grupos)

        niveles_plot <- niveles_originales
        if (invertir_segmentos_var) {
          niveles_plot <- rev(niveles_plot)
        }

        niveles_leyenda <- niveles_plot
        if (invertir_leyenda_var) {
          niveles_leyenda <- rev(niveles_leyenda)
        }

        if (!is.null(colores_grupos)) {
          colores_grupos <- colores_grupos[niveles_leyenda]
        }

        tab_multi$etiquetas_grupos <- stats::setNames(
          niveles_leyenda,
          tab_multi$cols_porcentaje
        )

        # 7. Limpiar claves "meta" que el graficador no conoce
        estilos_apiladas_clean <- estilos_barras_apiladas
        estilos_apiladas_clean$nota_pie         <- NULL
        estilos_apiladas_clean$nota_pie_derecha <- NULL
        estilos_apiladas_clean$pos_nota_pie     <- NULL
        estilos_apiladas_clean$listnames_invertir_segmentos <- NULL
        estilos_apiladas_clean$listnames_invertir_leyenda   <- NULL
        estilos_apiladas_clean$vars_invertir_segmentos      <- NULL
        estilos_apiladas_clean$vars_invertir_leyenda        <- NULL
        estilos_apiladas_clean$grosor_barras                <- NULL
        estilos_apiladas_clean$prefijo_barra_extra          <- NULL
        estilos_apiladas_clean$titulo_barra_extra           <- NULL
        estilos_apiladas_clean$color_barra_extra            <- NULL
        estilos_apiladas_clean$invertir_barras              <- NULL
        estilos_apiladas_clean$barra_extra_vjust <- NULL

        # N del bloque (base de personas, no suma de marcas)
        n_total_bloque <- max(tab_multi$data$n_base, na.rm = TRUE)
        N_texto <- if (is.finite(n_total_bloque) && n_total_bloque >= 0) {
          sprintf("%s", format(n_total_bloque, big.mark = ",", scientific = FALSE))
        } else {
          NULL
        }

        # 8. Construir gráfico multi-apilado
        # vjust global para barra extra, si existe en estilos
        barra_extra_vjust_global <- estilos_barras_apiladas$barra_extra_vjust %||% NULL

        args_multi_core <- list(
          data                = tab_multi$data,
          var_categoria       = "pregunta",
          var_n               = "n_base",
          cols_porcentaje     = tab_multi$cols_porcentaje,
          etiquetas_grupos    = tab_multi$etiquetas_grupos,
          escala_valor        = "proporcion_1",
          colores_grupos      = colores_grupos,
          mostrar_valores     = TRUE,
          titulo              = NULL,
          subtitulo           = NULL,
          nota_pie            = NULL,
          nota_pie_derecha    = NULL,
          pos_nota_pie        = NULL,

          mostrar_barra_extra = if (!is.null(preset_extra)) {
            TRUE
          } else {
            barra_extra == "total_n"
          },

          barra_extra_preset  = preset_extra %||% "ninguno",

          prefijo_barra_extra = if (!is.null(preset_extra)) {
            ""
          } else if (barra_extra == "total_n") {
            "N = "
          } else {
            ""
          },

          titulo_barra_extra = if (!is.null(preset_extra) || barra_extra != "total_n") {
            NULL
          } else {
            "Total"
          },

          invertir_segmentos  = invertir_segmentos_var,
          invertir_leyenda    = invertir_leyenda_var,
          invertir_barras     = invertir_barras_bloque,

          grosor_barras       = grosor_barras_bloque,

          exportar            = "rplot"
        )

        # Si hay vjust global, pasarlo al graficador
        if (!is.null(barra_extra_vjust_global)) {
          args_multi_core$barra_extra_vjust <- barra_extra_vjust_global
        }

        args_multi <- c(
          args_multi_core,
          estilos_apiladas_clean
        )

        p <- do.call(graficar_barras_apiladas, args_multi) +
          ggplot2::theme(
            axis.text.y = ggplot2::element_text(
              hjust  = 1,
              vjust  = 0.5,
              margin = ggplot2::margin(r = 6)
            ),
            plot.margin = ggplot2::margin(l = 20, r = 10, t = 5, b = 5)
          )

        idx <- length(plots_list) + 1L
        plots_list[[idx]]       <- p
        titulos_list[[idx]]     <- titulo_bloque
        N_texto_list[[idx]]     <- N_texto
        seccion_por_plot[idx]   <- sec

        log_list[[length(log_list) + 1]] <- tibble::tibble(
          seccion      = sec,
          var          = paste(vars_bloque, collapse = ", "),
          tipo_var     = "multi_apiladas",
          list_name    = list_name_bloque,
          override     = paste0("multi_apiladas=", bloque_id),
          tipo_grafico = "barras_apiladas_multi"
        )

        next
      }

      # -----------------------------------------------------------------------
      # Flujo normal
      # -----------------------------------------------------------------------
      tipo_v <- tipo_pregunta_spss(v, survey, sm_vars_force)
      if (tipo_v == "so_or_open") tipo_v <- "so"

      list_name_v <- NA_character_
      if ("list_name" %in% names(survey)) {
        tmp <- survey$list_name[survey$name == v]
        if (length(tmp)) list_name_v <- tmp[1]
      } else if ("list_norm" %in% names(survey)) {
        tmp <- survey$list_norm[survey$name == v]
        if (length(tmp)) list_name_v <- tmp[1]
      }

      override     <- NA_character_
      tipo_grafico <- NULL

      if (v %in% vars_dico) {
        override     <- "vars_dico"
        tipo_grafico <- "dico"
      } else if (v %in% vars_barras_apiladas) {
        override     <- "vars_barras_apiladas"
        tipo_grafico <- "barras_apiladas"
      } else if (v %in% vars_barras_agrupadas) {
        override     <- "vars_barras_agrupadas"
        tipo_grafico <- "barras_agrupadas"
      } else if (v %in% vars_radar) {
        override     <- "vars_radar"
        tipo_grafico <- "radar"
      } else if (!is.na(list_name_v) && list_name_v %in% listnames_dico) {
        override     <- paste0("listnames_dico=", list_name_v)
        tipo_grafico <- "dico"
      } else if (!is.na(list_name_v) && list_name_v %in% listnames_apiladas) {
        override     <- paste0("listnames_apiladas=", list_name_v)
        tipo_grafico <- "barras_apiladas"
      } else {
        if (tipo_v == "sm") {
          override     <- paste0("default_sm=", default_sm)
          tipo_grafico <- default_sm
        } else {
          override     <- paste0("default_so=", default_so)
          tipo_grafico <- default_so
        }
      }

      tab_freq <- .tab_freq_var(v)
      if (!nrow(tab_freq)) {
        log_list[[length(log_list) + 1]] <- tibble::tibble(
          seccion      = sec,
          var          = v,
          tipo_var     = tipo_v,
          list_name    = list_name_v,
          override     = override,
          tipo_grafico = NA_character_
        )
        next
      }

      # --- N de la pregunta (base por persona, no por alternativas) ---
      n_var <- .extraer_N_total(tab_freq)

      N_texto <- NULL
      if (is.finite(n_var) && n_var >= 0) {
        N_texto <- sprintf("%s", format(n_var, big.mark = ",", scientific = FALSE))
      }

      var_label    <- .titulo_var_safe(v)
      titulo_plot  <- NULL   # siempre sin título interno
      titulo_slide <- if (incluir_titulo_var) var_label else v

      # No queremos N dentro del gráfico, solo fuera en el Word
      nota_pie_plot      <- NULL
      nota_pie_derecha   <- NULL
      pos_nota_pie_plot  <- NULL

      if (mensajes_progreso) {
        message("   - ", v, " → ", tipo_grafico,
                " (list_name = ", list_name_v, ") [", override, "]")
      }

      p <- NULL
      tipo_grafico_final <- tipo_grafico

      if (tipo_grafico %in% c("barras_agrupadas", "barras_apiladas", "dico")) {

        if (tipo_grafico == "barras_agrupadas") {

          tab_agr <- .build_tab_barras_agrupadas(tab_freq, var_label)
          if (is.null(tab_agr) || !nrow(tab_agr)) next

          cols_porcentaje  <- "pct"
          etiquetas_series <- c(pct = "Porcentaje")

          colores_series <- estilos_barras_agrupadas$colores_series %||%
            c("Porcentaje" = "#1B679D")

          # ------------------------------------------------------------
          # Orientación segura (var → list_name → estilo global → default)
          # ------------------------------------------------------------
          ori_var <- NULL
          if (exists("orientacion_por_var", inherits = TRUE)) {
            ori_var <- orientacion_por_var[[v]]
          }

          ori_list <- NULL
          if (exists("orientacion_por_listname", inherits = TRUE) &&
              !is.null(list_name_v) && !is.na(list_name_v) && nzchar(list_name_v) &&
              list_name_v %in% names(orientacion_por_listname)) {
            ori_list <- orientacion_por_listname[[list_name_v]]
          }

          ori_def <- "horizontal"
          if (exists("orientacion_default", inherits = TRUE) &&
              is.character(orientacion_default) && length(orientacion_default) >= 1L) {
            ori_def <- orientacion_default[1]
          }

          ori_final <- ori_var %||%
            ori_list %||%
            estilos_barras_agrupadas$orientacion %||%
            ori_def

          # ancho máximo del eje Y según list_name
          ancho_eje_v <- estilos_barras_agrupadas$ancho_max_eje_y %||% NULL

          if (exists("ancho_eje_por_listname", inherits = TRUE) &&
              !is.null(list_name_v) && !is.na(list_name_v) && nzchar(list_name_v) &&
              list_name_v %in% names(ancho_eje_por_listname)) {

            ancho_eje_v <- ancho_eje_por_listname[[list_name_v]]
          }

          estilos_agrupadas_clean <- estilos_barras_agrupadas
          estilos_agrupadas_clean$ancho_max_eje_y <- NULL

          args_barras <- c(
            c(
              list(
                data               = tab_agr,
                var_categoria      = "categoria",
                var_n              = "n_base",
                cols_porcentaje    = cols_porcentaje,
                etiquetas_series   = etiquetas_series,
                escala_valor       = "proporcion_1",
                colores_series     = colores_series,
                mostrar_valores    = TRUE,
                titulo             = titulo_plot,
                subtitulo          = NULL,
                nota_pie           = NULL,
                nota_pie_derecha   = NULL,
                pos_nota_pie       = NULL,
                mostrar_barra_extra = barra_extra == "total_n",
                prefijo_barra_extra = if (barra_extra == "total_n") "N = " else "",
                titulo_barra_extra  = if (barra_extra == "total_n") "Total" else NULL,
                exportar            = "rplot",
                orientacion         = ori_final
              ),
              if (!is.null(ancho_eje_v)) list(ancho_max_eje_y = ancho_eje_v) else list()
            ),
            estilos_agrupadas_clean
          )

          p <- do.call(graficar_barras_agrupadas, args_barras)

          if (ori_final == "vertical") {
            p <- p +
              ggplot2::theme(
                axis.text.y  = ggplot2::element_blank(),
                axis.title.y = ggplot2::element_blank(),
                axis.ticks.y = ggplot2::element_blank(),
                axis.line.y  = ggplot2::element_blank()
              )
          }
        }

        if (tipo_grafico == "barras_apiladas") {

          tab_apil <- .build_tab_barras_apiladas(tab_freq, var_label)
          if (is.null(tab_apil)) next

          colores_grupos <- NULL
          preset_extra   <- NULL

          if (!is.na(list_name_v) &&
              !is.null(colores_apiladas_por_listname[[list_name_v]])) {

            obj_col <- colores_apiladas_por_listname[[list_name_v]]

            if (is.list(obj_col)) {
              if (!is.null(obj_col$colores)) {
                colores_grupos <- obj_col$colores
              }
              if (!is.null(obj_col$preset_barra_extra)) {
                preset_extra <- obj_col$preset_barra_extra
              }
            } else {
              colores_grupos <- obj_col
            }
          }

          ln_inv_seg <- estilos_barras_apiladas$listnames_invertir_segmentos
          ln_inv_seg <- if (is.null(ln_inv_seg)) character(0) else ln_inv_seg

          ln_inv_ley <- estilos_barras_apiladas$listnames_invertir_leyenda
          ln_inv_ley <- if (is.null(ln_inv_ley)) character(0) else ln_inv_ley

          vars_inv_seg <- estilos_barras_apiladas$vars_invertir_segmentos
          vars_inv_seg <- if (is.null(vars_inv_seg)) character(0) else vars_inv_seg

          vars_inv_ley <- estilos_barras_apiladas$vars_invertir_leyenda
          vars_inv_ley <- if (is.null(vars_inv_ley)) character(0) else vars_inv_ley

          invertir_segmentos_var <- (
            v %in% vars_inv_seg ||
              (!is.na(list_name_v) && list_name_v %in% ln_inv_seg)
          )

          invertir_leyenda_var <- (
            v %in% vars_inv_ley ||
              (!is.na(list_name_v) && list_name_v %in% ln_inv_ley)
          )

          estilos_apiladas_clean <- estilos_barras_apiladas
          estilos_apiladas_clean$listnames_invertir_segmentos <- NULL
          estilos_apiladas_clean$listnames_invertir_leyenda   <- NULL
          estilos_apiladas_clean$vars_invertir_segmentos      <- NULL
          estilos_apiladas_clean$vars_invertir_leyenda        <- NULL
          estilos_apiladas_clean$barra_extra_vjust <- NULL


          args_extra <- list()
          if (!is.null(colores_grupos)) {
            args_extra$colores_grupos <- colores_grupos
          }

          # vjust global para barra extra, si existe en estilos
          barra_extra_vjust_global <- estilos_barras_apiladas$barra_extra_vjust %||% NULL

          args_core <- list(
            data                = tab_apil$data,
            var_categoria       = "categoria",
            var_n               = "n_base",
            cols_porcentaje     = tab_apil$cols_porcentaje,
            etiquetas_grupos    = tab_apil$etiquetas_grupos,
            escala_valor        = "proporcion_1",
            mostrar_valores     = TRUE,
            titulo              = titulo_plot,
            subtitulo           = NULL,
            nota_pie            = nota_pie_plot,
            nota_pie_derecha    = nota_pie_derecha,
            pos_nota_pie        = pos_nota_pie_plot,

            mostrar_barra_extra = if (!is.null(preset_extra)) TRUE else (barra_extra == "total_n"),
            barra_extra_preset  = preset_extra,

            prefijo_barra_extra = if (!is.null(preset_extra)) {
              ""
            } else if (barra_extra == "total_n") {
              "N = "
            } else {
              ""
            },

            titulo_barra_extra = if (!is.null(preset_extra) || barra_extra != "total_n") {
              NULL
            } else {
              "Total"
            },

            color_barra_extra = if (!is.null(preset_extra)) {
              NULL
            } else {
              "#092147"
            },

            exportar           = "rplot",
            invertir_segmentos = invertir_segmentos_var,
            invertir_leyenda   = invertir_leyenda_var
          )

          # Si hay vjust global, pasarlo
          if (!is.null(barra_extra_vjust_global)) {
            args_core$barra_extra_vjust <- barra_extra_vjust_global
          }

          args_apiladas <- c(
            args_core,
            args_extra,
            estilos_apiladas_clean
          )

          if (!is.null(names(args_apiladas))) {
            args_apiladas <- args_apiladas[!duplicated(names(args_apiladas))]
          }

          p <- do.call(graficar_barras_apiladas, args_apiladas)

          p <- p +
            ggplot2::theme(
              axis.text.y  = ggplot2::element_blank(),
              axis.title.y = ggplot2::element_blank()
            )
        }

        if (tipo_grafico == "dico") {

          labels_dico <- NULL
          if (!is.null(dico_labels_por_var[[v]])) {
            labels_dico <- dico_labels_por_var[[v]]
          } else if (!is.na(list_name_v) &&
                     !is.null(dico_labels_por_listname[[list_name_v]])) {
            labels_dico <- dico_labels_por_listname[[list_name_v]]
          }

          if (is.null(labels_dico) || length(labels_dico) < 2) {
            warning(
              "En la variable '", v, "' no se encontraron etiquetas dicotómicas ",
              "en `dico_labels_por_var` ni `dico_labels_por_listname`. ",
              "Se usará barras agrupadas.",
              call. = FALSE
            )
            tipo_grafico_final <- "barras_agrupadas"

            tab_agr <- .build_tab_barras_agrupadas(tab_freq, var_label)
            if (is.null(tab_agr) || !nrow(tab_agr)) next

            cols_porcentaje  <- "pct"
            etiquetas_series <- c(pct = "Porcentaje")

            colores_series <- estilos_barras_agrupadas$colores_series %||%
              c("Porcentaje" = "#004B8D")

            args_barras <- c(
              list(
                data               = tab_agr,
                var_categoria      = "categoria",
                var_n              = "n_base",
                cols_porcentaje    = cols_porcentaje,
                etiquetas_series   = etiquetas_series,
                escala_valor       = "proporcion_1",
                colores_series     = colores_series,
                mostrar_valores    = TRUE,
                titulo             = titulo_plot,
                subtitulo          = NULL,
                nota_pie            = NULL,
                nota_pie_derecha    = NULL,
                pos_nota_pie        = NULL,
                mostrar_barra_extra = barra_extra == "total_n",
                prefijo_barra_extra = if (barra_extra == "total_n") "N = " else "",
                titulo_barra_extra  = if (barra_extra == "total_n") "Total" else NULL,
                exportar            = "rplot"
              ),
              estilos_barras_agrupadas
            )

            p <- do.call(graficar_barras_agrupadas, args_barras)

          } else {

            tab_dico <- .build_tab_dico(tab_freq, v, var_label, labels_dico)
            if (is.null(tab_dico) || !nrow(tab_dico)) next

            args_dico <- c(
              list(
                data              = tab_dico,
                var_indicador     = "indicador",
                var_porcentaje_si = "pct_si",
                var_n             = "n_total",
                escala_valor      = "proporcion_100",
                etiqueta_si       = labels_dico[1],
                etiqueta_no       = labels_dico[2],
                titulo            = titulo_plot,
                subtitulo         = NULL,
                nota_pie          = N_texto,
                incluir_n_en_titulo = FALSE,
                exportar          = "rplot"
              ),
              estilos_dico
            )

            p <- do.call(graficar_dico, args_dico)
          }
        }

      } else if (tipo_grafico == "radar") {

        warning(
          "Tipo de gráfico 'radar' señalado para la variable '", v,
          "', pero el constructor genérico aún no está implementado. ",
          "La variable se omitirá en este reporte.",
          call. = FALSE
        )
        next

      } else {
        warning(
          "Tipo de gráfico '", tipo_grafico, "' no reconocido para la variable '",
          v, "'. Se omitirá en este reporte.",
          call. = FALSE
        )
        next
      }

      log_list[[length(log_list) + 1]] <- tibble::tibble(
        seccion      = sec,
        var          = v,
        tipo_var     = tipo_v,
        list_name    = list_name_v,
        override     = override,
        tipo_grafico = tipo_grafico_final
      )

      idx <- length(plots_list) + 1L
      plots_list[[idx]]       <- p
      titulos_list[[idx]]     <- titulo_slide
      N_texto_list[[idx]]     <- N_texto
      seccion_por_plot[idx]   <- sec
    }
  }

  log_decisiones <- dplyr::bind_rows(log_list)

  # ---------------------------------------------------------------------------
  # 5. Construir DOCX
  # ---------------------------------------------------------------------------
  if (!solo_lista) {
    if (!requireNamespace("officer", quietly = TRUE)) {
      stop(
        "Para exportar a Word se requiere el paquete 'officer'.",
        call. = FALSE
      )
    }

    pulso_azul <- "#1B679D"

    # -------------------------------------------------------------------------
    # 5.1. Leer plantilla Word
    # -------------------------------------------------------------------------
    if (is.null(template_docx)) {
      template_interno <- system.file(
        "plantillas/plantilla_pulso.docx",
        package = "prosecnur"
      )

      if (nzchar(template_interno) && file.exists(template_interno)) {
        if (mensajes_progreso) {
          message("Usando plantilla Word interna: ", template_interno)
        }
        doc <- officer::read_docx(path = template_interno)
      } else {
        if (mensajes_progreso) {
          message(
            "No se encontró 'plantillas/plantilla_pulso.docx' dentro del paquete. ",
            "Se usará la plantilla por defecto de Word."
          )
        }
        doc <- officer::read_docx()
      }

    } else {
      if (!file.exists(template_docx)) {
        stop(
          "No se encontró el archivo de plantilla especificado en `template_docx`: ",
          template_docx,
          call. = FALSE
        )
      }
      if (mensajes_progreso) {
        message("Usando plantilla Word definida por el usuario: ", template_docx)
      }
      doc <- officer::read_docx(path = template_docx)
    }

    # -------------------------------------------------------------------------
    # 5.2. Primera página: índice de gráficos por sección
    # -------------------------------------------------------------------------
    if (length(plots_list)) {

      indice_titulo <- officer::fpar(
        officer::ftext(
          "ÍNDICE DE GRÁFICOS",
          prop = officer::fp_text(
            bold        = TRUE,
            italic      = FALSE,
            color       = pulso_azul,
            font.size   = 12,
            font.family = "Arial"
          )
        )
      )
      doc <- officer::body_add_fpar(doc, indice_titulo, style = "Normal")
      doc <- officer::body_add_par(doc, "", style = "Normal")

      if (length(seccion_por_plot) == length(plots_list)) {

        secciones_unicas <- unique(seccion_por_plot)

        for (sec in secciones_unicas) {
          if (is.na(sec) || !nzchar(sec)) next

          sec_fpar <- officer::fpar(
            officer::ftext(
              sec,
              prop = officer::fp_text(
                bold        = TRUE,
                italic      = FALSE,
                color       = "black",
                font.size   = 11,
                font.family = "Arial"
              )
            )
          )
          doc <- officer::body_add_fpar(doc, sec_fpar, style = "Normal")

          idx_sec <- which(seccion_por_plot == sec)

          for (i in idx_sec) {
            titulo_i <- titulos_list[[i]] %||% ""

            entrada_indice <- officer::fpar(
              officer::ftext(
                sprintf("Gráfico Nº %d. ", i),
                prop = officer::fp_text(
                  bold        = TRUE,
                  italic      = FALSE,
                  color       = pulso_azul,
                  font.size   = 10,
                  font.family = "Arial"
                )
              ),
              officer::ftext(
                titulo_i,
                prop = officer::fp_text(
                  bold        = FALSE,
                  italic      = TRUE,
                  color       = "black",
                  font.size   = 10,
                  font.family = "Arial"
                )
              )
            )

            doc <- officer::body_add_fpar(doc, entrada_indice, style = "Normal")
          }

          doc <- officer::body_add_par(doc, "", style = "Normal")
        }

      } else {
        for (i in seq_along(plots_list)) {
          titulo_i <- titulos_list[[i]] %||% ""

          entrada_indice <- officer::fpar(
            officer::ftext(
              sprintf("Gráfico Nº %d. ", i),
              prop = officer::fp_text(
                bold        = TRUE,
                italic      = FALSE,
                color       = pulso_azul,
                font.size   = 10,
                font.family = "Arial"
              )
            ),
            officer::ftext(
              titulo_i,
              prop = officer::fp_text(
                bold        = FALSE,
                italic      = TRUE,
                color       = "black",
                font.size   = 10,
                font.family = "Arial"
              )
            )
          )

          doc <- officer::body_add_fpar(doc, entrada_indice, style = "Normal")
        }
      }

      doc <- officer::body_add_break(doc)

      # -----------------------------------------------------------------------
      # 5.3. Gráficas: título externo + gráfico + pie de página
      # -----------------------------------------------------------------------
      for (i in seq_along(plots_list)) {

        p        <- plots_list[[i]]
        titulo_i <- titulos_list[[i]] %||% ""
        N_i      <- N_texto_list[[i]] %||% NA

        # ---------------- TÍTULO CENTRADO ----------------
        titulo_full <- sprintf("Gráfico Nº %d: %s", i, titulo_i)

        titulo_word <- officer::fpar(
          officer::ftext(
            titulo_full,
            prop = officer::fp_text(
              bold        = TRUE,
              italic      = TRUE,
              color       = pulso_azul,
              font.size   = 11,
              font.family = "Arial"
            )
          ),
          fp_p = officer::fp_par(text.align = "center")
        )

        doc <- officer::body_add_fpar(doc, titulo_word, style = "Normal")

        # ---------------- GRÁFICO ----------------
        width_word <- 17 / 2.54

        alto_sugerido <- attr(p, "alto_word_sugerido", exact = TRUE)
        height_word <- if (!is.null(alto_sugerido) && is.finite(alto_sugerido)) {
          alto_sugerido
        } else {
          3.5
        }

        doc <- officer::body_add_gg(
          doc,
          value  = p,
          width  = width_word,
          height = height_word
        )

        # ---------------- PIE CENTRADO ----------------
        pie_full <- NULL

        if (!is.null(fuente) && nzchar(fuente) && !is.na(N_i)) {
          pie_full <- paste0("N = ", N_i, fuente)
        } else if (!is.null(fuente) && nzchar(fuente)) {
          pie_full <- fuente
        } else if (!is.na(N_i)) {
          pie_full <- paste0("N = ", N_i)
        }

        if (!is.null(pie_full)) {
          pie_word <- officer::fpar(
            officer::ftext(
              pie_full,
              prop = officer::fp_text(
                color       = pulso_azul,
                font.size   = 7,
                bold        = TRUE,
                italic      = TRUE,
                font.family = "Arial"
              )
            ),
            fp_p = officer::fp_par(text.align = "center")
          )

          doc <- officer::body_add_fpar(doc, pie_word, style = "Normal")
        }

        doc <- officer::body_add_par(doc, "", style = "Normal")
      }
    }

    print(doc, target = path_word)
    if (mensajes_progreso) {
      message("Word generado en: ", normalizePath(path_word, winslash = "/"))
    }
  }

  invisible(list(
    plots          = plots_list,
    log_decisiones = log_decisiones
  ))
}
