#' Crear un reporte en PowerPoint a partir de resultados de encuesta
#'
#' Genera un archivo **.pptx** con una diapositiva por variable (y, de forma opcional,
#' algunas diapositivas con **dos gráficos**). El flujo general es:
#' (1) identificar las variables a reportar (desde `secciones` o desde el instrumento),
#' (2) calcular tablas de frecuencias,
#' (3) elegir el tipo de gráfico según reglas y prioridades (por variable, por `list_name`
#' o por defecto), y (4) armar el PowerPoint usando una plantilla.
#'
#' El tipo de gráfico se selecciona con un orden de preferencia:
#' primero sobrescrituras explícitas (`vars_*` y `listnames_*`), y si no hay coincidencias,
#' se aplican los valores por defecto (`default_so` y `default_sm`) según el tipo de pregunta.
#' Además, se admite un modo especial para construir **bloques de varias preguntas** en una
#' sola visualización apilada (`bloques_multi_apiladas`).
#'
#' El PowerPoint puede generarse desde una plantilla propia (`template_pptx`) o, si está
#' disponible, desde la plantilla incluida en el paquete. Si `solo_lista = TRUE`, la función
#' no escribe el archivo y únicamente devuelve los gráficos y el registro de decisiones.
#'
#' @param data `data.frame` o `tibble` con las variables (o sus dummies) a reportar.
#' @param instrumento Objeto de instrumento con al menos `survey` (y opcionalmente `choices`
#'   y `orders_list`). Si es `NULL`, se busca el atributo `instrumento_reporte` en `data`.
#' @param secciones Lista nombrada donde cada elemento contiene un vector de nombres de variables
#'   a incluir por sección. Si es `NULL`, se intenta inferir desde `survey$section` o `survey$seccion`.
#' @param path_ppt Ruta del archivo `.pptx` de salida.
#' @param fuente Texto opcional que se agrega como fuente/base en las diapositivas.
#' @param sm_vars_force Vector de variables que deben tratarse como selección múltiple (si aplica
#'   en la lógica de clasificación).
#' @param mostrar_todo Si `TRUE`, conserva opciones con frecuencia cero al construir tablas (si la
#'   tabla de frecuencias lo permite).
#' @param solo_lista Si `TRUE`, no se genera el PowerPoint. Se devuelven gráficos y el log.
#' @param incluir_titulo_var Si `TRUE`, usa el título de la pregunta como título de la diapositiva.
#' @param mensajes_progreso Si `TRUE`, imprime mensajes de avance durante el proceso.
#'
#' @param bloques_multi_apiladas Lista nombrada de bloques especiales para apiladas múltiples.
#'   Cada bloque debe incluir al menos `vars` (vector de variables) y puede incluir `titulo`
#'   y `wrap_y`.
#' @param pares_diapositiva Lista para definir pares de gráficos en una misma diapositiva.
#'   Cada elemento debe incluir `vars` (vector de longitud 2) y opcionalmente `titulo`.
#'
#' @param vars_dico,vars_barras_apiladas,vars_barras_agrupadas,vars_radar Vectores de nombres de
#'   variables que fuerzan el tipo de gráfico por variable.
#' @param listnames_dico,listnames_apiladas Vectores de `list_name` que fuerzan el tipo de gráfico
#'   por lista.
#'
#' @param dico_labels_por_var Lista con etiquetas dicotómicas por variable. Cada elemento debe
#'   ser un vector de dos etiquetas (por ejemplo, `c("Sí", "No")`).
#' @param dico_labels_por_listname Lista con etiquetas dicotómicas por `list_name`, con el mismo
#'   formato que `dico_labels_por_var`.
#'
#' @param colores_apiladas_por_listname Lista con definiciones de colores para barras apiladas
#'   por `list_name`. Puede incluir un vector nombrado de colores y, opcionalmente, parámetros
#'   extra asociados al estilo de barra adicional.
#'
#' @param default_so Tipo de gráfico por defecto para variables de opción única (por ejemplo,
#'   `"barras_apiladas"` o `"barras_agrupadas"`).
#' @param default_sm Tipo de gráfico por defecto para variables de selección múltiple.
#' @param barra_extra Define el tratamiento de una anotación adicional en gráficos de barras:
#'   `"ninguna"` o `"total_n"`.
#'
#' @param estilos_barras_agrupadas Lista de parámetros de estilo que se pasan al graficador de
#'   barras agrupadas.
#' @param estilos_barras_apiladas Lista de parámetros de estilo que se pasan al graficador de
#'   barras apiladas.
#' @param estilos_dico Lista de parámetros de estilo que se pasan al graficador dicotómico.
#'
#' @param template_pptx Ruta a una plantilla `.pptx` para definir layouts y estilos. Si es `NULL`,
#'   se intenta usar una plantilla interna y, en su defecto, la plantilla por defecto de PowerPoint.
#'
#' @param titulo_portada,subtitulo_portada,fecha_portada Textos opcionales para incluir una diapositiva
#'   de portada si la plantilla cuenta con un layout compatible.
#' @param mostrar_resumen_n Si `TRUE`, considera incluir un resumen de N (según el layout disponible).
#'
#' @param debug_ph_bordes Si `TRUE`, activa bordes de depuración en los gráficos que usan canvas
#'   (útil para revisar alineaciones internas).
#'
#' @return Devuelve invisiblemente una lista con:
#' \describe{
#'   \item{plots}{Lista de gráficos generados (en el mismo orden de salida).}
#'   \item{log_decisiones}{Tabla con las decisiones tomadas por variable (tipo, override, `list_name`, etc.).}
#' }
#' Si `solo_lista = TRUE`, el objeto retornado contiene los gráficos y el log, pero no se escribe el PPT.
#'
#' @examples
#' \dontrun{
#' out <- reporte_ppt(
#'   data = rp_data,
#'   instrumento = rp_inst,
#'   secciones = list("Sección 1" = c("p1", "p2")),
#'   path_ppt = "reporte.pptx",
#'   fuente = "Pulso PUCP"
#' )
#' }
#'
#' @family reporte
#' @export
reporte_ppt <- function(
    data,
    instrumento      = NULL,
    secciones        = NULL,
    path_ppt         = "reporte_ppt.pptx",
    fuente           = NULL,
    sm_vars_force    = NULL,
    mostrar_todo     = FALSE,
    solo_lista       = FALSE,
    incluir_titulo_var = TRUE,
    mensajes_progreso  = TRUE,

    # Bloques de varias vars apiladas
    bloques_multi_apiladas = NULL,

    # Pares de gráficos en una misma diapositiva
    pares_diapositiva = NULL,

    # Sobrescritura por variable
    vars_dico             = NULL,
    vars_barras_apiladas  = NULL,
    vars_barras_agrupadas = NULL,
    vars_radar            = NULL,

    # Overrides por list_name
    listnames_dico        = NULL,
    listnames_apiladas    = NULL,
    listnames_barras_agrupadas  = NULL,

    # Etiquetas explícitas para dicotómicas
    dico_labels_por_var      = list(),
    dico_labels_por_listname = list(),

    # Colores para barras apiladas por list_name
    colores_apiladas_por_listname = list(),

    # Defaults por tipo de pregunta
    default_so = c("barras_apiladas", "barras_agrupadas"),
    default_sm = c("barras_agrupadas", "barras_apiladas"),

    # Barra extra en barras apiladas/agrupadas
    barra_extra = c("ninguna", "total_n"),

    # Estilos por tipo de gráfico
    estilos_barras_agrupadas = list(),
    estilos_barras_apiladas  = list(),
    estilos_dico             = list(),

    # Plantilla PPT
    template_pptx = NULL,

    # Texto de portada
    titulo_portada    = NULL,
    subtitulo_portada = NULL,
    fecha_portada     = NULL,

    # Resumen de N en bloque derecho
    mostrar_resumen_n = TRUE,

    # ==========================
    # DEBUG: bordes morados de placeholders internos (canvas)
    # ==========================
    debug_ph_bordes = FALSE
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
  listnames_barras_agrupadas <- listnames_barras_agrupadas %||% character(0)

  if (!is.data.frame(data)) {
    stop("`data` debe ser un data.frame o tibble.", call. = FALSE)
  }

  # ---------------------------------------------------------------------------
  # Helpers internos (robustos a cambios de firmas)
  # ---------------------------------------------------------------------------

  .has_var_or_dummies <- function(df, v) {
    if (v %in% names(df)) return(TRUE)
    # dummies típicos: var/cod o var.cod
    pat1 <- paste0("^", gsub("([\\W])", "\\\\\\1", v), "/")
    pat2 <- paste0("^", gsub("([\\W])", "\\\\\\1", v), "\\.")
    any(grepl(pat1, names(df))) || any(grepl(pat2, names(df)))
  }

  .safe_named_args <- function(fun, args) {
    if (!length(args)) return(args)
    fn <- tryCatch(match.fun(fun), error = function(e) NULL)
    if (is.null(fn)) return(args)
    fml <- names(formals(fn))
    if (is.null(fml)) return(list())
    args[names(args) %in% fml]
  }

  .safe_call <- function(fun, args) {
    args2 <- .safe_named_args(fun, args)
    # silenciar warnings conocidos de dependencias (aes_string, lifecycle, etc.)
    suppressWarnings(do.call(fun, args2))
  }

  # Enfoque post-canvas: por defecto, si el graficador soporta usar_canvas,
  # se fuerza TRUE para PPT. Si debug está activo, también.
  .force_canvas_style <- function(estilos) {
    estilos <- estilos %||% list()
    estilos$usar_canvas <- estilos$usar_canvas %||% TRUE
    if (isTRUE(debug_ph_bordes)) estilos$usar_canvas <- TRUE
    estilos
  }

  estilos_barras_apiladas  <- .force_canvas_style(estilos_barras_apiladas)
  estilos_barras_agrupadas <- .force_canvas_style(estilos_barras_agrupadas)
  estilos_dico             <- .force_canvas_style(estilos_dico)

  # ---------------------------------------------------------------------------
  # 1. Instrumento y survey / choices / orders_list
  # ---------------------------------------------------------------------------
  if (is.null(instrumento)) {
    instrumento <- attr(data, "instrumento_reporte", exact = TRUE)
    if (is.null(instrumento)) {
      stop("No se proporcionó `instrumento` y `data` no tiene atributo `instrumento_reporte`.", call. = FALSE)
    }
  }

  survey  <- instrumento$survey
  choices <- instrumento$choices %||% NULL

  if (is.null(survey) || !all(c("name", "label") %in% names(survey))) {
    stop("El `instrumento` no contiene un `survey` con columnas `name` y `label`.", call. = FALSE)
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
      stop("No se especificaron `secciones` y el `survey` no tiene columna `section` ni `seccion`.", call. = FALSE)
    }

    secciones_df <- survey |>
      dplyr::filter(
        !is.na(.data[[seccion_col]]),
        !is.na(.data$name)
      ) |>
      dplyr::select(seccion = !!rlang::sym(seccion_col), name)

    if (nrow(secciones_df) == 0) {
      stop("No se pudieron inferir secciones desde `survey$", seccion_col, "`.", call. = FALSE)
    }

    secciones <- split(secciones_df$name, secciones_df$seccion)
  }

  SECCIONES <- lapply(secciones, function(vars) {
    vars[vapply(vars, function(v) .has_var_or_dummies(data, v), logical(1))]
  })
  SECCIONES <- SECCIONES[vapply(SECCIONES, length, integer(1)) > 0L]

  if (length(SECCIONES) == 0L) {
    stop("Después de filtrar por presencia en `data` (variable o dummies), ninguna sección tiene variables válidas.", call. = FALSE)
  }

  # ---------------------------------------------------------------------------
  # Precomputar variables incluidas en bloques multi-apilados
  # ---------------------------------------------------------------------------
  if (!is.null(bloques_multi_apiladas) && length(bloques_multi_apiladas) > 0) {
    vars_multi_all <- unique(unlist(lapply(bloques_multi_apiladas, `[[`, "vars")))
  } else {
    vars_multi_all <- character(0)
  }

  # ---------------------------------------------------------------------------
  # 3. Helpers de título / tablas (dependen de funciones del paquete)
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

    # Extraer N correcto desde la fila "Total" (si existe)
    N_total <- NA_real_
    if ("Opciones" %in% names(tab) && "n" %in% names(tab)) {
      idx_tot <- which(tab$Opciones == "Total")
      if (length(idx_tot)) {
        N_total <- suppressWarnings(as.numeric(tab$n[idx_tot[1]]))
      }
    }

    tab2 <- tab |>
      dplyr::filter(.data$Opciones != "Total") |>
      dplyr::filter(!is.na(.data$n) & .data$n > 0)

    if (is.finite(N_total)) {
      attr(tab2, "N_total") <- N_total
    }

    tab2
  }

  .N_total_from_tab <- function(tab, total_casos = NULL) {
    N <- attr(tab, "N_total", exact = TRUE)
    if (is.null(N) || !is.finite(N)) {
      N <- sum(tab$n, na.rm = TRUE)
    }
    if (!is.null(total_casos) &&
        is.finite(total_casos) && total_casos > 0 &&
        is.finite(N) && N > total_casos) {
      N <- total_casos
    }
    N
  }

  # Reparte porcentajes enteros que suman 100 (para apiladas / dico)
  .pct_enteros_100 <- function(n) {
    n <- as.numeric(n)
    if (!length(n) || all(is.na(n))) return(numeric(0))
    n[is.na(n)] <- 0
    total <- sum(n)

    if (!is.finite(total) || total <= 0) return(rep(0L, length(n)))

    raw_pct   <- n / total * 100
    floor_pct <- floor(raw_pct)
    resid <- as.integer(round(100 - sum(floor_pct)))
    frac <- raw_pct - floor_pct

    if (resid > 0) {
      ord <- order(frac, decreasing = TRUE, na.last = TRUE)
      idx <- head(ord, resid)
      floor_pct[idx] <- floor_pct[idx] + 1L
    } else if (resid < 0) {
      resid_neg <- abs(resid)
      ord <- order(frac, decreasing = FALSE, na.last = TRUE)
      idx <- head(ord, resid_neg)
      floor_pct[idx] <- pmax(0L, floor_pct[idx] - 1L)
    }
    floor_pct
  }

  .build_tab_barras_agrupadas <- function(tab_freq, var_label) {
    if (!nrow(tab_freq)) return(NULL)

    n_total <- .N_total_from_tab(tab_freq, total_casos)

    pct_raw <- tab_freq$pct
    if (all(is.na(pct_raw))) return(NULL)

    max_pct <- max(pct_raw, na.rm = TRUE)
    if (is.finite(max_pct) && max_pct <= 1 + 1e-8) {
      pct_0_100 <- pct_raw * 100
    } else {
      pct_0_100 <- pct_raw
    }

    pct_int  <- round(pct_0_100)
    pct_prop <- pct_int / 100

    tibble::tibble(
      categoria = tab_freq$Opciones,
      n_base    = n_total,
      pct       = pct_prop
    )
  }

  .build_tab_barras_apiladas <- function(tab_freq, var_label) {
    if (!nrow(tab_freq)) return(NULL)

    n_total <- .N_total_from_tab(tab_freq, total_casos)
    n_cat   <- nrow(tab_freq)

    pct_int <- .pct_enteros_100(tab_freq$n)
    cols_pct <- paste0("pct_", seq_len(n_cat))

    df_wide <- tibble::tibble(
      categoria = var_label %||% "",
      n_base    = n_total
    )
    for (i in seq_len(n_cat)) {
      df_wide[[cols_pct[i]]] <- pct_int[i] / 100
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
    colores_apiladas_por_listname = list(),
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
      if (!nrow(tab)) next

      N_total_v <- NA_real_
      if ("Opciones" %in% names(tab) && "n" %in% names(tab)) {
        idx_tot <- which(tab$Opciones == "Total")
        if (length(idx_tot)) {
          N_total_v <- suppressWarnings(as.numeric(tab$n[idx_tot[1]]))
        }
      }

      tab <- tab |>
        dplyr::filter(.data$Opciones != "Total") |>
        dplyr::filter(!is.na(.data$n) & .data$n > 0)
      if (!nrow(tab)) next

      label_v <- .titulo_var_safe(v)
      if (requireNamespace("stringr", quietly = TRUE)) {
        label_v <- stringr::str_wrap(label_v, width = wrap_y)
      }

      n_total <- if (is.finite(N_total_v)) N_total_v else sum(tab$n, na.rm = TRUE)
      pct_int <- .pct_enteros_100(tab$n)

      listas[[v]] <- list(
        label     = label_v,
        n_total   = n_total,
        opciones  = tab$Opciones,
        pct_int   = pct_int
      )

      all_opts <- union(all_opts, tab$Opciones)
    }

    if (!length(listas)) return(NULL)

    # Orden formal de opciones (prioridad: paleta nombrada > choices)
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

      if (!is.null(colores_apiladas_por_listname[[list_name_block]])) {
        pal <- colores_apiladas_por_listname[[list_name_block]]
        if (is.list(pal) && !is.null(pal$colores)) pal <- pal$colores
        niveles_formales <- names(pal)
      }

      if (!length(niveles_formales) &&
          !is.null(choices) &&
          "list_name" %in% names(choices) &&
          "label" %in% names(choices)) {

        niveles_formales <- as.character(
          choices$label[choices$list_name == list_name_block]
        )
      }

      niveles_formales <- niveles_formales[!is.na(niveles_formales) & niveles_formales != ""]
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
      etiquetas_grupos = etiquetas_grupos,
      list_name_block  = list_name_block
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
        "En la variable '", var, "' no se encontraron ambas categorías indicadas en `labels_dico`.",
        call. = FALSE
      )
      return(NULL)
    }

    n_pos <- sub$n[sub$Opciones == pos_lab][1]
    n_neg <- sub$n[sub$Opciones == neg_lab][1]
    denom <- n_pos + n_neg
    if (!is.finite(denom) || denom <= 0) return(NULL)

    pct_pair   <- .pct_enteros_100(c(n_pos, n_neg))
    pct_si_int <- pct_pair[1]

    indicador_val <- if (incluir_titulo_var) "" else (var_label %||% var)

    tibble::tibble(
      indicador = indicador_val,
      pct_si    = pct_si_int,
      n_total   = denom
    )
  }

  # ---------------------------------------------------------------------------
  # 4. Recorrido por secciones y variables
  # ---------------------------------------------------------------------------
  plots_list       <- list()
  titulos_list     <- list()
  resumenN_list    <- list()
  log_list         <- list()
  seccion_por_plot <- character(0)
  vars_por_plot    <- character(0)

  total_casos <- nrow(data)

  for (sec in names(SECCIONES)) {
    vars_sec <- SECCIONES[[sec]]

    if (mensajes_progreso) message("Procesando sección: ", sec)

    for (v in vars_sec) {

      # -----------------------------------------------------------------------
      # (NUEVO) Multi-apiladas: si la variable pertenece a un bloque especial
      # -----------------------------------------------------------------------
      if (v %in% vars_multi_all) {

        bloque_id <- names(Filter(function(x) v %in% x$vars, bloques_multi_apiladas))[1]
        bloque_info   <- bloques_multi_apiladas[[bloque_id]]
        vars_bloque   <- bloque_info$vars
        titulo_bloque <- bloque_info$titulo %||% .titulo_var_safe(v)
        wrap_y        <- bloque_info$wrap_y %||% 50

        if (v != vars_bloque[1]) next

        if (mensajes_progreso) {
          message("   - [multi_apiladas] ", paste(vars_bloque, collapse = ", "),
                  " → barras_apiladas_multi (bloque = ", bloque_id, ")")
        }

        tab_multi <- .build_tab_barras_apiladas_multi_vars(
          vars          = vars_bloque,
          data          = data,
          survey        = survey,
          choices       = choices,
          orders_list   = orders_list,
          sm_vars_force = sm_vars_force,
          mostrar_todo  = mostrar_todo,
          colores_apiladas_por_listname = colores_apiladas_por_listname,
          wrap_y        = wrap_y
        )
        if (is.null(tab_multi)) next

        list_name_bloque <- tab_multi$list_name_block %||% NA_character_

        colores_grupos <- NULL
        preset_extra   <- NULL
        if (!is.na(list_name_bloque) &&
            !is.null(colores_apiladas_por_listname[[list_name_bloque]])) {

          obj_col <- colores_apiladas_por_listname[[list_name_bloque]]
          if (is.list(obj_col)) {
            if (!is.null(obj_col$colores)) colores_grupos <- obj_col$colores
            if (!is.null(obj_col$preset_barra_extra)) preset_extra <- obj_col$preset_barra_extra
          } else {
            colores_grupos <- obj_col
          }
        }

        # Centro leyenda (var/listname) si existen objetos en el entorno
        centro_cowplot <- NULL
        if (exists("centro_leyenda_por_bloque", inherits = TRUE) &&
            bloque_id %in% names(centro_leyenda_por_bloque)) {
          centro_cowplot <- centro_leyenda_por_bloque[[bloque_id]]
        }
        if (is.null(centro_cowplot) &&
            exists("centro_leyenda_por_listname", inherits = TRUE) &&
            !is.na(list_name_bloque) && nzchar(list_name_bloque) &&
            list_name_bloque %in% names(centro_leyenda_por_listname)) {
          centro_cowplot <- centro_leyenda_por_listname[[list_name_bloque]]
        }

        # invertir por reglas (si existen)
        ln_inv_seg <- estilos_barras_apiladas$listnames_invertir_segmentos %||% character(0)
        ln_inv_ley <- estilos_barras_apiladas$listnames_invertir_leyenda   %||% character(0)
        vars_inv_seg <- estilos_barras_apiladas$vars_invertir_segmentos %||% character(0)
        vars_inv_ley <- estilos_barras_apiladas$vars_invertir_leyenda   %||% character(0)

        v_ref <- vars_bloque[1]
        invertir_segmentos_var <- (v_ref %in% vars_inv_seg) ||
          (!is.na(list_name_bloque) && list_name_bloque %in% ln_inv_seg)

        invertir_leyenda_var <- (v_ref %in% vars_inv_ley) ||
          (!is.na(list_name_bloque) && list_name_bloque %in% ln_inv_ley)

        estilos_apiladas_clean <- estilos_barras_apiladas
        estilos_apiladas_clean$listnames_invertir_segmentos <- NULL
        estilos_apiladas_clean$listnames_invertir_leyenda   <- NULL
        estilos_apiladas_clean$vars_invertir_segmentos      <- NULL
        estilos_apiladas_clean$vars_invertir_leyenda        <- NULL
        estilos_apiladas_clean$prefijo_barra_extra <- NULL
        estilos_apiladas_clean$mostrar_barra_extra <- NULL
        estilos_apiladas_clean$barra_extra_preset  <- NULL
        estilos_apiladas_clean$titulo_barra_extra  <- NULL

        args_multi <- c(
          list(
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

            mostrar_barra_extra = if (!is.null(preset_extra)) TRUE else (barra_extra == "total_n"),
            barra_extra_preset  = preset_extra %||% "ninguno",
            prefijo_barra_extra = if (!is.null(preset_extra)) "" else if (barra_extra == "total_n") "N = " else "",
            titulo_barra_extra  = NULL,

            invertir_segmentos  = invertir_segmentos_var,
            invertir_leyenda    = invertir_leyenda_var,

            debug_ph_bordes     = isTRUE(debug_ph_bordes),

            exportar            = "rplot"
          ),
          if (!is.null(centro_cowplot)) list(centro_cowplot = centro_cowplot) else list(),
          estilos_apiladas_clean
        )

        p <- .safe_call(graficar_barras_apiladas, args_multi)

        n_total_bloque <- max(tab_multi$data$n_base, na.rm = TRUE)

        idx <- length(plots_list) + 1L
        plots_list[[idx]]       <- p
        titulos_list[[idx]]     <- titulo_bloque
        resumenN_list[[idx]]    <- sprintf("N = %s", format(n_total_bloque, big.mark = ",", scientific = FALSE))
        seccion_por_plot[idx]   <- sec
        vars_por_plot[idx]      <- v

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
      } else if (!is.na(list_name_v) && list_name_v %in% listnames_barras_agrupadas) {
        override     <- paste0("listnames_barras_agrupadas=", list_name_v)
        tipo_grafico <- "barras_agrupadas"
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

      # resumen N y ratio
      n_var <- .N_total_from_tab(tab_freq, total_casos)
      if (is.finite(n_var) && n_var >= 0 && total_casos > 0) {
        ratio <- n_var / total_casos * 100
        resumen_n_txt <- sprintf(
          "N = %s | Ratio de respuestas: %.1f%%",
          format(n_var, big.mark = ",", scientific = FALSE),
          ratio
        )
      } else if (is.finite(n_var)) {
        resumen_n_txt <- sprintf("N = %s", format(n_var, big.mark = ",", scientific = FALSE))
      } else {
        resumen_n_txt <- NULL
      }

      var_label    <- .titulo_var_safe(v)
      titulo_slide <- if (incluir_titulo_var) var_label else NULL

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
          colores_series <- estilos_barras_agrupadas$colores_series %||% c("Porcentaje" = "#004B8D")

          ori_final <- estilos_barras_agrupadas$orientacion %||% "horizontal"
          if (exists("orientacion_por_var", inherits = TRUE) && v %in% names(orientacion_por_var)) {
            ori_final <- orientacion_por_var[[v]]
          } else if (exists("orientacion_por_listname", inherits = TRUE) &&
                     !is.na(list_name_v) && nzchar(list_name_v) &&
                     list_name_v %in% names(orientacion_por_listname)) {
            ori_final <- orientacion_por_listname[[list_name_v]]
          } else if (exists("orientacion_default", inherits = TRUE) &&
                     is.character(orientacion_default) && length(orientacion_default) >= 1L) {
            ori_final <- orientacion_default[1]
          }

          ancho_eje_v <- estilos_barras_agrupadas$ancho_max_eje_y %||% NULL
          if (exists("ancho_eje_por_listname", inherits = TRUE) &&
              !is.na(list_name_v) && nzchar(list_name_v) &&
              list_name_v %in% names(ancho_eje_por_listname)) {
            ancho_eje_v <- ancho_eje_por_listname[[list_name_v]]
          }

          estilos_agrupadas_clean <- estilos_barras_agrupadas
          estilos_agrupadas_clean$ancho_max_eje_y <- NULL

          args_barras <- c(
            list(
              data             = tab_agr,
              var_categoria    = "categoria",
              var_n            = "n_base",
              cols_porcentaje  = cols_porcentaje,
              etiquetas_series = etiquetas_series,
              escala_valor     = "proporcion_1",
              colores_series   = colores_series,
              mostrar_valores  = TRUE,

              titulo           = NULL,
              subtitulo        = NULL,
              nota_pie         = NULL,

              mostrar_barra_extra = barra_extra == "total_n",
              prefijo_barra_extra = if (barra_extra == "total_n") "N = " else "N = ",
              titulo_barra_extra  = if (barra_extra == "total_n") "Total" else NULL,

              debug_ph_bordes     = isTRUE(debug_ph_bordes),

              exportar            = "rplot",
              orientacion         = ori_final
            ),
            if (!is.null(ancho_eje_v)) list(ancho_max_eje_y = ancho_eje_v) else list(),
            estilos_agrupadas_clean
          )

          p <- .safe_call(graficar_barras_agrupadas, args_barras)
        }

        if (tipo_grafico == "barras_apiladas") {

          tab_apil <- .build_tab_barras_apiladas(tab_freq, var_label)
          if (is.null(tab_apil)) next

          colores_grupos <- NULL
          preset_extra   <- NULL

          if (!is.na(list_name_v) && !is.null(colores_apiladas_por_listname[[list_name_v]])) {
            obj_col <- colores_apiladas_por_listname[[list_name_v]]
            if (is.list(obj_col)) {
              if (!is.null(obj_col$colores)) colores_grupos <- obj_col$colores
              if (!is.null(obj_col$preset_barra_extra)) preset_extra <- obj_col$preset_barra_extra
            } else {
              colores_grupos <- obj_col
            }
          }

          centro_cowplot <- NULL
          if (exists("centro_leyenda_por_var", inherits = TRUE) && v %in% names(centro_leyenda_por_var)) {
            centro_cowplot <- centro_leyenda_por_var[[v]]
          }
          if (is.null(centro_cowplot) &&
              exists("centro_leyenda_por_listname", inherits = TRUE) &&
              !is.na(list_name_v) && nzchar(list_name_v) &&
              list_name_v %in% names(centro_leyenda_por_listname)) {
            centro_cowplot <- centro_leyenda_por_listname[[list_name_v]]
          }

          ln_inv_seg <- estilos_barras_apiladas$listnames_invertir_segmentos %||% character(0)
          ln_inv_ley <- estilos_barras_apiladas$listnames_invertir_leyenda   %||% character(0)
          vars_inv_seg <- estilos_barras_apiladas$vars_invertir_segmentos %||% character(0)
          vars_inv_ley <- estilos_barras_apiladas$vars_invertir_leyenda   %||% character(0)

          invertir_segmentos_var <- (v %in% vars_inv_seg) ||
            (!is.na(list_name_v) && list_name_v %in% ln_inv_seg)

          invertir_leyenda_var <- (v %in% vars_inv_ley) ||
            (!is.na(list_name_v) && list_name_v %in% ln_inv_ley)

          estilos_apiladas_clean <- estilos_barras_apiladas
          estilos_apiladas_clean$listnames_invertir_segmentos <- NULL
          estilos_apiladas_clean$listnames_invertir_leyenda   <- NULL
          estilos_apiladas_clean$vars_invertir_segmentos      <- NULL
          estilos_apiladas_clean$vars_invertir_leyenda        <- NULL
          estilos_apiladas_clean$prefijo_barra_extra <- NULL
          estilos_apiladas_clean$mostrar_barra_extra <- NULL
          estilos_apiladas_clean$barra_extra_preset  <- NULL
          estilos_apiladas_clean$titulo_barra_extra  <- NULL

          args_apiladas <- c(
            list(
              data                = tab_apil$data,
              var_categoria       = "categoria",
              var_n               = "n_base",
              cols_porcentaje     = tab_apil$cols_porcentaje,
              etiquetas_grupos    = tab_apil$etiquetas_grupos,
              escala_valor        = "proporcion_1",
              colores_grupos      = colores_grupos,
              mostrar_valores     = TRUE,

              titulo              = NULL,
              subtitulo           = NULL,
              nota_pie            = NULL,

              mostrar_barra_extra = if (!is.null(preset_extra)) TRUE else (barra_extra == "total_n"),
              barra_extra_preset  = preset_extra %||% "ninguno",
              prefijo_barra_extra = if (!is.null(preset_extra)) "" else if (barra_extra == "total_n") "N = " else "",
              titulo_barra_extra  = NULL,

              debug_ph_bordes     = isTRUE(debug_ph_bordes),

              exportar           = "rplot",
              invertir_segmentos = invertir_segmentos_var,
              invertir_leyenda   = invertir_leyenda_var
            ),
            if (!is.null(centro_cowplot)) list(centro_cowplot = centro_cowplot) else list(),
            estilos_apiladas_clean
          )

          p <- .safe_call(graficar_barras_apiladas, args_apiladas)
        }

        if (tipo_grafico == "dico") {

          labels_dico <- NULL
          if (!is.null(dico_labels_por_var[[v]])) {
            labels_dico <- dico_labels_por_var[[v]]
          } else if (!is.na(list_name_v) && !is.null(dico_labels_por_listname[[list_name_v]])) {
            labels_dico <- dico_labels_por_listname[[list_name_v]]
          }

          if (is.null(labels_dico) || length(labels_dico) < 2) {
            warning(
              "En la variable '", v, "' no se encontraron etiquetas dicotómicas en `dico_labels_por_var` ni `dico_labels_por_listname`. Se usará barras agrupadas.",
              call. = FALSE
            )
            tipo_grafico_final <- "barras_agrupadas"

            tab_agr <- .build_tab_barras_agrupadas(tab_freq, var_label)
            if (is.null(tab_agr) || !nrow(tab_agr)) next

            args_barras <- c(
              list(
                data             = tab_agr,
                var_categoria    = "categoria",
                var_n            = "n_base",
                cols_porcentaje  = "pct",
                etiquetas_series = c(pct = "Porcentaje"),
                escala_valor     = "proporcion_1",
                colores_series   = estilos_barras_agrupadas$colores_series %||% c("Porcentaje" = "#004B8D"),
                mostrar_valores  = TRUE,

                titulo           = NULL,
                subtitulo        = NULL,
                nota_pie         = NULL,

                mostrar_barra_extra = barra_extra == "total_n",
                prefijo_barra_extra = if (barra_extra == "total_n") "N = " else "N = ",
                titulo_barra_extra  = if (barra_extra == "total_n") "Total" else NULL,

                debug_ph_bordes     = isTRUE(debug_ph_bordes),

                exportar            = "rplot"
              ),
              estilos_barras_agrupadas
            )

            p <- .safe_call(graficar_barras_agrupadas, args_barras)

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

                titulo            = NULL,
                subtitulo         = NULL,
                nota_pie          = NULL,

                debug_ph_bordes     = isTRUE(debug_ph_bordes),

                incluir_n_en_titulo = FALSE,
                exportar          = "rplot"
              ),
              estilos_dico
            )

            p <- .safe_call(graficar_dico, args_dico)
          }
        }

      } else if (tipo_grafico == "radar") {

        warning(
          "Tipo de gráfico 'radar' señalado para la variable '", v,
          "', pero el constructor genérico aún no está implementado. La variable se omitirá.",
          call. = FALSE
        )
        next

      } else {
        warning("Tipo de gráfico '", tipo_grafico, "' no reconocido para la variable '", v, "'. Se omitirá.", call. = FALSE)
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
      resumenN_list[[idx]]    <- resumen_n_txt
      seccion_por_plot[idx]   <- sec
      vars_por_plot[idx]      <- v
    }
  }

  log_decisiones <- dplyr::bind_rows(log_list)

  # ---------------------------------------------------------------------------
  # 5. Construir pares de diapositiva → índices de plots
  # ---------------------------------------------------------------------------
  pares_indices <- list()

  if (!is.null(pares_diapositiva) && length(pares_diapositiva) > 0) {
    for (nm in names(pares_diapositiva)) {

      vars_pair <- pares_diapositiva[[nm]]$vars
      if (length(vars_pair) != 2) next

      v1 <- vars_pair[1]
      v2 <- vars_pair[2]

      idx1 <- match(v1, vars_por_plot)
      idx2 <- match(v2, vars_por_plot)

      if (!is.na(idx1) && !is.na(idx2) && idx1 != idx2) {
        idx_ordered <- sort(c(idx1, idx2))
        pares_indices[[nm]] <- list(
          idx1   = idx_ordered[1],
          idx2   = idx_ordered[2],
          titulo = pares_diapositiva[[nm]]$titulo %||% NULL
        )
      }
    }
  }

  indices_segundos_pares <- if (length(pares_indices)) vapply(pares_indices, function(x) x$idx2, integer(1)) else integer(0)
  indices_primeros_pares <- if (length(pares_indices)) vapply(pares_indices, function(x) x$idx1, integer(1)) else integer(0)

  .usar_layout_unabarra <- function(var_i, log_decisiones, tiene_layout_unabarra) {
    if (!isTRUE(tiene_layout_unabarra)) return(FALSE)
    if (is.null(log_decisiones) || !nrow(log_decisiones)) return(FALSE)
    fila_i <- which(log_decisiones$var == var_i)
    if (!length(fila_i)) return(FALSE)
    identical(log_decisiones$tipo_grafico[fila_i[1]], "barras_apiladas")
  }

  # ---------------------------------------------------------------------------
  # 6. PPT (export)
  # ---------------------------------------------------------------------------
  if (!solo_lista) {

    if (!requireNamespace("officer", quietly = TRUE) ||
        !requireNamespace("rvg", quietly = TRUE)) {
      stop("Para exportar a PPT se requieren los paquetes 'officer' y 'rvg'.", call. = FALSE)
    }

    # Helper: ph_location_type compatible (officer >= 0.6.7 usa type_idx)
    .ph_loc_type <- function(type, idx = NULL) {
      if (is.null(idx)) return(officer::ph_location_type(type = type))
      tryCatch(
        officer::ph_location_type(type = type, type_idx = idx),
        error = function(e) officer::ph_location_type(type = type, id = idx)
      )
    }

    # 6.1. Leer plantilla
    if (is.null(template_pptx)) {
      template_interno <- system.file("plantillas/plantilla_16_9.pptx", package = "prosecnur")
      if (nzchar(template_interno) && file.exists(template_interno)) {
        if (mensajes_progreso) message("Usando plantilla interna 16:9: ", template_interno)
        doc <- officer::read_pptx(path = template_interno)
      } else {
        if (mensajes_progreso) {
          message("No se encontró 'plantilla_16_9.pptx' dentro del paquete. Se usará la plantilla por defecto de PowerPoint.")
        }
        doc <- officer::read_pptx()
      }
    } else {
      if (!file.exists(template_pptx)) {
        stop("No se encontró el archivo de plantilla especificado en `template_pptx`: ", template_pptx, call. = FALSE)
      }
      if (mensajes_progreso) message("Usando plantilla definida por el usuario: ", template_pptx)
      doc <- officer::read_pptx(path = template_pptx)
    }

    # 6.2. Info de layouts
    layout_info <- tryCatch(officer::layout_summary(doc), error = function(e) NULL)

    tiene_layout_graficos        <- FALSE
    layout_graficos              <- "Blank"
    usar_pic_placeholder         <- FALSE
    tiene_layout_title_slide     <- FALSE
    tiene_layout_contraportada   <- FALSE
    tiene_layout_section_header  <- FALSE

    tiene_layout_unabarra <- FALSE
    layout_unabarra       <- NULL

    if (!is.null(layout_info)) {
      if ("Graficos2" %in% layout_info$layout) {
        tiene_layout_graficos <- TRUE
        layout_graficos       <- "Graficos2"
        usar_pic_placeholder  <- TRUE
      } else if ("Graficos" %in% layout_info$layout) {
        tiene_layout_graficos <- TRUE
        layout_graficos       <- "Graficos"
        usar_pic_placeholder  <- TRUE
      }

      if ("Graficos_unabarra" %in% layout_info$layout) {
        tiene_layout_unabarra <- TRUE
        layout_unabarra       <- "Graficos_unabarra"
        if (mensajes_progreso) message("Layout 'Graficos_unabarra' disponible para gráficos de una sola barra.")
      }

      if ("Title Slide" %in% layout_info$layout)   tiene_layout_title_slide <- TRUE
      if ("Contraportada" %in% layout_info$layout) tiene_layout_contraportada <- TRUE
      if ("Section Header" %in% layout_info$layout) tiene_layout_section_header <- TRUE
    }

    tiene_layout_doble <- FALSE
    layout_doble <- NULL
    if (!is.null(layout_info) && "Graficos_2columnas" %in% layout_info$layout) {
      tiene_layout_doble <- TRUE
      layout_doble <- "Graficos_2columnas"
      if (mensajes_progreso) message("Layout 'Graficos_2columnas' disponible para gráficos dobles.")
    }

    if (mensajes_progreso) {
      if (tiene_layout_graficos) {
        message("Las diapositivas de gráficos usarán el layout '", layout_graficos, "'.")
      } else {
        message("No se encontró un layout 'Graficos' ni 'Graficos2'; se usará 'Blank' a pantalla completa.")
      }
    }

    # 6.3. Portada
    if (tiene_layout_title_slide &&
        (( !is.null(titulo_portada)    && nzchar(titulo_portada)) ||
         ( !is.null(subtitulo_portada) && nzchar(subtitulo_portada)) ||
         ( !is.null(fecha_portada)     && nzchar(fecha_portada)))) {

      if (mensajes_progreso) message("Agregando diapositiva de portada (Title Slide).")

      doc <- officer::add_slide(doc, layout = "Title Slide", master = "Office Theme")

      if (!is.null(titulo_portada) && nzchar(titulo_portada)) {
        loc_title <- tryCatch(.ph_loc_type("ctrTitle"),
                              error = function(e) .ph_loc_type("title"))
        doc <- tryCatch(officer::ph_with(doc, titulo_portada, location = loc_title),
                        error = function(e) doc)
      }

      if (!is.null(subtitulo_portada) && nzchar(subtitulo_portada)) {
        doc <- tryCatch(
          officer::ph_with(doc, subtitulo_portada, location = .ph_loc_type("subTitle")),
          error = function(e) doc
        )
      }

      if (!is.null(fecha_portada) && nzchar(fecha_portada)) {
        doc <- tryCatch(
          officer::ph_with(doc, fecha_portada, location = .ph_loc_type("dt")),
          error = function(e) doc
        )
      }
    } else if (( !is.null(titulo_portada) || !is.null(subtitulo_portada) || !is.null(fecha_portada)) && mensajes_progreso) {
      if (!tiene_layout_title_slide) message("Se solicitaron textos de portada, pero la plantilla no tiene layout 'Title Slide'.")
    }

    # 6.4. Diapositivas de gráficos
    if (length(plots_list)) {

      extract_pretty_N <- function(n_txt) {
        if (is.null(n_txt) || !nzchar(n_txt)) return(NULL)
        n_part   <- sub("^N\\s*=\\s*([^|]+).*", "\\1", n_txt)
        n_digits <- gsub("[^0-9]", "", n_part)
        if (!nzchar(n_digits)) return(NULL)
        n_num <- suppressWarnings(as.numeric(n_digits))
        if (!is.finite(n_num)) return(NULL)
        format(n_num, big.mark = ",", scientific = FALSE)
      }

      for (i in seq_along(plots_list)) {

        if (i %in% indices_segundos_pares) next

        # Pares
        if (i %in% indices_primeros_pares && tiene_layout_doble) {

          par_nm   <- names(pares_indices)[which(indices_primeros_pares == i)]
          info_par <- pares_indices[[par_nm]]
          j        <- info_par$idx2

          p1 <- plots_list[[i]]
          p2 <- plots_list[[j]]

          t1 <- titulos_list[[i]] %||% ""
          t2 <- titulos_list[[j]] %||% ""
          titulo_final <- info_par$titulo %||% paste0(t1, " / ", t2)

          N_pretty_i <- extract_pretty_N(resumenN_list[[i]])
          base1 <- if (!is.null(N_pretty_i)) {
            if (!is.null(fuente) && nzchar(fuente)) paste0("Base: ", N_pretty_i, " ", fuente) else paste0("Base: ", N_pretty_i)
          } else {
            fuente %||% ""
          }
          base2 <- " "

          # Section header si cambia sección
          if (tiene_layout_section_header && length(seccion_por_plot) >= i) {
            sec_i <- seccion_por_plot[i]
            if (!is.null(sec_i) && nzchar(sec_i) && (i == 1L || !identical(sec_i, seccion_por_plot[i - 1L]))) {
              doc <- officer::add_slide(doc, layout = "Section Header", master = "Office Theme")
              loc_sec_title <- tryCatch(.ph_loc_type("title"),
                                        error = function(e) tryCatch(.ph_loc_type("ctrTitle"),
                                                                     error = function(e2) NULL))
              if (!is.null(loc_sec_title)) {
                doc <- tryCatch(officer::ph_with(doc, sec_i, location = loc_sec_title), error = function(e) doc)
              }
            }
          }

          doc <- officer::add_slide(doc, layout = layout_doble, master = "Office Theme")

          doc <- tryCatch(officer::ph_with(doc, titulo_final, location = .ph_loc_type("title")),
                          error = function(e) doc)

          # Gráficos (type_idx en vez de id)
          doc <- officer::ph_with(doc, rvg::dml(ggobj = p1, bg = "transparent"),
                                  location = .ph_loc_type("pic", idx = 2))
          doc <- officer::ph_with(doc, rvg::dml(ggobj = p2, bg = "transparent"),
                                  location = .ph_loc_type("pic", idx = 1))

          doc <- officer::ph_with(doc, base1, location = .ph_loc_type("body", idx = 2))
          doc <- officer::ph_with(doc, base2, location = .ph_loc_type("body", idx = 3))

          next
        }

        # Section header si cambia sección
        if (tiene_layout_section_header && length(seccion_por_plot) >= i) {
          sec_i <- seccion_por_plot[i]
          if (!is.null(sec_i) && nzchar(sec_i) && (i == 1L || !identical(sec_i, seccion_por_plot[i - 1L]))) {
            doc <- officer::add_slide(doc, layout = "Section Header", master = "Office Theme")
            loc_sec_title <- tryCatch(.ph_loc_type("title"),
                                      error = function(e) tryCatch(.ph_loc_type("ctrTitle"),
                                                                   error = function(e2) NULL))
            if (!is.null(loc_sec_title)) {
              doc <- tryCatch(officer::ph_with(doc, sec_i, location = loc_sec_title), error = function(e) doc)
            }
          }
        }

        # Layout unabarra si aplica
        layout_i <- layout_graficos
        if (.usar_layout_unabarra(vars_por_plot[i], log_decisiones, tiene_layout_unabarra)) {
          layout_i <- layout_unabarra
          if (mensajes_progreso) message("   · Usando layout '", layout_i, "' para gráfico de una sola barra (", vars_por_plot[i], ").")
        }

        doc <- officer::add_slide(doc, layout = layout_i, master = "Office Theme")

        # Título del slide
        st <- titulos_list[[i]] %||% NULL
        if (!is.null(st) && nzchar(st)) {
          loc_gtitle <- tryCatch(.ph_loc_type("title"),
                                 error = function(e) tryCatch(.ph_loc_type("ctrTitle"),
                                                              error = function(e2) NULL))
          if (!is.null(loc_gtitle)) {
            doc <- tryCatch(officer::ph_with(doc, st, location = loc_gtitle), error = function(e) doc)
          }
        }

        # Insertar gráfico
        loc_pic <- if (usar_pic_placeholder) .ph_loc_type("pic") else officer::ph_location_fullsize()
        doc <- officer::ph_with(doc, rvg::dml(ggobj = plots_list[[i]], bg = "transparent"), location = loc_pic)

        # Base izquierda (body idx = 2)
        if (tiene_layout_graficos) {
          loc_fuente <- tryCatch(.ph_loc_type("body", idx = 2), error = function(e) NULL)
          if (!is.null(loc_fuente)) {
            N_pretty <- extract_pretty_N(resumenN_list[[i]])
            base_texto <- if (!is.null(N_pretty)) {
              if (!is.null(fuente) && nzchar(fuente)) paste0("Base: ", N_pretty, " ", fuente) else paste0("Base: ", N_pretty)
            } else {
              fuente %||% ""
            }
            doc <- tryCatch(officer::ph_with(doc, base_texto, location = loc_fuente), error = function(e) doc)
          }
        }

        # Bloque derecho (body idx = 3): vacío
        if (tiene_layout_graficos) {
          loc_resumen <- tryCatch(.ph_loc_type("body", idx = 3), error = function(e) NULL)
          if (!is.null(loc_resumen)) {
            doc <- tryCatch(officer::ph_with(doc, " ", location = loc_resumen), error = function(e) doc)
          }
        }
      }
    }

    # 6.5. Contraportada
    if (tiene_layout_contraportada) {
      if (mensajes_progreso) message("Agregando diapositiva de contraportada.")
      doc <- officer::add_slide(doc, layout = "Contraportada", master = "Office Theme")

      if (!is.null(fecha_portada) && nzchar(fecha_portada)) {
        doc <- tryCatch(
          officer::ph_with(doc, fecha_portada, location = .ph_loc_type("dt")),
          error = function(e) doc
        )
      }
    }

    # 6.6. Guardar
    print(doc, target = path_ppt)
    if (mensajes_progreso) {
      message("PPT generado en: ", normalizePath(path_ppt, winslash = "/"))
    }
  }

  invisible(list(
    plots          = plots_list,
    log_decisiones = log_decisiones
  ))
}
