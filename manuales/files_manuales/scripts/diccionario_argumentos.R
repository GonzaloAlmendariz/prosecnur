`%||%` <- function(x, y) {
  if (is.null(x) || !length(x)) y else x
}

manual_regex_escape <- function(x) {
  gsub("([][{}()+*^$|\\\\?.])", "\\\\\\1", x)
}

manual_rd_to_text <- function(x) {
  if (is.null(x) || !length(x)) {
    return("")
  }

  if (is.character(x)) {
    return(paste(x, collapse = ""))
  }

  if (is.list(x)) {
    tag <- attr(x, "Rd_tag", exact = TRUE)
    inner <- paste(vapply(x, manual_rd_to_text, character(1)), collapse = "")

    if (identical(tag, "\\code")) {
      return(paste0("`", inner, "`"))
    }

    return(inner)
  }

  paste(as.character(x), collapse = "")
}

manual_clean_text <- function(x) {
  x <- gsub("[\r\n\t]+", " ", x)
  x <- gsub("\\s+", " ", x)
  trimws(x)
}

manual_short_text <- function(x, max_chars = 220) {
  x <- manual_clean_text(x)

  if (!nzchar(x)) {
    return("")
  }

  pos <- regexpr("([.!?])\\s", x, perl = TRUE)[1]

  if (!is.na(pos) && pos > 0 && pos <= max_chars) {
    return(substr(x, 1, pos))
  }

  if (nchar(x) > max_chars) {
    return(paste0(substr(x, 1, max_chars - 1), "…"))
  }

  x
}

manual_rd_section <- function(rd, tag) {
  idx <- which(vapply(
    rd,
    function(node) identical(attr(node, "Rd_tag", exact = TRUE), tag),
    logical(1)
  ))

  if (!length(idx)) {
    return(NULL)
  }

  rd[[idx[1]]]
}

manual_rd_info <- function(fun, package = "prosecnur") {
  h <- do.call(utils::help, list(topic = fun, package = package))

  if (length(h) == 0) {
    return(NULL)
  }

  rd <- utils:::.getHelpFile(h)

  title <- manual_short_text(
    manual_rd_to_text(manual_rd_section(rd, "\\title")),
    max_chars = 180
  )

  description <- manual_short_text(
    manual_rd_to_text(manual_rd_section(rd, "\\description")),
    max_chars = 260
  )

  arguments_node <- manual_rd_section(rd, "\\arguments")
  arguments <- list()

  if (!is.null(arguments_node)) {
    items <- Filter(
      function(node) identical(attr(node, "Rd_tag", exact = TRUE), "\\item"),
      arguments_node
    )

    for (item in items) {
      arg_name <- manual_clean_text(manual_rd_to_text(item[[1]]))
      arg_desc <- manual_short_text(manual_rd_to_text(item[[2]]), max_chars = 260)
      arguments[[arg_name]] <- arg_desc
    }
  }

  list(
    title = title,
    description = description,
    arguments = arguments
  )
}

manual_default_label <- function(x) {
  txt <- manual_clean_text(paste(deparse(x), collapse = " "))

  if (!nzchar(txt)) {
    return("Requerido")
  }

  if (nchar(txt) > 42) {
    txt <- paste0(substr(txt, 1, 39), "...")
  }

  paste0("`", txt, "`")
}

manual_fallback_arg_desc <- function(arg) {
  switch(
    arg,
    path = "Ruta del archivo que la función necesita leer o escribir.",
    path_instrumento = "Ruta del instrumento que se usará como base del proceso.",
    path_instrumento_in = "Ruta del instrumento original antes de adaptarlo.",
    path_instrumento_out = "Ruta donde se guardará el instrumento ya adaptado.",
    path_datos = "Ruta de la base de datos con la que se trabajará.",
    path_data_adaptada = "Ruta de la base ya adaptada que se usará como insumo.",
    path_plantilla = "Ruta de la plantilla de codificación ya preparada o resuelta.",
    path_familias = "Ruta de la hoja de familias o clasificación de variables.",
    path_xlsx = "Ruta del Excel que la función dejará como salida.",
    path_sav = "Ruta del archivo `.sav` que se quiere exportar.",
    path_sps = "Ruta del archivo `.sps` complementario que se quiere generar.",
    path_ppt = "Ruta del PowerPoint final que se quiere generar.",
    path_docx = "Ruta del documento Word final que se quiere generar.",
    data = "Base principal sobre la que la función calcula, grafica o exporta resultados.",
    dat = "Base de datos que se usa como insumo dentro del flujo.",
    instrumento = "Objeto de instrumento o metadatos que acompaña a la base.",
    inst = "Objeto de instrumento que ayuda a interpretar variables, labels y opciones.",
    plantilla = "Objeto o tabla de trabajo donde ya están las decisiones de codificación.",
    split = "Resultado de la clasificación de familias que organiza qué variables entran al flujo.",
    survey = "Tabla o metadata de la hoja survey del instrumento.",
    choices = "Tabla o metadata de la hoja choices del instrumento.",
    var = "Nombre de la variable principal que se quiere procesar o graficar.",
    vars = "Conjunto de variables sobre las que se aplicará la operación.",
    title = "Título visible para la persona que leerá la salida.",
    subtitle = "Bajada corta que acompaña el título principal.",
    date = "Fecha o texto temporal que se mostrará en la salida.",
    plot = "Gráfico o pieza visual que se insertará dentro del layout.",
    presets = "Reglas visuales generales para colores, tamaños y estilo.",
    presets_ppt = "Configuración visual que se usará en la salida PPT.",
    presets_word = "Configuración visual que se usará en la salida Word.",
    template_pptx = "Plantilla de PowerPoint sobre la que se construirá el archivo final.",
    master = "Nombre del master o tema de la plantilla PPT que se usará al exportar.",
    plan = "Plan ya armado con la narrativa y los bloques visuales del reporte.",
    env_diapos = "Entorno donde `diapo()` fue acumulando las diapositivas.",
    output_file = "Nombre o ruta del archivo final que se va a exportar.",
    sheet = "Nombre de la hoja del Excel con la que se quiere trabajar.",
    sheet_survey = "Nombre de la hoja del XLSForm donde están las preguntas.",
    sheet_choices = "Nombre de la hoja del XLSForm donde están las opciones de respuesta.",
    lang = "Idioma o columna de etiquetas que quiere priorizar al leer el instrumento.",
    prefer_label = "Columna de label que quiere privilegiar si el instrumento tiene varias.",
    secciones = "Lista o selección de secciones que quiere incluir en la salida.",
    SECCIONES = "Lista o selección de secciones que quiere incluir en la salida.",
    cruces = "Variables que se usarán para abrir tablas comparativas.",
    ord = "Información extra de orden o etiquetas que ayuda a enriquecer la salida.",
    orden = "Orden en que quiere ver los resultados dentro de la salida.",
    var_peso = "Variable de ponderación, si la base usa pesos.",
    dummy_vars = "Variables dummy adicionales que conviene tratar como opciones sí/no.",
    dummies_na_to_zero = "Regla sobre cómo tratar valores faltantes en variables dummy.",
    ordinal_vars = "Variables que quiere tratar como ordinales en el reporte.",
    ordinal_list_names = "Listas del instrumento que quiere tratar como ordinales.",
    listas_ordinales = "Listas del instrumento que deben respetar un orden de respuesta.",
    vars_fecha = "Variables que conviene tratar como fechas en la salida final.",
    vars_hora = "Variables que conviene tratar como horas en la salida final.",
    vars_datetime = "Variables que conviene tratar como fecha y hora en la salida final.",
    sm_vars = "Variables `select_multiple` que entrarán al flujo de adaptación.",
    sm_vars_force = "Variables `select_multiple` que quiere forzar dentro de una salida aunque no entren solas.",
    so_parent_vars = "Variables `select_one` padre que se adaptarán como preguntas principales.",
    so_child_vars = "Variables `select_one` hijas o derivadas que también se adaptarán.",
    int_vars = "Variables enteras que quiere recodificar o adaptar por rangos o nuevos grupos.",
    integer_vars = "Variables enteras que quiere convertir en nuevas categorías recodificadas.",
    out_path = "Ruta donde se guardará la salida generada por la función.",
    include_familias = "Indica si la salida debe incorporar también la hoja de familias.",
    choices_order = "Regla para ordenar los códigos u opciones nuevas que crea la función.",
    paint = "Si quiere colorear visualmente lo nuevo para revisarlo más fácil.",
    autofiltro = "Si quiere dejar filtros listos en el Excel para facilitar la revisión.",
    congelar_encabezado = "Si quiere inmovilizar el encabezado para navegar mejor el Excel.",
    incluir_text_vars = "Si quiere incluir también preguntas de texto en la plantilla de familias.",
    verbose = "Si quiere que la función vaya informando lo que encuentra o hace.",
    verbose_sps = "Si quiere mensajes mientras se construye el archivo para SPSS.",
    compress = "Si quiere guardar el `.sav` con compresión.",
    decimales_2 = "Variables que conviene guardar con dos decimales en SPSS.",
    codigos_solo_si_presentes = "Códigos especiales que solo quiere mostrar si realmente aparecen en los datos.",
    numericas = "Variables numéricas que quiere tratar de forma explícita en la salida.",
    fuente = "Texto de fuente o pie de página que acompañará la salida.",
    mostrar_todo = "Si quiere ver toda la salida sin recortar bloques menos relevantes.",
    modo = "Modo general con el que la función organizará la salida.",
    user_na = "Indica si ciertas respuestas deben entrar como valores perdidos de usuario.",
    base = "Ajustes generales que servirán de base para el resto de presets.",
    barras_apiladas = "Reglas visuales específicas para gráficos de barras apiladas.",
    multi_apiladas = "Reglas visuales específicas para gráficos con varias barras apiladas.",
    barras_agrupadas = "Reglas visuales específicas para gráficos de barras agrupadas.",
    barras_numericas = "Reglas visuales para gráficos basados en variables numéricas.",
    boxplot = "Reglas visuales para boxplots.",
    pie = "Reglas visuales para gráficos de torta.",
    donut = "Reglas visuales para gráficos tipo dona.",
    radar_tabla = "Reglas visuales para el layout de radar con tabla.",
    dim_heatmap = "Reglas visuales para heatmaps de dimensiones.",
    dim_radar = "Reglas visuales para radares de dimensiones.",
    dim_foda = "Reglas visuales para salidas FODA de dimensiones.",
    numerico = "Reglas visuales para tarjetas o salidas numéricas simples.",
    debug = "Opciones extra para revisar o afinar el plan visual.",
    slide = "Diapositiva o bloque ya armado que quiere agregar al plan.",
    env = "Entorno donde se guardará o leerá el plan acumulado.",
    strict_diapos = "Si quiere exigir que el plan venga exactamente desde `diapo()`.",
    mensajes_progreso = "Si quiere ver mensajes de avance mientras se exporta.",
    solo_lista = "Si quiere obtener solo la lista/plan sin exportar todavía el archivo final.",
    build_render_meta = "Si quiere además construir metadatos internos del render.",
    col_enumerador = "Nombre de la columna que identifica al encuestador o enumerador.",
    cols_corte = "Columnas con las que quiere cortar o resumir el reporte.",
    ... = "Argumentos adicionales para afinar el comportamiento de la función.",
    "Argumento usado por la función para controlar una parte específica del proceso."
  )
}

manual_bad_arg_desc <- function(x) {
  x <- manual_clean_text(x)
  x %in% c(
    "", "character.", "logical.", "list.", "numeric.", "integer.", "double.",
    "data.frame.", "tibble.", "function.", "matrix.", "character", "logical",
    "list", "numeric", "integer", "double", "data.frame", "tibble", "function"
  )
}

manual_fallback_fun_desc <- function(fun) {
  switch(
    fun,
    leer_instrumento_xlsform = "Lee el instrumento en formato XLSForm y deja listas las preguntas y opciones para el resto del flujo.",
    leer_datos = "Lee la base de datos sin cambiar sus valores y conserva un mapa entre nombres originales y nombres limpios.",
    escribir_plantilla_familias = "Crea la primera hoja de trabajo donde el equipo decide qué variables entran realmente a codificación.",
    leer_familias_clasificar = "Lee la hoja de familias ya editada y organiza las variables según el tipo de trabajo que recibirán.",
    construir_plantilla_desde_familias = "Convierte la clasificación de familias en una plantilla de codificación lista para que el equipo recodifique.",
    exportar_plantilla_codificacion_xlsx = "Exporta la plantilla de codificación a Excel con una estructura pensada para revisión y edición manual.",
    ppra_adaptar_data = "Aplica la plantilla resuelta a la base de datos y genera una versión adaptada y consistente.",
    ppra_adaptar_instrumento = "Actualiza el instrumento para que documente las nuevas variables y categorías creadas en la adaptación.",
    reporte_instrumento = "Organiza el instrumento en un objeto de metadatos que luego alimenta tablas, codebooks y exportaciones.",
    reporte_data = "Toma la base ya adaptada y la deja preparada para reportes, tablas y exportaciones compatibles con SPSS.",
    reporte_codebook = "Genera un libro de códigos en Excel a partir de la base ya preparada para reporte.",
    reporte_spss = "Exporta la base a formatos compatibles con SPSS para intercambio o análisis posterior.",
    reporte_frecuencias = "Produce tablas simples de frecuencia a partir de la base preparada.",
    reporte_cruces = "Produce tablas cruzadas simples para comparar resultados entre grupos.",
    reporte_enumeradores = "Genera un reporte de seguimiento de campo resumido por encuestador o enumerador.",
    surveymonkey_leer = "Lee el archivo exportado desde SurveyMonkey para empezar el flujo con esa fuente.",
    surveymonkey_xlsform = "Convierte o prepara la metadata de SurveyMonkey para que el paquete la pueda tratar como instrumento.",
    surveymonkey_data = "Ajusta la base proveniente de SurveyMonkey para conectarla con el flujo de reportes y planes.",
    p_presets = "Reúne reglas visuales comunes para que todos los gráficos y slides compartan el mismo estilo.",
    p_reset = "Limpia el plan acumulado para empezar una nueva narrativa visual desde cero.",
    p_barras_apiladas = "Construye una pieza gráfica de barras apiladas lista para entrar a una diapositiva.",
    p_barras_agrupadas = "Construye una pieza gráfica de barras agrupadas lista para entrar a una diapositiva.",
    p_barras_multiapiladas = "Construye una pieza visual pensada para comparar varias barras apiladas en un mismo bloque.",
    p_numerico = "Construye una pieza visual centrada en uno o varios indicadores numéricos.",
    p_pie = "Construye un gráfico de torta listo para usar dentro del plan visual.",
    p_donut = "Construye un gráfico tipo dona listo para usar dentro del plan visual.",
    p_radar_tabla = "Construye una pieza combinada de radar y tabla para presentar resultados resumidos.",
    p_text = "Construye un bloque de texto para destacar mensajes o interpretaciones dentro del slide.",
    p_slide_title = "Arma una portada para el plan visual.",
    p_slide_section = "Arma una diapositiva de separación entre secciones del reporte.",
    p_slide_1 = "Arma un layout con un gráfico principal a pantalla completa.",
    p_slide_2 = "Arma un layout con dos piezas visuales en una misma diapositiva.",
    p_slide_text_l = "Arma un layout donde el texto acompaña al gráfico desde la izquierda.",
    p_slide_text_r = "Arma un layout donde el texto acompaña al gráfico desde la derecha.",
    p_slide_poblacion_2 = "Arma un layout pensado para mostrar perfiles o composiciones de población en dos bloques.",
    diapo = "Agrega una diapositiva ya armada al plan acumulado que luego se exportará.",
    p_get_plan = "Recupera el plan acumulado para revisarlo, guardarlo o reutilizarlo.",
    p_plan = "Construye un plan visual como objeto completo en una sola estructura.",
    graficar_ppt = "Exporta gráficos sueltos a PowerPoint, normalmente uno por diapositiva.",
    reporte_ppt_plan = "Toma un plan visual ya armado y lo exporta a un PowerPoint final.",
    w_presets = "Define reglas visuales equivalentes para que el mismo plan pueda salir también a Word.",
    reporte_word_plan = "Toma el mismo plan visual y lo convierte en un documento Word final.",
    "Función del paquete usada en este manual."
  )
}

manual_md_escape <- function(x) {
  x <- gsub("\\|", "\\\\|", x)
  x <- gsub("<", "&lt;", x, fixed = TRUE)
  x <- gsub(">", "&gt;", x, fixed = TRUE)
  x
}

manual_exported_mentions <- function(qmd_path, package = "prosecnur") {
  txt <- paste(readLines(qmd_path, warn = FALSE, encoding = "UTF-8"), collapse = "\n")
  exports <- sort(getNamespaceExports(package))

  positions <- vapply(exports, function(fun) {
    fun_rx <- manual_regex_escape(fun)

    hits <- c(
      regexpr(paste0("`", fun_rx, "\\(\\)`"), txt, perl = TRUE)[1],
      regexpr(paste0("`", fun_rx, "`"), txt, perl = TRUE)[1],
      regexpr(paste0("(?<![A-Za-z0-9_.])", fun_rx, "\\s*\\("), txt, perl = TRUE)[1]
    )

    hits <- hits[hits > 0]

    if (!length(hits)) Inf else min(hits)
  }, numeric(1))

  mentions <- exports[is.finite(positions)]
  mentions[order(positions[is.finite(positions)], mentions)]
}

manual_function_dictionary <- function(fun, package = "prosecnur") {
  obj <- tryCatch(
    get(fun, envir = asNamespace(package), inherits = FALSE),
    error = function(e) NULL
  )

  if (is.null(obj)) {
    return(NULL)
  }

  rd <- tryCatch(manual_rd_info(fun, package = package), error = function(e) NULL)
  formal_list <- formals(obj)

  if (is.null(formal_list)) {
    formal_list <- pairlist()
  }

  arg_names <- names(formal_list) %||% character(0)

  rows <- lapply(arg_names, function(arg) {
    arg_desc <- rd$arguments[[arg]] %||% ""

    if (manual_bad_arg_desc(arg_desc)) {
      arg_desc <- manual_fallback_arg_desc(arg)
    }

    list(
      argumento = arg,
      descripcion = arg_desc,
      default = manual_default_label(formal_list[[arg]])
    )
  })

  list(
    fun = fun,
    title = rd$title %||% paste("Función", fun),
    description = rd$description %||% manual_fallback_fun_desc(fun),
    rows = rows
  )
}

manual_render_argument_dictionary <- function(
  qmd_path,
  package = "prosecnur",
  prioridad = character(0)
) {
  funciones <- manual_exported_mentions(qmd_path = qmd_path, package = package)

  if (length(prioridad)) {
    prioridad <- prioridad[prioridad %in% funciones]
    funciones <- c(prioridad, setdiff(funciones, prioridad))
  }

  cat("# Anexo. Diccionario de argumentos\n\n")
  cat(
    "Este anexo reúne las funciones de `", package,
    "` que aparecen en esta guía. Para cada una se resume para qué sirve y qué hace cada argumento con un lenguaje de uso, no de implementación.\n\n",
    sep = ""
  )

  if (!length(funciones)) {
    cat("No se detectaron funciones documentadas de `", package, "` en este manual.\n", sep = "")
    return(invisible(character(0)))
  }

  for (fun in funciones) {
    info <- manual_function_dictionary(fun, package = package)

    if (is.null(info)) {
      next
    }

    cat("## ", fun, "()\n\n", sep = "")
    cat("**Para qué sirve.** ", manual_md_escape(info$description), "\n\n", sep = "")

    if (!length(info$rows)) {
      cat("Esta función no expone argumentos relevantes para este anexo.\n\n")
      next
    }

    cat("| Argumento | Qué hace | Por defecto |\n")
    cat("|---|---|---|\n")

    for (row in info$rows) {
      cat(
        "| `", manual_md_escape(row$argumento), "` | ",
        manual_md_escape(row$descripcion), " | ",
        manual_md_escape(row$default), " |\n",
        sep = ""
      )
    }

    cat("\n")
  }

  invisible(funciones)
}
