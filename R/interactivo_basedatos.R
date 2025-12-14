# =============================================================================
# Tab 2: Base de datos (UI + server) — SM dummies visibles + diccionario elegante
# =============================================================================
#' @keywords internal
#' @noRd

.ui_tab_base_datos <- function(ctx) {

  shiny::sidebarLayout(
    shiny::sidebarPanel(
      width = 3,
      shiny::h3("Diccionario"),
      shiny::p("Información de variables con categorías codificadas."),

      shiny::selectInput(
        inputId  = "data_seccion",
        label    = "Sección",
        choices  = stats::setNames(ctx$secciones_nombres, ctx$secciones_nombres),
        selected = ctx$secciones_nombres[1]
      ),

      shiny::selectInput(
        inputId  = "dicc_var",
        label    = "Variable",
        choices  = c(),
        selected = NULL
      ),

      shiny::div(
        class = "cardbox",
        style = "padding: 10px; margin-top: 10px;",
        shiny::uiOutput("diccionario_detalle")
      ),

      shiny::hr(),

      shiny::h3("Vista"),
      shiny::div(
        class = "toggle-row",
        shiny::span(class = "toggle-label", "Códigos"),
        shiny::tags$label(
          class = "switch",
          shiny::tags$input(id = "vista_etiquetas", type = "checkbox", checked = "checked"),
          shiny::tags$span(class = "slider")
        ),
        shiny::span(class = "toggle-label", "Etiquetas")
      )
    ),

    shiny::mainPanel(
      width = 9,

      shiny::fluidRow(
        shiny::column(
          width = 12,
          shiny::div(
            class = "cardbox",
            shiny::div(
              class = "cardbox-header",
              shiny::div(class = "cardbox-title", "Base de datos"),
              shiny::div(class = "cardbox-subtitle", "Búsqueda, ordenamiento y paginación disponibles.")
            ),
            DT::dataTableOutput("tabla_data")
          )
        )
      ),

      shiny::div(style = "height: 48px;")
    )
  )
}

#' @keywords internal
#' @noRd
.server_tab_base_datos <- function(ctx, input, output, session) {

  data        <- ctx$data
  instrumento <- ctx$instrumento

  `%||%` <- get0("%||%", ifnotfound = function(x, y) if (!is.null(x)) x else y)

  # ---------------------------------------------------------------------------
  # Helpers: labels / list_name / map code->label choices
  # ---------------------------------------------------------------------------
  .obtener_label_var <- get0(
    ".obtener_label_var",
    ifnotfound = function(var, instrumento, data = NULL) {
      surv <- instrumento$survey
      if (!is.null(surv) && all(c("name","label") %in% names(surv)) && var %in% surv$name) {
        lab <- surv$label[surv$name == var][1]
        if (!is.na(lab) && nzchar(as.character(lab))) return(as.character(lab))
      }
      if (!is.null(data) && var %in% names(data)) {
        vl <- attr(data[[var]], "label", exact = TRUE)
        if (!is.null(vl) && nzchar(as.character(vl))) return(as.character(vl))
      }
      as.character(var)
    }
  )

  .get_list_name <- function(var) {
    surv <- instrumento$survey %||% NULL
    if (is.null(surv) || !all(c("name","list_name") %in% names(surv))) return(NA_character_)
    ln <- as.character(surv$list_name[surv$name == var][1])
    if (is.na(ln) || !nzchar(ln)) NA_character_ else ln
  }

  .choice_map <- function(var) {
    ln <- .get_list_name(var)
    ch <- instrumento$choices %||% NULL
    if (is.null(ch) || !all(c("list_name","name") %in% names(ch))) return(list())

    label_col <- if ("label" %in% names(ch)) "label" else {
      cand <- grep("^label(::|$)", names(ch), value = TRUE)
      if (length(cand)) cand[1] else NULL
    }
    if (is.null(label_col) || !label_col %in% names(ch)) return(list())
    if (is.na(ln) || !nzchar(ln)) return(list())

    chv <- ch[ch$list_name == ln, , drop = FALSE]
    if (!nrow(chv)) return(list())

    as.list(stats::setNames(as.character(chv[[label_col]]), as.character(chv$name)))
  }

  .is_sm_madre <- function(v) {
    v %in% (ctx$sm_madres %||% character(0)) ||
      v %in% (ctx$vars_sm_madres %||% character(0)) ||
      v %in% (ctx$vars_sm_madres_all %||% character(0)) ||
      v %in% (ctx$vars_diccionario_sm %||% character(0)) ||
      v %in% names(ctx$sm_cols_map %||% list())
  }

  .sm_cols <- function(v) {
    cols <- (ctx$sm_cols_map[[v]] %||% character(0))
    cols <- cols[cols %in% names(data)]
    cols
  }

  # Etiqueta de una dummy: "Pregunta — Opción"
  .label_dummy <- function(col_dummy) {
    # espera: var_madre.code (o var_madre_recod.code si tuvieras eso)
    madre <- sub("\\..*$", "", col_dummy)
    code  <- sub("^.*\\.", "", col_dummy)

    preg <- .obtener_label_var(madre, instrumento, data = NULL)

    map <- .choice_map(madre)
    opt <- as.character(map[[code]] %||% code)

    paste0(preg, " — ", opt)
  }

  # ---------------------------------------------------------------------------
  # Diccionario: variables por sección
  # ---------------------------------------------------------------------------
  dicc_vars_por_seccion <- lapply(ctx$secciones_limpias, function(vs) {
    intersect(vs, ctx$vars_diccionario_all)
  })

  shiny::observe({
    sec <- input$data_seccion
    vars_sec <- dicc_vars_por_seccion[[sec]] %||% character(0)

    if (!length(vars_sec)) {
      shiny::updateSelectInput(session, "dicc_var", choices = c(), selected = NULL)
    } else {
      ch <- stats::setNames(vars_sec, vapply(vars_sec, ctx$label_var, character(1)))
      shiny::updateSelectInput(session, "dicc_var", choices = ch, selected = vars_sec[1])
    }
  })

  output$diccionario_detalle <- shiny::renderUI({
    v <- input$dicc_var

    if (is.null(v) || !nzchar(v) || !v %in% ctx$vars_diccionario_all) {
      return(shiny::div(style="font-size:12px;color:#5f6b7a;", "Sin variables codificadas disponibles."))
    }

    fila <- instrumento$survey[instrumento$survey$name == v, , drop = FALSE]
    tipo_survey <- if (nrow(fila)) tolower(as.character(fila$type[1])) else ""

    es_so <- grepl("^select_one\\b", tipo_survey)
    es_sm <- grepl("^select_multiple\\b", tipo_survey)

    # Etiqueta: para SM usar etiqueta de la madre (aunque no exista como columna)
    etq <- .obtener_label_var(v, instrumento, data = data)

    # Medición: para SO usar attr si existe; para SM forzar NOMINAL elegante
    meas <- if (es_so && v %in% names(data)) attr(data[[v]], "measure", exact = TRUE) else NULL
    meas <- if (!is.null(meas) && nzchar(as.character(meas))) toupper(as.character(meas)) else {
      if (es_sm) "NOMINAL" else "—"
    }

    tipo <- if (es_so) "Selección única" else if (es_sm) "Selección múltiple" else "Variable codificada"

    shiny::tagList(
      shiny::div(class="dicc-kv",
                 shiny::div(class="dicc-k","Variable"), shiny::div(class="dicc-v", v),
                 shiny::div(class="dicc-k","Etiqueta"), shiny::div(class="dicc-v", as.character(etq)),
                 shiny::div(class="dicc-k","Tipo"),     shiny::div(class="dicc-v", tipo),
                 shiny::div(class="dicc-k","Medición"), shiny::div(class="dicc-v", meas)
      ),
      shiny::hr(),
      shiny::div(style="font-size:12px;font-weight:800;color:#002457;margin-bottom:6px;", "Categorías"),
      DT::DTOutput("dicc_opciones")
    )
  })

  output$dicc_opciones <- DT::renderDT({
    v <- input$dicc_var
    if (is.null(v) || !nzchar(v) || !v %in% ctx$vars_diccionario_all) return(NULL)

    fila <- instrumento$survey[instrumento$survey$name == v, , drop = FALSE]
    tipo_survey <- if (nrow(fila)) tolower(as.character(fila$type[1])) else ""
    es_so <- grepl("^select_one\\b", tipo_survey)
    es_sm <- grepl("^select_multiple\\b", tipo_survey)

    ln <- if (nrow(fila) && "list_name" %in% names(fila)) as.character(fila$list_name[1]) else NA_character_
    ch <- instrumento$choices %||% NULL

    opts_df <- NULL
    if (!is.null(ch) && all(c("list_name","name") %in% names(ch)) &&
        !is.na(ln) && nzchar(ln)) {

      label_col <- if ("label" %in% names(ch)) "label" else {
        cand <- grep("^label(::|$)", names(ch), value = TRUE)
        if (length(cand)) cand[1] else NULL
      }

      if (!is.null(label_col) && label_col %in% names(ch)) {
        chv <- ch[ch$list_name == ln, c("name", label_col), drop = FALSE]
        if (nrow(chv)) {
          opts_df <- data.frame(
            Código   = as.character(chv$name),
            Etiqueta = as.character(chv[[label_col]]),
            stringsAsFactors = FALSE
          )
        }
      }
    }

    if (is.null(opts_df) || !nrow(opts_df)) {
      opts_df <- data.frame(Código = character(0), Etiqueta = character(0), stringsAsFactors = FALSE)
    }

    # Ocultar códigos perdidos salvo que se observen (SO o SM)
    cod_perd <- as.character(ctx$codigos_perdidos %||% character(0))
    if (length(cod_perd) > 0 && nrow(opts_df) > 0) {

      vals_obs <- character(0)

      if (es_so && v %in% names(data)) {
        x <- as.character(data[[v]])
        vals_obs <- unique(x[!is.na(x)])

      } else if (es_sm) {

        cols <- .sm_cols(v)
        if (length(cols)) {
          m <- data[, cols, drop = FALSE]
          m <- as.data.frame(lapply(m, function(z) suppressWarnings(as.numeric(as.character(z)))))
          cols_on <- cols[colSums(m == 1, na.rm = TRUE) > 0]
          if (length(cols_on)) {
            choice_codes <- sub(paste0("^", v, "\\."), "", cols_on)
            vals_obs <- unique(choice_codes)
          }
        }
      }

      keep_perd <- if (length(vals_obs)) intersect(cod_perd, vals_obs) else character(0)

      es_perd <- opts_df$Código %in% cod_perd
      opts_df <- opts_df[!es_perd | (opts_df$Código %in% keep_perd), , drop = FALSE]
    }

    DT::datatable(
      opts_df,
      rownames = FALSE,
      options = list(
        paging    = FALSE,
        searching = FALSE,
        info      = FALSE,
        language  = list(search = "Buscar:", zeroRecords = "Sin resultados")
      )
    )
  })

  # ---------------------------------------------------------------------------
  # 🔥 Base de datos: columnas visibles por sección (con expansión SM -> dummies)
  # ---------------------------------------------------------------------------
  vars_data_por_seccion <- lapply(ctx$secciones_limpias, function(vs) {

    vs0 <- intersect(vs, ctx$vars_data_visibles %||% names(data))

    # Expandir: si hay SM madre en la sección, añadir sus dummies (aunque la madre no sea columna)
    sm_madres_sec <- intersect(vs, names(ctx$sm_cols_map %||% list()))
    sm_dummies <- unique(unlist(lapply(sm_madres_sec, .sm_cols), use.names = FALSE))

    # Dejar solo columnas existentes en data
    cols <- unique(c(vs0[vs0 %in% names(data)], sm_dummies))
    cols
  })
  vars_data_por_seccion <- vars_data_por_seccion[vapply(vars_data_por_seccion, length, integer(1)) > 0]

  data_base_filtrada <- shiny::reactive({
    sec  <- input$data_seccion
    cols <- vars_data_por_seccion[[sec]] %||% character(0)

    if (!length(cols)) cols <- head(names(data), 10)

    data[, cols, drop = FALSE]
  })

  data_base_vista <- shiny::reactive({

    df <- data_base_filtrada()
    use_labels <- isTRUE(input$vista_etiquetas)

    if (use_labels) {

      # Primero: valores con labels (SO) si tu helper lo hace
      df2 <- ctx$.to_labels_df(df)

      # Luego: renombrar columnas:
      cn <- vapply(names(df2), function(vcol) {

        # Dummy SM: contiene un punto y su madre está en sm_cols_map
        madre <- sub("\\..*$", "", vcol)
        es_dummy_sm <- grepl("\\.", vcol) && madre %in% names(ctx$sm_cols_map %||% list())

        if (es_dummy_sm) {
          return(.label_dummy(vcol))
        }

        # SO/otras: label attr si existe; sino label del instrumento; sino nombre
        lab <- if (vcol %in% names(data)) attr(data[[vcol]], "label", exact = TRUE) else NULL
        if (!is.null(lab) && nzchar(as.character(lab))) return(as.character(lab))

        .obtener_label_var(vcol, instrumento, data = NULL)
      }, character(1))

      names(df2) <- cn
      return(df2)
    }

    df
  })

  # ---------------------------------------------------------------------------
  # Tabla DT
  # ---------------------------------------------------------------------------
  output$tabla_data <- DT::renderDataTable({

    df <- data_base_vista()
    use_labels <- isTRUE(input$vista_etiquetas)

    col_w <- if (use_labels) 240 else 130

    cb_txt <- paste0(
      "function(settings) {
  var api = this.api();
  var thead = $(api.table().header());

  function escapeRegex(s) {
    return s.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&');
  }

  if ($(thead).find('tr').length < 2) {
    var filterRow = $('<tr class=\"dt-filter-row\">').appendTo(thead);

    api.columns().every(function() {
      var col = this;
      var th  = $('<th>').appendTo(filterRow);

      var uniq = col.data().unique().toArray()
        .filter(function(x){ return x !== null && x !== undefined && x !== ''; });

      uniq.sort();

      if (uniq.length <= 20) {

        var sel = $('<select multiple></select>')
          .css({
            'width':'100%',
            'font-size':'11px',
            'box-sizing':'border-box'
          })
          .appendTo(th);

        $('<option></option>').attr('value','__ALL__').text('(Todos)').appendTo(sel);

        uniq.forEach(function(v){
          $('<option></option>').attr('value', v).text(v).appendTo(sel);
        });

        var $sel = $(sel).selectize({
          plugins: ['remove_button'],
          maxItems: null,
          closeAfterSelect: false,
          hideSelected: false,
          placeholder: 'Filtrar...',
          dropdownParent: 'body',

          render: {
            option: function(item, escape) {
              var label = item.text || item.value;
              var isAll = (item.value === '__ALL__');
              return '<div style=\"display:flex;align-items:center;gap:8px;\">'
                + '<input type=\"checkbox\" style=\"pointer-events:none;\"/>'
                + '<span>' + escape(label) + '</span>'
                + (isAll ? '<span style=\"margin-left:auto;color:#5f6b7a;font-weight:700;\">*</span>' : '')
                + '</div>';
            },
            item: function(item, escape) {
              return '<div>' + escape(item.text || item.value) + '</div>';
            }
          },

          onChange: function(vals) {
            vals = vals || [];

            if (vals.length === 0 || vals.indexOf('__ALL__') >= 0) {
              col.search('').draw();
              return;
            }

            var rx = '^(' + vals.map(escapeRegex).join('|') + ')$';
            col.search(rx, true, false).draw();
          }
        });

        var inst = $sel[0].selectize;
        var $ctrl = $(inst.$control);
        $ctrl.css({
          'border':'1px solid #e6e9f2',
          'border-radius':'10px',
          'min-height':'30px',
          'padding':'2px 4px',
          'box-shadow':'none'
        });

      } else {

        var inp = $('<input type=\"text\" placeholder=\"Filtrar\"/>')
          .css({
            'width':'100%',
            'border':'1px solid #e6e9f2',
            'border-radius':'10px',
            'padding':'6px 8px',
            'font-size':'11px',
            'box-sizing':'border-box'
          })
          .appendTo(th);

        inp.on('keyup change clear', function() {
          if (col.search() !== this.value) {
            col.search(this.value).draw();
          }
        });
      }
    });
  }
}"
    )

    cb <- DT::JS(cb_txt)

    DT::datatable(
      df,
      rownames   = FALSE,
      extensions = c("Scroller"),
      options = list(
        destroy     = TRUE,
        serverSide  = FALSE,
        autoWidth   = FALSE,
        columnDefs  = list(list(width = paste0(col_w, "px"), targets = "_all")),
        deferRender = TRUE,
        scrollX     = TRUE,
        scrollY     = 560,
        scroller    = TRUE,
        pageLength  = 15,
        lengthMenu  = c(10, 15, 25, 50),
        initComplete = cb,
        language = list(
          lengthMenu   = "Mostrando _MENU_ registros",
          search       = "Buscar:",
          info         = "Mostrando _START_ a _END_ de _TOTAL_ registros",
          infoEmpty    = "Mostrando 0 a 0 de 0 registros",
          infoFiltered = "(filtrado de _MAX_ registros)",
          zeroRecords  = "Sin resultados",
          paginate     = list(previous = "Anterior", `next` = "Siguiente")
        )
      )
    )
  })
}
