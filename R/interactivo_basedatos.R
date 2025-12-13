# =============================================================================
# Tab 2: Base de datos (UI + server)
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

    etq <- if (es_so && v %in% names(data)) {
      attr(data[[v]], "label", exact = TRUE) %||% ctx$label_var(v)
    } else {
      .obtener_label_var(v, instrumento, data = NULL)
    }

    meas <- if (es_so && v %in% names(data)) attr(data[[v]], "measure", exact = TRUE) else NULL
    meas <- if (!is.null(meas) && nzchar(as.character(meas))) toupper(as.character(meas)) else "—"

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
    if (!is.null(ch) && all(c("list_name","name","label") %in% names(ch)) &&
        !is.na(ln) && nzchar(ln)) {

      chv <- ch[ch$list_name == ln, c("name","label"), drop = FALSE]
      if (nrow(chv)) {
        opts_df <- data.frame(
          Código   = as.character(chv$name),
          Etiqueta = as.character(chv$label),
          stringsAsFactors = FALSE
        )
      }
    }

    if (is.null(opts_df) || !nrow(opts_df)) {
      opts_df <- data.frame(Código = character(0), Etiqueta = character(0), stringsAsFactors = FALSE)
    }

    cod_perd <- as.character(ctx$codigos_perdidos %||% character(0))
    if (length(cod_perd) > 0 && nrow(opts_df) > 0) {

      vals_obs <- character(0)

      if (es_so && v %in% names(data)) {
        x <- as.character(data[[v]])
        vals_obs <- unique(x[!is.na(x)])

      } else if (es_sm) {
        cols <- ctx$sm_cols_map[[v]] %||% character(0)
        cols <- cols[cols %in% names(data)]
        if (length(cols)) {
          m <- data[, cols, drop = FALSE]
          m <- as.data.frame(lapply(m, function(z) suppressWarnings(as.numeric(as.character(z)))))
          any_one <- apply(m, 1, function(r) any(r == 1, na.rm = TRUE))
          if (any(any_one, na.rm = TRUE)) {
            cols_on <- cols[colSums(m == 1, na.rm = TRUE) > 0]
            choice_codes <- sub(paste0("^", v, "(_recod)?\\."), "", cols_on)
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
        language  = list(
          search      = "Buscar:",
          zeroRecords = "Sin resultados"
        )
      )
    )
  })

  # columnas visibles por sección
  vars_data_por_seccion <- lapply(ctx$secciones_limpias, function(v) {
    intersect(v, ctx$vars_data_visibles)
  })
  vars_data_por_seccion <- vars_data_por_seccion[vapply(vars_data_por_seccion, length, integer(1)) > 0]

  data_base_filtrada <- shiny::reactive({
    sec  <- input$data_seccion
    cols <- vars_data_por_seccion[[sec]] %||% character(0)

    if (!length(cols)) cols <- head(ctx$vars_data_visibles, 10)

    data[, cols, drop = FALSE]
  })

  data_base_vista <- shiny::reactive({
    df <- data_base_filtrada()

    use_labels <- isTRUE(input$vista_etiquetas)

    if (use_labels) {
      df <- ctx$.to_labels_df(df)

      cn <- vapply(names(df), function(v) {
        lab <- attr(data[[v]], "label", exact = TRUE)
        if (!is.null(lab) && nzchar(as.character(lab))) as.character(lab) else v
      }, character(1))
      names(df) <- cn
    }

    df
  })

  output$tabla_data <- DT::renderDataTable({

    df <- data_base_vista()
    use_labels <- isTRUE(input$vista_etiquetas)

    col_w <- if (use_labels) 220 else 120

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
        destroy    = TRUE,
        serverSide = FALSE,
        autoWidth  = FALSE,
        columnDefs = list(list(width = paste0(col_w, "px"), targets = "_all")),
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
