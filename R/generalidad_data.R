#' Adaptar la base para reporte y estructura tipo SPSS
#'
#' `reporte_data()` toma una base de datos ya adaptada (tras la evaluación de
#' consistencia y la recodificación/adaptación del instrumento y la data) y un
#' objeto de instrumento creado con [reporte_instrumento()], y realiza tres
#' tareas principales:
#' \enumerate{
#'   \item Asigna metadatos: etiquetas de variables (`label`), etiquetas de
#'         valor (`labels`) y nivel de medición sugerido (`measure`).
#'   \item Adapta las variables de elección única (`select_one`) para que
#'         utilicen códigos en lugar de etiquetas, solo cuando detecta que en
#'         la base hay valores que coinciden con las etiquetas del instrumento.
#'         Los códigos ya presentes se conservan sin modificación.
#'   \item Normaliza las variables derivadas de `select_multiple`:
#'         dummies en 0/1, con `label` de la opción, `labels` = `"No"/"Sí"`,
#'         y nombres `var.codigo` (punto) en lugar de `var/opción`, tal como
#'         se espera en un flujo orientado a SPSS. La variable madre se
#'         reconstruye solo si no existe ya en la base.
#' }
#'
#' La función no escribe ningún archivo ni crea objetos `labelled_spss`. La
#' conversión final a objetos SPSS (por ejemplo con [haven::labelled_spss()])
#' debe realizarse en una etapa posterior (por ejemplo, en una función
#' `reporte_spss()`).
#'
#' @param data Un `data.frame` o `tibble` que contenga la base de datos ya
#'   adaptada (nombres finales de variables) tras las fases de consistencia y
#'   recodificación.
#' @param instrumento Objeto devuelto por [reporte_instrumento()], que
#'   contiene los metadatos del XLSForm (survey, choices, diccionarios, etc.).
#' @param var_peso (Opcional) Nombre de la variable de peso en `data`. Si se
#'   proporciona y existe en la base, se registra como atributo `var_peso`.
#' @param dummy_vars (Opcional) Vector con nombres adicionales de variables
#'   dummy (0/1) a las que se desee asignar etiquetas de valor genéricas
#'   `"No"` / `"Sí"`. Todas las variables cuyo nombre contenga `"/"` se tratan
#'   automáticamente como dummies de `select_multiple`, por lo que este
#'   argumento es únicamente complementario.
#' @param dummies_na_to_zero Lógico; si `TRUE` (por defecto), las dummies de
#'   `select_multiple` se normalizan de modo que los `NA` se imputan a 0.
#'   Si `FALSE`, se dejan como `NA` cuando no hay información.
#'
#' @return El mismo objeto `data` pero:
#'   \itemize{
#'     \item Con variables `select_one` en códigos cuando se detectan labels
#'           provenientes del instrumento.
#'     \item Con dummies 0/1 renombradas `var.codigo` (punto) y etiquetadas con
#'           `label` (texto de la opción) y `labels` = `"No"` / `"Sí"`.
#'     \item Con variables madre de `select_multiple` reconstruidas como
#'           cadenas de códigos separados por `";"` cuando no existían en la
#'           base original.
#'     \item Con atributos `label`, `labels` y `measure` según el instrumento.
#'     \item Con atributos a nivel de objeto:
#'       \describe{
#'         \item{instrumento_reporte}{Metadatos completos del instrumento.}
#'         \item{var_peso}{Nombre de la variable de peso (si se proporcionó).}
#'         \item{vars_fecha}{Variables declaradas como fecha en el instrumento.}
#'         \item{vars_hora}{Variables declaradas como hora.}
#'         \item{vars_datetime}{Variables declaradas como fecha-hora.}
#'       }
#'     \item Con clase adicional `"prosecnur_reporte_tbl"` para facilitar su uso
#'           en otras funciones de reporte.
#'   }
#'
#' @export
reporte_data <- function(data,
                         instrumento,
                         var_peso           = NULL,
                         dummy_vars         = NULL,
                         dummies_na_to_zero = TRUE) {

  if (!is.data.frame(data)) {
    stop("`data` debe ser un data.frame o tibble.", call. = FALSE)
  }

  if (is.null(instrumento) || !is.list(instrumento)) {
    stop("`instrumento` debe ser un objeto lista devuelto por `reporte_instrumento()`.",
         call. = FALSE)
  }

  if (!inherits(instrumento, "prosecnur_instrumento")) {
    warning("`instrumento` no tiene clase 'prosecnur_instrumento'. ",
            "Se usará igualmente, asumiendo la estructura esperada.",
            call. = FALSE)
  }

  survey             <- instrumento$survey
  choices            <- instrumento$choices
  var_labels         <- instrumento$var_labels
  measure_rules      <- instrumento$measure_rules
  dicc_label_to_code <- instrumento$dicc_label_to_code
  dicc_code_to_label <- instrumento$dicc_code_to_label

  # -------------------------------------------------------------------------
  # Helpers internos
  # -------------------------------------------------------------------------
  clean_dummy_name <- function(x) {
    base <- gsub("/", ".", x)
    base <- iconv(base, from = "", to = "ASCII//TRANSLIT")
    base <- tolower(base)
    base <- gsub(" ", ".", base)
    base <- gsub("[^a-z0-9._]", "_", base)
    base <- gsub("_+", "_", base)
    base <- gsub("\\.+", ".", base)
    base <- gsub("^[_\\.]+|[_\\.]+$", "", base)
    base
  }

  canon_txt <- function(x) {
    x <- as.character(x)
    x <- tolower(x)
    x <- iconv(x, from = "", to = "ASCII//TRANSLIT")
    x <- gsub("[^a-z0-9]+", " ", x)
    x <- trimws(x)
    x <- gsub("\\s+", " ", x)
    x
  }

  # -------------------------------------------------------------------------
  # 1) Recode select_one: de labels a códigos cuando haya labels presentes
  # -------------------------------------------------------------------------
  if (!is.null(survey) &&
      all(c("name", "type", "list_name") %in% names(survey)) &&
      !is.null(dicc_label_to_code)) {

    so_vars <- survey$name[grepl("^select_one", survey$type)]
    so_vars <- intersect(so_vars, names(data))

    for (v in so_vars) {
      ln <- survey$list_name[survey$name == v][1]
      if (is.na(ln) || !ln %in% names(dicc_label_to_code)) next

      map_lab_to_code <- dicc_label_to_code[[ln]]   # labels -> codes (char)
      if (is.null(map_lab_to_code) || length(map_lab_to_code) == 0L) next

      codes_vec  <- as.character(unname(map_lab_to_code))
      labels_vec <- as.character(names(map_lab_to_code))

      x_chr <- as.character(data[[v]])
      vals  <- unique(trimws(x_chr[!is.na(x_chr) & x_chr != ""]))

      if (!length(vals)) next

      # ¿hay al menos algún valor que coincida con un label del diccionario?
      prop_labels <- mean(vals %in% labels_vec)
      prop_codes  <- mean(vals %in% codes_vec)  # se mantiene por si luego se usa

      # Si hay labels presentes, recodificamos esos labels a código.
      # Los códigos existentes se mantienen porque .default = x_chr
      if (prop_labels > 0) {
        x_rec <- dplyr::recode(
          x_chr,
          !!!map_lab_to_code,
          .default = x_chr
        )
        data[[v]] <- x_rec
      }
    }
  }

  # -------------------------------------------------------------------------
  # 2) Tratamiento de select_multiple: dummies 0/1, renombrar y madre (si falta)
  # -------------------------------------------------------------------------
  if (!is.null(survey) && all(c("name", "type", "list_name") %in% names(survey))) {

    sm_vars <- survey$name[grepl("^select_multiple", survey$type)]

    for (v in sm_vars) {

      # dummies detectadas por patrón var/... (estado previo a limpiar nombres)
      dummy_cols <- grep(paste0("^", v, "/"), names(data), value = TRUE)
      if (length(dummy_cols) == 0L) next

      # 2.1 Normalizar a 0/1 (con control de NA -> 0 o NA)
      data[dummy_cols] <- lapply(data[dummy_cols], function(col) {
        x_num <- suppressWarnings(as.numeric(col))
        res   <- rep(NA_real_, length(x_num))

        res[!is.na(x_num) & x_num == 1] <- 1
        res[!is.na(x_num) & x_num != 1] <- 0

        if (dummies_na_to_zero) {
          res[is.na(x_num)] <- 0
        }
        res
      })

      # 2.2 Etiquetar dummies con label de la opción (si choices está disponible)
      if (!is.null(choices)) {
        ln <- survey$list_name[survey$name == v][1]
        if (!is.na(ln)) {
          for (d in dummy_cols) {
            code_opt_raw <- sub("^[^/]+/", "", d)   # sufijo después de "/"

            lab_opt <- NA_character_

            # a) si tenemos dicc_label_to_code podemos intentar canonizar label
            if (!is.null(dicc_label_to_code[[ln]])) {
              dict_lc   <- dicc_label_to_code[[ln]]  # label -> code
              labs_orig <- names(dict_lc)

              suf_canon  <- canon_txt(code_opt_raw)
              labs_canon <- canon_txt(labs_orig)

              idx <- match(suf_canon, labs_canon)
              if (!is.na(idx)) {
                code_match <- dict_lc[[labs_orig[idx]]]
                lab_opt <- choices$label[
                  choices$list_name == ln &
                    as.character(choices$name) == as.character(code_match)
                ][1]
              }
            }

            # b) si no encontramos nada, intento directo por nombre en choices
            if (is.na(lab_opt) || is.null(lab_opt)) {
              lab_opt <- choices$label[
                choices$list_name == ln &
                  canon_txt(choices$name) == canon_txt(code_opt_raw)
              ][1]
            }

            if (!is.na(lab_opt) && !is.null(lab_opt)) {
              attr(data[[d]], "label") <- as.character(lab_opt)
            }
          }
        }
      }

      # 2.3 Renombrar dummies usando CODE y no LABEL en la parte después de "/"
      if (!is.null(dicc_label_to_code)) {
        ln <- survey$list_name[survey$name == v][1]
        if (!is.na(ln) && ln %in% names(dicc_label_to_code)) {
          dict_lc    <- dicc_label_to_code[[ln]]  # label -> code
          labs_orig  <- names(dict_lc)
          labs_canon <- canon_txt(labs_orig)

          for (d in dummy_cols) {
            suf_raw   <- sub("^.+/", "", d)
            suf_canon <- canon_txt(suf_raw)

            idx <- match(suf_canon, labs_canon)
            if (!is.na(idx)) {
              code_opt     <- dict_lc[[labs_orig[idx]]]
              nuevo_nombre <- paste0(v, "/", code_opt)
              names(data)[names(data) == d] <- nuevo_nombre
            }
          }
        }
      }

      # actualizar dummy_cols después de posibles renombres
      dummy_cols <- grep(paste0("^", v, "/"), names(data), value = TRUE)
      if (length(dummy_cols) == 0L) next

      # 2.4 Reconstruir variable madre (códigos separados por ";")
      #     SOLO si no existe ya en `data`
      if (!v %in% names(data)) {
        codigos <- sub(paste0("^", v, "/"), "", dummy_cols)
        mat <- as.matrix(data[, dummy_cols, drop = FALSE])

        madre_vec <- apply(mat, 1, function(row) {
          sel <- row == 1
          if (!any(sel)) return(NA_character_)
          paste(codigos[sel], collapse = ";")
        })

        # insertar madre al inicio de las dummies
        pos_dummy1 <- which(names(data) == dummy_cols[1])[1]
        data[[v]] <- madre_vec

        if (!is.na(pos_dummy1) && pos_dummy1 > 1) {
          data <- dplyr::relocate(
            data,
            dplyr::all_of(v),
            .before = dplyr::all_of(dummy_cols[1])
          )
        }
      }

      # 2.5 Asignar labels "No"/"Sí" a dummies
      for (d in dummy_cols) {
        attr(data[[d]], "labels") <- c(`0` = "No", `1` = "Sí")
      }
    }
  }

  # -------------------------------------------------------------------------
  # 3) Asignar etiquetas de variable (attr(, "label"))
  # -------------------------------------------------------------------------
  if (!is.null(var_labels) && length(var_labels) > 0L) {
    vars_comunes <- intersect(names(data), names(var_labels))
    for (v in vars_comunes) {
      attr(data[[v]], "label") <- as.character(var_labels[[v]])
    }
  }

  # -------------------------------------------------------------------------
  # 4) Asignar value-labels para select_one (attr(, "labels"))
  # -------------------------------------------------------------------------
  if (!is.null(survey) && all(c("name", "type", "list_name") %in% names(survey))) {

    so_vars <- unique(survey$name[grepl("^select_one", survey$type)])
    so_comunes <- intersect(so_vars, names(data))

    for (v in so_comunes) {
      ln <- survey$list_name[survey$name == v][1]
      if (is.na(ln) || is.null(ln)) next
      if (is.null(dicc_code_to_label) || !ln %in% names(dicc_code_to_label)) next

      labs <- dicc_code_to_label[[ln]]  # code -> label
      if (is.null(labs) || length(labs) == 0L) next

      labs <- stats::setNames(as.character(labs),
                              nm = as.character(names(labs)))
      attr(data[[v]], "labels") <- labs
    }
  }

  # -------------------------------------------------------------------------
  # 5) Dummies adicionales declaradas en dummy_vars (sin "/")
  # -------------------------------------------------------------------------
  if (!is.null(dummy_vars) && length(dummy_vars) > 0L) {
    dummy_add <- intersect(dummy_vars, names(data))
    if (length(dummy_add) > 0L) {
      for (d in dummy_add) {
        attr(data[[d]], "labels") <- c(`0` = "No", `1` = "Sí")
      }
    }
  }

  # -------------------------------------------------------------------------
  # 6) Asignar nivel de medición sugerido (attr(, "measure"))
  # -------------------------------------------------------------------------
  if (!is.null(measure_rules) &&
      all(c("name", "measure_sugerida") %in% names(measure_rules))) {

    vars_comunes_m <- intersect(names(data), measure_rules$name)

    for (v in vars_comunes_m) {
      m <- measure_rules$measure_sugerida[measure_rules$name == v][1]
      if (!is.na(m) && !is.null(m) && nzchar(m)) {
        attr(data[[v]], "measure") <- as.character(m)
      }
    }
  }

  # -------------------------------------------------------------------------
  # 7) Renombrar dummies con "/" a nombres compatibles SPSS (".")
  # -------------------------------------------------------------------------
  dummy_idx <- grepl("/", names(data))
  if (any(dummy_idx)) {
    names(data)[dummy_idx] <- clean_dummy_name(names(data)[dummy_idx])
  }

  # -------------------------------------------------------------------------
  # 8) Registrar variable de peso y variables temporales a nivel de objeto
  # -------------------------------------------------------------------------
  if (!is.null(var_peso)) {
    if (!var_peso %in% names(data)) {
      warning("`var_peso = ", var_peso, "` no se encuentra en `data`. ",
              "Se registrará igual en el atributo, pero la variable no existe.",
              call. = FALSE)
    }
    attr(data, "var_peso") <- var_peso
  }

  attr(data, "vars_fecha")    <- instrumento$vars_fecha
  attr(data, "vars_hora")     <- instrumento$vars_hora
  attr(data, "vars_datetime") <- instrumento$vars_datetime

  attr(data, "instrumento_reporte") <- instrumento

  # -------------------------------------------------------------------------
  # 9) Añadir clase para identificar este objeto en el resto del flujo
  # -------------------------------------------------------------------------
  class(data) <- c("prosecnur_reporte_tbl", class(data))

  data
}
