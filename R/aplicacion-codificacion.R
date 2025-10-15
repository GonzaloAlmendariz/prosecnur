# =============================================================================
# ppra_adaptar_data_simple() — 1 salida, sin __tpl, inserta y colorea a la derecha
# SM: <parent>_recod (texto con códigos)
# SO-padre: <parent>_recod
# SO-hijo: <parent>_<alias>_recod (p.ej. CEPR01_why_recod)
# Colores: SM (verde), SO (azul). Solo se colorea la columna objetivo.
# =============================================================================

suppressPackageStartupMessages({
  library(readxl); library(dplyr); library(stringr); library(purrr); library(tidyr); library(tibble)
})

`%||%` <- function(a,b) if (is.null(a) || (length(a)==1 && is.na(a))) b else a
nz      <- function(x) !is.na(x) & nzchar(trimws(as.character(x)))

# -------- utilidades básicas --------
pick_join_key <- function(df){
  cands <- c("_uuid","uuid","meta_instance_id","instanceid","_id",
             "_index","Codigo pulso","Código pulso","Pulso_code","pulso_code")
  hit <- cands[cands %in% names(df)]
  if (length(hit)) hit[1] else NA_character_
}

.norm01 <- function(v){
  if (is.numeric(v)) return(ifelse(is.na(v), NA_integer_, ifelse(v!=0,1L,0L)))
  s <- tolower(trimws(as.character(v)))
  s <- iconv(s, from="", to="ASCII//TRANSLIT")
  out <- rep(NA_integer_, length(s))
  out[s %in% c("1","t","true","si","s","yes","y","verdadero")] <- 1L
  out[s %in% c("0","f","false","no","n","falso")]               <- 0L
  out
}

.normcode <- function(x){
  x <- trimws(as.character(x))
  x <- iconv(x, from="", to="ASCII//TRANSLIT")
  tolower(x)
}

leer_datos_generico <- function(path_data, sheet = NULL){
  ext <- tolower(tools::file_ext(path_data))
  if (ext %in% c("csv","txt")) {
    suppressWarnings(readr::read_csv(path_data, show_col_types = FALSE))
  } else {
    readxl::read_excel(path_data, sheet = sheet %||% 1)
  }
}

# -------- plantilla: resolver hoja --------
.ppra_clean <- function(x) gsub("[^A-Za-z0-9]+", "", tolower(trimws(as.character(x))))
ppra_resolve_template_sheet <- function(path_plantilla, parent_var){
  sh <- readxl::excel_sheets(path_plantilla)
  hit <- which(sh == parent_var); if (length(hit)) return(sh[hit[1]])
  hit <- which(tolower(sh) == tolower(parent_var)); if (length(hit)) return(sh[hit[1]])
  p31 <- substr(parent_var, 1, 31)
  hit <- which(sh == p31 | tolower(sh) == tolower(p31)); if (length(hit)) return(sh[hit[1]])
  cl_parent <- .ppra_clean(parent_var); cl_sheets <- .ppra_clean(sh)
  hit <- which(cl_sheets == cl_parent); if (length(hit)) return(sh[hit[1]])
  d <- adist(cl_parent, cl_sheets); j <- which.min(d)
  if (length(j) && is.finite(d[j]) && d[j] <= 5) return(sh[j])
  NA_character_
}

# -------- XLSForm: choices por parent --------
ppra_get_choices_parent <- function(path_instrumento, parent_var){
  survey  <- readxl::read_excel(path_instrumento, sheet = "survey")
  choices <- readxl::read_excel(path_instrumento, sheet = "choices")

  i <- which(survey$name == parent_var)
  if (!length(i)) {
    clean <- function(x) gsub("[^A-Za-z0-9_]+","", as.character(x))
    i <- which(clean(as.character(survey$name)) == clean(parent_var))
  }
  if (!length(i)) stop("No encontré la pregunta en survey: ", parent_var)

  ty <- as.character(survey$type[i][1] %||% "")
  ln <- trimws(sub("^\\S+\\s+", "", ty))
  if (!nzchar(ln)) stop("No pude determinar list_name desde 'type' para: ", parent_var)

  nmsl <- tolower(names(choices))
  label_es_col <- names(choices)[match(TRUE, nmsl %in% c(
    "label::spanish (es)","label::spanish(es)","label::spanish_es",
    "label_spanish_es","label::spanish","label::es"
  ))]
  if (is.na(label_es_col)) label_es_col <- if ("label" %in% names(choices)) "label" else NA_character_

  ch <- choices[ trimws(choices$list_name) == trimws(ln), , drop = FALSE ]
  if (!nrow(ch)) stop("El list_name '", ln, "' no tiene choices en 'choices'.")

  tibble::tibble(
    order = seq_len(nrow(ch)),
    code  = as.character(ch$name),
    label = if (!is.na(label_es_col)) as.character(ch[[label_es_col]]) else as.character(ch$name)
  )
}

# ====================== SM ======================
# Construye SOLO <parent>_recod con:
# - clásico: Parent/code
# - overrides plantilla: Parent/code_recod (0/1) o _RECOD
# - nuevas hijas: ^Parent/Nueva_recod -> integradas
# - normalización de códigos (acentos/caso) y “Other” unificado/filtrado
ppra_sm_parent_recod <- function(df, parent, path_instrumento, path_plantilla){
  nms <- names(df)

  # catálogo clásico
  ch <- ppra_get_choices_parent(path_instrumento, parent)
  classic        <- as.character(ch$code)
  classic_norm   <- .normcode(classic)

  # hoja de plantilla (opcional)
  tpl <- NULL
  if (file.exists(path_plantilla)) {
    sheet <- ppra_resolve_template_sheet(path_plantilla, parent)
    if (!is.na(sheet)) {
      tpl <- readxl::read_excel(path_plantilla, sheet = sheet)
      names(tpl) <- trimws(names(tpl))
    }
  }

  kd <- pick_join_key(df)
  if (is.null(kd) || is.na(kd)) stop("No se encontró clave de unión en datos para SM: ", parent)

  # “tmp” con columnas de plantilla (sin dejar __tpl al final)
  if (!is.null(tpl)) {
    kt <- pick_join_key(tpl); if (is.null(kt) || is.na(kt)) stop("Plantilla sin clave para ", parent)
    names(tpl)[names(tpl)==kt] <- kd
    tmp <- left_join(df[, kd, drop = FALSE], tpl, by = kd)
  } else {
    tmp <- df[, kd, drop = FALSE]
  }
  tnames <- names(tmp)

  # matriz clásico desde crudo
  mat <- matrix(NA_integer_, nrow = nrow(df), ncol = length(classic))
  colnames(mat) <- classic
  for (code in classic) {
    raw_col <- paste0(parent, "/", code)
    if (raw_col %in% nms) {
      v <- .norm01(df[[raw_col]])
      mat[, code] <- v
    }
  }

  # overrides clásico desde plantilla
  if (!is.null(tpl)) {
    for (code in classic) {
      rc1 <- paste0(parent, "/", code, "_RECOD")
      rc2 <- paste0(parent, "/", code, "_recod")
      if (rc1 %in% tnames || rc2 %in% tnames) {
        v <- if (rc1 %in% tnames) tmp[[rc1]] else tmp[[rc2]]
        v <- .norm01(v)
        w1 <- which(!is.na(v) & v==1L); if (length(w1)) mat[w1, code] <- 1L
        w0 <- which(!is.na(v) & v==0L); if (length(w0)) mat[w0, code] <- 0L
      }
    }
  }

  # nuevas hijas en plantilla: si su “base” normalizada coincide con un clásico,
  # trata como override del clásico; si no, añádela como token nuevo.
  new_tokens <- vector("list", nrow(df))
  if (!is.null(tpl)) {
    rx_new <- paste0("^", parent, "/[^/]+_(?i:recod)$")
    new_cols <- tnames[grepl(rx_new, tnames, perl = TRUE)]
    if (length(new_cols)) {
      for (cc in new_cols) {
        base <- sub("^.+/", "", cc)
        base <- sub("_(?i:recod)$", "", base)
        base_norm <- .normcode(base)

        v <- .norm01(tmp[[cc]])

        if (length(classic) && base_norm %in% classic_norm) {
          # mapear a su código clásico canónico
          can_code <- classic[match(base_norm, classic_norm)]
          w1 <- which(!is.na(v) & v==1L); if (length(w1)) mat[w1, can_code] <- 1L
          w0 <- which(!is.na(v) & v==0L); if (length(w0)) mat[w0, can_code] <- 0L
        } else {
          # es realmente una NUEVA categoría
          sel <- which(!is.na(v) & v==1L)
          if (length(sel)) {
            for (i in sel) new_tokens[[i]] <- unique(c(new_tokens[[i]], base))
          }
        }
      }
    }
  }

  # construir tokens finales (clásicos + nuevas)
  parent_recod <- character(nrow(df))
  for (i in seq_len(nrow(df))){
    from_classic <- colnames(mat)[which(mat[i, ] == 1L)]
    from_new     <- new_tokens[[i]] %||% character(0)

    # si “Other” clásico está 0 por override, eliminar cualquier variante en nuevos
    other_idx <- match("other", .normcode(from_classic))
    if (is.na(other_idx)) {
      # quizá llegó por “new” en otra grafía: si el clásico Other está en 0, bórralo de from_new
      if ("Other" %in% classic) {
        # ¿está en 0 explícito?
        j <- which(colnames(mat)=="Other")
        if (length(j) && !is.na(mat[i,j]) && mat[i,j]==0L) {
          from_new <- from_new[ .normcode(from_new) != "other" ]
        }
      }
    }

    allcodes <- unique(c(from_classic, from_new))
    parent_recod[i] <- if (length(allcodes)) paste(allcodes, collapse = " ") else NA_character_
  }

  out_col <- paste0(parent, "_recod")
  df[[out_col]] <- parent_recod

  list(df = df, new_col = out_col)
}

# ====================== SO (padre) ======================
ppra_so_parent <- function(df, parent, path_instrumento, path_plantilla){
  ch <- ppra_get_choices_parent(path_instrumento, parent)
  cat_codes <- as.character(ch$code); cat_labs <- as.character(ch$label)
  lab_by_code <- function(code){ i <- match(code, cat_codes); ifelse(is.na(i), NA_character_, cat_labs[i]) }

  tpl <- NULL
  if (file.exists(path_plantilla)) {
    sheet <- ppra_resolve_template_sheet(path_plantilla, parent)
    if (!is.na(sheet)) {
      tpl <- readxl::read_excel(path_plantilla, sheet = sheet)
      names(tpl) <- trimws(names(tpl))
    }
  }
  kd <- pick_join_key(df); stopifnot(!is.na(kd))
  tmp <- if (!is.null(tpl)) {
    kt <- pick_join_key(tpl); stopifnot(!is.na(kt))
    names(tpl)[names(tpl)==kt] <- kd
    left_join(df[, kd, drop = FALSE], tpl, by = kd)
  } else df[, kd, drop = FALSE]
  tnames <- names(tmp)
  .pick <- function(cands) { h <- cands[cands %in% tnames]; if (length(h)) h[1] else NA_character_ }

  rec_tpl  <- .pick(c(paste0(parent,"_RECOD"), paste0(parent,"_recod")))
  labr_tpl <- .pick(c(paste0(parent,"_LABEL_RECOD"), paste0(parent,"_label_recod")))
  lab_tpl  <- .pick(c(paste0(parent,"_LABEL"),       paste0(parent,"_label")))
  orec_tpl <- .pick(c(paste0(parent,"_OTHER_RECOD"), paste0(parent,"_other_recod")))
  olab_tpl <- .pick(c(paste0(parent,"_OTHER_RECOD_LABEL"), paste0(parent,"_other_recod_label")))

  base_code <- as.character(df[[parent]]); base_code[base_code==""] <- NA_character_
  code_final <- base_code
  if (!is.na(rec_tpl)) {
    x <- trimws(as.character(tmp[[rec_tpl]])); x[x==""] <- NA_character_
    i <- which(!is.na(x)); if (length(i)) code_final[i] <- x[i]
  }
  is_other <- tolower(code_final) == "other" | tolower(base_code) == "other"
  if (any(is_other) && !is.na(orec_tpl)) {
    y <- trimws(as.character(tmp[[orec_tpl]])); y[y==""] <- NA_character_
    j <- which(is_other & !is.na(y)); if (length(j)) code_final[j] <- y[j]
  }

  label_final <- rep(NA_character_, nrow(df))
  if (!is.na(labr_tpl)) {
    z <- trimws(as.character(tmp[[labr_tpl]])); z[z==""] <- NA_character_
    label_final[!is.na(z)] <- z[!is.na(z)]
  }
  need <- is.na(label_final) & !is.na(code_final)
  if (any(need)) label_final[need] <- lab_by_code(code_final[need])
  need <- is.na(label_final)
  if (any(need) && !is.na(lab_tpl)) {
    w <- trimws(as.character(tmp[[lab_tpl]])); w[w==""] <- NA_character_
    label_final[need & !is.na(w)] <- w[need & !is.na(w)]
  }
  need <- is.na(label_final)
  if (any(need) && !is.na(olab_tpl)) {
    q <- trimws(as.character(tmp[[olab_tpl]])); q[q==""] <- NA_character_
    label_final[need & !is.na(q)] <- q[need & !is.na(q)]
  }

  out_col <- paste0(parent, "_recod")
  df[[out_col]] <- code_final
  lab_col <- paste0(parent, "_recod_label")
  if (any(nz(label_final))) df[[lab_col]] <- label_final

  list(df = df, new_col = out_col)
}

# ====================== SO (hijo) ======================
ALIAS_CHILD <- c(
  "why","por_que","porque","motivo","razon","razón","motivo_principal","motivo_ppal","causa",
  "other","otro","especificar","especifique","otro_especificar","otro_cual","cual"
)

resolve_child_recod_col <- function(parent, colnames_tpl){
  nm <- trimws(colnames_tpl)
  nm_norm <- tolower(gsub("[^a-z0-9_]+","_", iconv(nm,"","ASCII//TRANSLIT")))
  parent_norm <- tolower(gsub("[^a-z0-9_]+","_", iconv(parent,"","ASCII//TRANSLIT")))
  aliases_norm <- tolower(gsub("[^a-z0-9_]+","_", iconv(ALIAS_CHILD,"","ASCII//TRANSLIT")))

  # candidatos exactos parent_alias_recod
  cand <- paste0(parent_norm, "_", aliases_norm, "_recod")
  hit <- which(nm_norm %in% cand)
  if (length(hit)) return(nm[hit[1]])

  # regex tolerante
  rx <- paste0("^", parent_norm, "(_)?(", paste(aliases_norm, collapse="|"), ")_recod$")
  j <- which(grepl(rx, nm_norm, perl = TRUE))
  if (length(j)) nm[j[1]] else NA_character_
}

ppra_so_child <- function(df, parent, path_plantilla){
  tpl <- NULL
  if (file.exists(path_plantilla)) {
    sheet <- ppra_resolve_template_sheet(path_plantilla, parent)
    if (!is.na(sheet)) {
      tpl <- readxl::read_excel(path_plantilla, sheet = sheet)
      names(tpl) <- trimws(names(tpl))
    }
  }
  if (is.null(tpl)) {
    message("[SO-hijo] No hay hoja para '", parent, "'. Omito.")
    return(list(df=df, new_col=character(0)))
  }
  kd <- pick_join_key(df); kt <- pick_join_key(tpl)
  if (is.na(kd) || is.na(kt)) {
    message("[SO-hijo] Sin clave de unión para '", parent, "'.")
    return(list(df=df, new_col=character(0)))
  }
  names(tpl)[names(tpl)==kt] <- kd

  src <- resolve_child_recod_col(parent, names(tpl))
  if (is.na(src)) {
    message("[SO-hijo] No hallé hija *_recod (why/otro) para '", parent, "'.")
    return(list(df=df, new_col=character(0)))
  }

  # nombre de salida estandarizado
  alias <- sub(paste0("^", parent, "_"), "", sub("_recod$","", src))
  out_col <- paste0(parent, "_", alias, "_recod")
  if (!(out_col %in% names(df))) df[[out_col]] <- NA_character_

  tmp <- left_join(df[, c(kd), drop = FALSE], tpl[, c(kd, src), drop = FALSE], by = kd)
  val <- trimws(as.character(tmp[[src]])); val[val==""] <- NA_character_
  i <- which(!is.na(val)); if (length(i)) df[[out_col]][i] <- val[i]

  list(df = df, new_col = out_col)
}

# ====================== Insertar a la derecha + colorear ======================
insert_right_of <- function(df, anchor, cols_to_insert){
  cols_to_insert <- cols_to_insert[cols_to_insert %in% names(df)]
  if (!length(cols_to_insert)) return(df)
  nms <- names(df)
  apos <- match(anchor, nms); if (is.na(apos)) apos <- 0L
  base <- setdiff(nms, cols_to_insert)
  left <- if (apos<=0L) character(0) else base[seq_len(apos)]
  right<- if (apos<=0L) base else base[(apos+1):length(base)]
  df[, c(left, cols_to_insert, right), drop = FALSE]
}

ppra_color_columns <- function(df, path, sm_cols = character(0), sop_cols = character(0), soh_cols = character(0)){
  if (is.null(path)) return(invisible(NULL))
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    warning("No se encontró 'openxlsx'. No se exporta Excel.")
    return(invisible(NULL))
  }
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "DATA")
  openxlsx::writeData(wb, "DATA", df)

  verde <- "#DFF5DF" # SM
  azul  <- "#DCEBFF" # SO (padre e hijo)

  paint <- function(cols, color){
    if (!length(cols)) return()
    idx <- match(cols, names(df))
    idx <- idx[!is.na(idx)]
    if (!length(idx)) return()
    openxlsx::addStyle(wb, "DATA", openxlsx::createStyle(fgFill = color),
                       rows = 1:(nrow(df)+1), cols = idx, gridExpand = TRUE, stack = TRUE)
  }

  paint(unique(sm_cols),  verde)
  paint(unique(sop_cols), azul)
  paint(unique(soh_cols), azul)

  openxlsx::freezePane(wb, "DATA", firstActiveRow = 2)
  openxlsx::setColWidths(wb, "DATA", cols = 1:ncol(df), widths = "auto")
  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
  invisible(path)
}

# ====================== FUNCIÓN PRINCIPAL ======================
#' @title ppra_adaptar_data
#' @description
#' Función **limpia y universal**:
#' - **SM**: genera `<parent>_recod` (con overrides y nuevas hijas, normalización de códigos),
#'   lo inserta a la derecha del `parent` y lo colorea **verde**.
#' - **SO-padre**: genera `<parent>_recod`, lo inserta y colorea **azul**.
#' - **SO-hijo**: genera `<parent>_<alias>_recod` (p.ej. `CEPR01_why_recod`),
#'   lo inserta y colorea **azul**.
#' No deja columnas técnicas ni requiere “slim”.
#'
#' @param path_instrumento Ruta al XLSForm (survey/choices)
#' @param path_datos Ruta a los datos (CSV/XLSX)
#' @param path_plantilla Ruta a la plantilla de codificación (hojas por parent)
#' @param sm_vars Vector de SM (p.ej. c("services","futuros_servicios"))
#' @param so_parent_vars Vector de SO que recodifican el **padre**
#' @param so_child_vars Vector de SO que recodifican el **hijo** (why/otro)
#' @param out_path Ruta del XLSX de salida
#' @return data.frame adaptado (también guardado en `out_path` si se provee)
#' @export
ppra_adaptar_data <- function(path_instrumento,
                                     path_datos,
                                     path_plantilla,
                                     sm_vars        = character(0),
                                     so_parent_vars = character(0),
                                     so_child_vars  = character(0),
                                     out_path       = NULL){

  stopifnot(file.exists(path_instrumento), file.exists(path_datos), file.exists(path_plantilla))

  df <- leer_datos_generico(path_datos)

  sm_cols_to_color  <- character(0)
  sop_cols_to_color <- character(0)
  soh_cols_to_color <- character(0)

  # SM
  if (length(sm_vars)) {
    for (pv in sm_vars) {
      res <- ppra_sm_parent_recod(df, pv, path_instrumento, path_plantilla)
      df  <- res$df
      df  <- insert_right_of(df, pv, res$new_col)
      sm_cols_to_color <- c(sm_cols_to_color, res$new_col)
    }
  }

  # SO padre
  if (length(so_parent_vars)) {
    for (pv in so_parent_vars) {
      res <- ppra_so_parent(df, pv, path_instrumento, path_plantilla)
      df  <- res$df
      df  <- insert_right_of(df, pv, res$new_col)
      sop_cols_to_color <- c(sop_cols_to_color, res$new_col)
    }
  }

  # SO hijo
  if (length(so_child_vars)) {
    for (pv in so_child_vars) {
      res <- ppra_so_child(df, pv, path_plantilla)
      df  <- res$df
      if (length(res$new_col)) {
        df  <- insert_right_of(df, pv, res$new_col)
        soh_cols_to_color <- c(soh_cols_to_color, res$new_col)
      }
    }
  }

  # Export + color
  if (!is.null(out_path)) {
    ppra_color_columns(df, out_path,
                       sm_cols  = sm_cols_to_color,
                       sop_cols = sop_cols_to_color,
                       soh_cols = soh_cols_to_color)
  }

  df
}
