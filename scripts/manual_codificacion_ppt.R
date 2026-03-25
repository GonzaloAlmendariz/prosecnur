# =============================================================================
# Manual de Capacitación PPT — Codificación de Preguntas Abiertas (PPRA)
# =============================================================================
# Genera un PowerPoint instruccional usando officer directamente.
# Incluye diagramas explicativos generados con ggplot2.
# Requiere: officer, ggplot2
#
# USO:
#   1. Ejecutar este script para generar el PPT.
#   2. Las slides con screenshots usarán imágenes de scripts/screenshots/.
#      Si no existen, se generan placeholders automáticos.
#   3. Los diagramas se generan automáticamente en scripts/diagramas/.
#   4. El PPT se guarda en scripts/Manual_Codificacion_PPRA.pptx
# =============================================================================

library(officer)
library(ggplot2)

# ── Configuración ──────────────────────────────────────────────────────────────

TEMPLATE_PPTX <- system.file("plantillas/plantilla_16_9.pptx", package = "prosecnur")
if (!nzchar(TEMPLATE_PPTX)) {
  TEMPLATE_PPTX <- file.path(dirname(getwd()), "inst/plantillas/plantilla_16_9.pptx")
}

# Si se ejecuta desde scripts/
if (!file.exists(TEMPLATE_PPTX)) {
  TEMPLATE_PPTX <- system.file("plantillas/plantilla_16_9.pptx", package = "prosecnur")
}

SCREENSHOTS_DIR <- file.path(getwd(), "screenshots")
DIAGRAMAS_DIR   <- file.path(getwd(), "diagramas")
OUTPUT_PATH     <- file.path(getwd(), "Manual_Codificacion_PPRA.pptx")

if (!dir.exists(SCREENSHOTS_DIR)) dir.create(SCREENSHOTS_DIR, recursive = TRUE)
if (!dir.exists(DIAGRAMAS_DIR))   dir.create(DIAGRAMAS_DIR, recursive = TRUE)

PROYECTO <- "ACNUR / Pulso"
FECHA    <- format(Sys.Date(), "%d de %B de %Y")

# Modo de render de diagramas: "vector" (rvg) o "png".
# Se puede forzar con PPRA_DIAGRAM_RENDER_MODE=vector|png
DIAGRAM_VECTOR_AVAILABLE <- requireNamespace("rvg", quietly = TRUE)
render_mode_env <- tolower(trimws(Sys.getenv("PPRA_DIAGRAM_RENDER_MODE", "")))
if (render_mode_env %in% c("vector", "png")) {
  DIAGRAM_RENDER_MODE <- render_mode_env
} else {
  DIAGRAM_RENDER_MODE <- if (DIAGRAM_VECTOR_AVAILABLE) "vector" else "png"
}
if (identical(DIAGRAM_RENDER_MODE, "vector") && !DIAGRAM_VECTOR_AVAILABLE) {
  message("rvg no está disponible; se usará fallback PNG para diagramas.")
  DIAGRAM_RENDER_MODE <- "png"
}
message("Modo de render de diagramas: ", DIAGRAM_RENDER_MODE)

# ── Estilos de texto ──────────────────────────────────────────────────────────

fp_title    <- fp_text(font.size = 28, bold = TRUE, color = "#1F3864", font.family = "Calibri")
fp_subtitle <- fp_text(font.size = 18, italic = TRUE, color = "#4472C4", font.family = "Calibri")
fp_body     <- fp_text(font.size = 14, color = "#333333", font.family = "Calibri")
fp_bold     <- fp_text(font.size = 14, bold = TRUE, color = "#333333", font.family = "Calibri")
fp_small    <- fp_text(font.size = 12, color = "#666666", font.family = "Calibri")
fp_code     <- fp_text(font.size = 11, color = "#1B4332", font.family = "Courier New")
fp_code_b   <- fp_text(font.size = 11, color = "#1B4332", font.family = "Courier New", bold = TRUE)
fp_warning  <- fp_text(font.size = 14, bold = TRUE, color = "#C00000", font.family = "Calibri")
fp_tip      <- fp_text(font.size = 14, color = "#2E7D32", font.family = "Calibri")

# ── Helpers: slides ───────────────────────────────────────────────────────────

get_screenshot <- function(name, width = 8, height = 4.5) {
  path <- file.path(SCREENSHOTS_DIR, paste0(name, ".png"))
  if (file.exists(path)) return(path)
  p <- ggplot() +
    annotate("rect", xmin = 0, xmax = 10, ymin = 0, ymax = 6,
             fill = "#F0F0F0", color = "#CCCCCC", linewidth = 2) +
    annotate("text", x = 5, y = 3.5, label = paste0("[Screenshot]\n", name),
             size = 6, color = "#999999", fontface = "italic") +
    annotate("text", x = 5, y = 1.5,
             label = "Reemplace este archivo en scripts/screenshots/",
             size = 4, color = "#BBBBBB") +
    theme_void() +
    coord_fixed(ratio = height / (width * 0.6))
  ggsave(path, p, width = width, height = height, dpi = 150, bg = "white")
  path
}

add_title_slide <- function(doc, title, subtitle = NULL, date = NULL) {
  doc <- add_slide(doc, layout = "Title Slide", master = "Office Theme")
  doc <- ph_with(doc, value = title, location = ph_location_type(type = "ctrTitle"))
  if (!is.null(subtitle))
    doc <- ph_with(doc, value = subtitle, location = ph_location_type(type = "subTitle"))
  if (!is.null(date))
    doc <- ph_with(doc, value = date, location = ph_location_type(type = "dt"))
  doc
}

add_section_slide <- function(doc, title, subtitle = NULL) {
  doc <- add_slide(doc, layout = "Section Header", master = "Office Theme")
  doc <- ph_with(doc, value = title, location = ph_location_type(type = "title"))
  if (!is.null(subtitle))
    doc <- ph_with(doc, value = subtitle, location = ph_location_type(type = "body", type_idx = 1))
  doc
}

add_full_img_slide <- function(doc, title, img_path) {
  doc <- add_slide(doc, layout = "Graficos", master = "Office Theme")
  doc <- ph_with(doc, value = title, location = ph_location_type(type = "title"))
  doc <- ph_with(doc, value = external_img(img_path, width = 9, height = 5),
                 location = ph_location_type(type = "pic"))
  doc
}

get_placeholder_box <- function(doc, layout, master = "Office Theme", type = "pic", type_idx = 1) {
  props <- layout_properties(doc)
  props <- props[
    props$name == layout &
      props$master_name == master &
      props$type == type,
    ,
    drop = FALSE
  ]
  if (!nrow(props)) {
    stop("No se encontró placeholder tipo '", type, "' para layout '", layout, "'.", call. = FALSE)
  }
  ord <- order(ifelse(is.na(props$type_idx), Inf, props$type_idx), props$id)
  props <- props[ord, , drop = FALSE]
  idx <- max(1, min(type_idx, nrow(props)))
  row <- props[idx, , drop = FALSE]
  list(left = row$offx, top = row$offy, width = row$cx, height = row$cy)
}

center_fit_in_box <- function(content_width, content_height, box_width, box_height) {
  content_ratio <- content_width / content_height
  box_ratio <- box_width / box_height
  if (content_ratio >= box_ratio) {
    width <- box_width
    height <- box_width / content_ratio
    left <- 0
    top <- (box_height - height) / 2
  } else {
    height <- box_height
    width <- box_height * content_ratio
    left <- (box_width - width) / 2
    top <- 0
  }
  list(left = left, top = top, width = width, height = height)
}

add_full_diagram_slide <- function(doc, title, diagram_asset) {
  doc <- add_slide(doc, layout = "Graficos", master = "Office Theme")
  doc <- ph_with(doc, value = title, location = ph_location_type(type = "title"))

  box <- get_placeholder_box(doc, layout = "Graficos", master = "Office Theme", type = "pic", type_idx = 1)
  fit <- center_fit_in_box(
    content_width = diagram_asset$width,
    content_height = diagram_asset$height,
    box_width = box$width,
    box_height = box$height
  )
  loc <- ph_location(
    left = box$left + fit$left,
    top = box$top + fit$top,
    width = fit$width,
    height = fit$height
  )

  use_vector <- identical(DIAGRAM_RENDER_MODE, "vector") &&
    DIAGRAM_VECTOR_AVAILABLE &&
    !is.null(diagram_asset$plot)

  if (use_vector) {
    doc <- ph_with(
      doc,
      value = rvg::dml(ggobj = diagram_asset$plot, bg = "transparent"),
      location = loc
    )
  } else {
    doc <- ph_with(
      doc,
      value = external_img(diagram_asset$png_path, width = fit$width, height = fit$height),
      location = loc
    )
  }
  doc
}

add_text_slide <- function(doc, title, content_block) {
  doc <- add_slide(doc, layout = "Graficos", master = "Office Theme")
  doc <- ph_with(doc, value = title, location = ph_location_type(type = "title"))
  doc <- ph_with(doc, value = content_block, location = ph_location_type(type = "body", type_idx = 2))
  doc
}

add_img_text_r_slide <- function(doc, title, img_path, content_block) {
  doc <- add_slide(doc, layout = "right_grafico_texto", master = "Office Theme")
  doc <- ph_with(doc, value = title, location = ph_location_type(type = "title"))
  doc <- ph_with(doc, value = external_img(img_path, width = 5.5, height = 4),
                 location = ph_location_type(type = "pic", type_idx = 1))
  doc <- ph_with(doc, value = content_block,
                 location = ph_location_type(type = "body", type_idx = 2))
  doc
}

add_img_text_l_slide <- function(doc, title, img_path, content_block) {
  doc <- add_slide(doc, layout = "left_grafico_texto", master = "Office Theme")
  doc <- ph_with(doc, value = title, location = ph_location_type(type = "title"))
  doc <- ph_with(doc, value = content_block,
                 location = ph_location_type(type = "body", type_idx = 2))
  doc <- ph_with(doc, value = external_img(img_path, width = 5.5, height = 4),
                 location = ph_location_type(type = "pic", type_idx = 1))
  doc
}

add_two_img_slide <- function(doc, title, img_left, img_right) {
  doc <- add_slide(doc, layout = "Graficos_2columnas", master = "Office Theme")
  doc <- ph_with(doc, value = title, location = ph_location_type(type = "title"))
  doc <- ph_with(doc, value = external_img(img_left, width = 5, height = 4),
                 location = ph_location_type(type = "pic", type_idx = 2))
  doc <- ph_with(doc, value = external_img(img_right, width = 5, height = 4),
                 location = ph_location_type(type = "pic", type_idx = 1))
  doc
}

# ── Helpers: diagramas ────────────────────────────────────────────────────────

# Tema base para todos los diagramas
theme_diagrama <- function() {
  theme_void() +
    theme(
      plot.background = element_rect(fill = "transparent", color = NA),
      panel.background = element_rect(fill = "transparent", color = NA),
      plot.margin = margin(10, 10, 10, 10)
    )
}

# Dibuja una caja con texto (nodo de flowchart)
draw_box <- function(x, y, w, h, label, fill = "#FFFFFF", border = "#333333",
                     text_color = "#333333", text_size = 3.5, fontface = "plain",
                     radius = 0.02) {
  list(
    annotate("rect", xmin = x - w/2, xmax = x + w/2,
             ymin = y - h/2, ymax = y + h/2,
             fill = fill, color = border, linewidth = 0.8),
    annotate("text", x = x, y = y, label = label,
             size = text_size, color = text_color, fontface = fontface,
             lineheight = 0.9)
  )
}

# Dibuja una flecha vertical entre dos puntos
draw_arrow_v <- function(x, y_from, y_to, color = "#555555") {
  annotate("segment", x = x, xend = x, y = y_from, yend = y_to,
           arrow = arrow(length = unit(0.15, "cm"), type = "closed"),
           color = color, linewidth = 0.6)
}

# Dibuja una flecha horizontal
draw_arrow_h <- function(y, x_from, x_to, color = "#555555") {
  annotate("segment", x = x_from, xend = x_to, y = y, yend = y,
           arrow = arrow(length = unit(0.15, "cm"), type = "closed"),
           color = color, linewidth = 0.6)
}

# Genera y guarda un diagrama y devuelve metadatos para render híbrido
save_diagram <- function(p, name, width = 9, height = 5.5) {
  path <- file.path(DIAGRAMAS_DIR, paste0(name, ".png"))
  png_device <- if (requireNamespace("ragg", quietly = TRUE)) ragg::agg_png else "png"
  ggsave(path, p, width = width, height = height, dpi = 320, bg = "transparent", device = png_device)
  list(
    name = name,
    plot = p,
    png_path = path,
    width = width,
    height = height
  )
}

# =============================================================================
# GENERACIÓN DE DIAGRAMAS
# =============================================================================

message("Generando diagramas...")

# ── Diagrama 1: Flujo general de codificación ────────────────────────────────

diag_flujo_general <- ggplot() +
  # Cajas principales
  draw_box(2, 9, 3, 0.9, "1. FAMILIAS.xlsx\n(Definir variables)", fill = "#E8F0FE", border = "#4472C4", text_size = 3.2) +
  draw_box(6, 9, 3, 0.9, "XLSForm\n(Instrumento)", fill = "#E8F0FE", border = "#4472C4", text_size = 3.2) +
  draw_box(10, 9, 3, 0.9, "Datos crudos\n(.xlsx / .csv)", fill = "#E8F0FE", border = "#4472C4", text_size = 3.2) +

  # Flechas convergentes
  draw_arrow_v(2, 8.55, 7.55) +
  draw_arrow_v(6, 8.55, 7.55) +
  draw_arrow_v(10, 8.55, 7.55) +

  # Lineas horizontales de convergencia
  annotate("segment", x = 2, xend = 10, y = 7.5, yend = 7.5, color = "#555555", linewidth = 0.6) +
  draw_arrow_v(6, 7.5, 6.55) +

  # Paso 2: Generar plantilla
  draw_box(6, 6.1, 5, 0.9, "2. Generar plantilla de codificacion\nconstruir_plantilla + exportar_plantilla", fill = "#FFF2CC", border = "#BF8F00", text_size = 3.2, fontface = "bold") +
  draw_arrow_v(6, 5.65, 4.85) +

  # Paso 3: Codificador
  draw_box(6, 4.4, 5, 0.9, "3. Codificador llena la plantilla\n(Solo Excel, no requiere R)", fill = "#DFF5DF", border = "#2E7D32", text_size = 3.2, fontface = "bold") +

  # Icono persona
  annotate("text", x = 3, y = 4.4, label = "\U0001F464", size = 8) +

  draw_arrow_v(6, 3.95, 3.15) +

  # Paso 4 y 5: Aplicar
  draw_box(4.5, 2.7, 4, 0.9, "4. ppra_adaptar_data()\nAplicar codificacion a datos", fill = "#DCEBFF", border = "#1F4E79", text_size = 3.2) +
  draw_box(9, 2.7, 3.5, 0.9, "5. ppra_adaptar_instrumento()\nActualizar XLSForm", fill = "#E6D9F2", border = "#5B2C6F", text_size = 3.2) +

  annotate("segment", x = 6, xend = 9, y = 3.1, yend = 3.15, color = "#555555", linewidth = 0.6,
           arrow = arrow(length = unit(0.15, "cm"), type = "closed")) +
  draw_arrow_v(4.5, 2.25, 1.45) +
  draw_arrow_v(9, 2.25, 1.45) +

  # Resultado
  draw_box(4.5, 1, 3.5, 0.9, "Datos recodificados\n(.xlsx con *_recod)", fill = "#E2EFDA", border = "#548235", text_size = 3.2) +
  draw_box(9, 1, 3.5, 0.9, "Instrumento adaptado\n(nuevas categorias)", fill = "#E2EFDA", border = "#548235", text_size = 3.2) +

  # Etiquetas de rol
  annotate("label", x = 11.8, y = 6.1, label = "Tecnico R", fill = "#FFF2CC",
           color = "#BF8F00", size = 2.8, fontface = "italic", label.size = 0.3) +
  annotate("label", x = 11.8, y = 4.4, label = "Codificador", fill = "#DFF5DF",
           color = "#2E7D32", size = 2.8, fontface = "italic", label.size = 0.3) +
  annotate("label", x = 11.8, y = 2.7, label = "Tecnico R", fill = "#FFF2CC",
           color = "#BF8F00", size = 2.8, fontface = "italic", label.size = 0.3) +

  xlim(0.5, 12.5) + ylim(0.3, 9.8) +
  theme_diagrama()

path_flujo_general <- save_diagram(diag_flujo_general, "flujo_general")

# ── Diagrama 2: Mapa de zonas de una hoja ────────────────────────────────────

diag_zonas <- ggplot() +
  # Fondo de la "hoja Excel"
  annotate("rect", xmin = 0.5, xmax = 12.5, ymin = 0.5, ymax = 8, fill = "#FAFAFA", color = "#CCCCCC") +

  # Fila 1 (técnica)
  annotate("rect", xmin = 0.5, xmax = 12.5, ymin = 7.2, ymax = 8, fill = "#F0F0F0", color = "#999999") +
  annotate("text", x = 0.2, y = 7.6, label = "Fila 1", size = 2.5, color = "#999999", hjust = 1, fontface = "bold") +
  annotate("text", x = 6.5, y = 7.6, label = "Nombres tecnicos (NO EDITAR)", size = 3, color = "#666666", fontface = "italic") +

  # Fila 2 (etiquetas)
  annotate("rect", xmin = 0.5, xmax = 12.5, ymin = 6.4, ymax = 7.2, fill = "#F5F5F5", color = "#999999") +
  annotate("text", x = 0.2, y = 6.8, label = "Fila 2", size = 2.5, color = "#999999", hjust = 1, fontface = "bold") +
  annotate("text", x = 6.5, y = 6.8, label = "Etiquetas legibles (NO EDITAR)", size = 3, color = "#666666", fontface = "italic") +

  # Zona ID (azul oscuro)
  annotate("rect", xmin = 0.5, xmax = 2.5, ymin = 0.5, ymax = 6.4, fill = "#DDE3EA", color = "#1F3864", linewidth = 1.2) +
  annotate("text", x = 1.5, y = 5.5, label = "ID", size = 4.5, color = "#1F3864", fontface = "bold") +
  annotate("text", x = 1.5, y = 4.8, label = "_uuid\n_index\nCodigo pulso", size = 2.5, color = "#1F3864", lineheight = 0.9) +
  annotate("text", x = 1.5, y = 3.5, label = "NO EDITAR", size = 2.8, color = "#C00000", fontface = "bold") +

  # Zona Referencia (gris)
  annotate("rect", xmin = 2.5, xmax = 5.5, ymin = 0.5, ymax = 6.4, fill = "#F7F7F8", color = "#808080", linewidth = 1.2) +
  annotate("text", x = 4, y = 5.5, label = "Referencia", size = 4, color = "#555555", fontface = "bold") +
  annotate("text", x = 4, y = 4.8, label = "Valores originales\nLabels\nSeleccionadas", size = 2.5, color = "#555555", lineheight = 0.9) +
  annotate("text", x = 4, y = 3.5, label = "NO EDITAR", size = 2.8, color = "#C00000", fontface = "bold") +

  # Zona Editable (verde)
  annotate("rect", xmin = 5.5, xmax = 8.5, ymin = 0.5, ymax = 6.4, fill = "#EAF7E6", color = "#2E7D32", linewidth = 1.5) +
  annotate("text", x = 7, y = 5.5, label = "EDITABLE", size = 4.5, color = "#2E7D32", fontface = "bold") +
  annotate("text", x = 7, y = 4.8, label = "*_recod\n(columnas de\nrecodificacion)", size = 2.5, color = "#2E7D32", lineheight = 0.9) +
  annotate("text", x = 7, y = 3.5, label = "AQUI SE\nCODIFICA", size = 3.2, color = "#2E7D32", fontface = "bold", lineheight = 0.85) +
  # Borde destacado
  annotate("rect", xmin = 5.5, xmax = 8.5, ymin = 0.5, ymax = 6.4, fill = NA, color = "#2E7D32", linewidth = 2, linetype = "solid") +

  # Zona Control (naranja)
  annotate("rect", xmin = 8.5, xmax = 10, ymin = 0.5, ymax = 6.4, fill = "#FFF2E8", color = "#C65911", linewidth = 1.2) +
  annotate("text", x = 9.25, y = 5.5, label = "Control", size = 3.5, color = "#C65911", fontface = "bold") +
  annotate("text", x = 9.25, y = 4.6, label = "Notas y\nobservaciones", size = 2.3, color = "#C65911", lineheight = 0.9) +
  annotate("text", x = 9.25, y = 3.5, label = "Opcional", size = 2.5, color = "#C65911", fontface = "italic") +

  # Separador
  annotate("rect", xmin = 10, xmax = 10.3, ymin = 0.5, ymax = 6.4, fill = "#FFFFFF", color = "#DDDDDD") +

  # Zona Auxiliar (rojo/salmón)
  annotate("rect", xmin = 10.3, xmax = 12.5, ymin = 0.5, ymax = 6.4, fill = "#FCE5CD", color = "#C00000", linewidth = 1.2) +
  annotate("text", x = 11.4, y = 5.5, label = "Auxiliar", size = 3.5, color = "#C00000", fontface = "bold") +
  annotate("text", x = 11.4, y = 4.6, label = "nuevo_codigo\nnueva_etiqueta", size = 2.3, color = "#C00000", lineheight = 0.9) +
  annotate("text", x = 11.4, y = 3.5, label = "Nuevas\ncategorias", size = 2.5, color = "#C00000", fontface = "italic", lineheight = 0.85) +

  # Flecha "Datos comienzan aquí"
  annotate("segment", x = 0.2, xend = 0.5, y = 3, yend = 3,
           arrow = arrow(length = unit(0.15, "cm"), type = "closed"), color = "#333333") +
  annotate("text", x = -0.4, y = 3, label = "Datos\n(fila 3+)", size = 2.5, color = "#333333", fontface = "bold", lineheight = 0.85) +

  xlim(-1, 13) + ylim(0, 8.5) +
  theme_diagrama()

path_zonas <- save_diagram(diag_zonas, "mapa_zonas_hoja")

# ── Diagrama 3: Árbol de decisión "¿Qué tipo de variable?" ───────────────────

diag_decision <- ggplot() +
  # Pregunta central
  draw_box(6, 9, 5, 1, "Que tipo de variable\nestoy codificando?", fill = "#E8F0FE", border = "#1F3864", text_size = 3.8, fontface = "bold") +

  # Ramas
  annotate("segment", x = 3, xend = 1.5, y = 8.5, yend = 7.5, color = "#555555", linewidth = 0.6,
           arrow = arrow(length = unit(0.15, "cm"), type = "closed")) +
  annotate("segment", x = 4.5, xend = 4.5, y = 8.5, yend = 7.5, color = "#555555", linewidth = 0.6,
           arrow = arrow(length = unit(0.15, "cm"), type = "closed")) +
  annotate("segment", x = 7.5, xend = 7.5, y = 8.5, yend = 7.5, color = "#555555", linewidth = 0.6,
           arrow = arrow(length = unit(0.15, "cm"), type = "closed")) +
  annotate("segment", x = 9, xend = 10.5, y = 8.5, yend = 7.5, color = "#555555", linewidth = 0.6,
           arrow = arrow(length = unit(0.15, "cm"), type = "closed")) +

  # 4 tipos
  draw_box(1.5, 7, 2.4, 0.9, "select_multiple", fill = "#DCEBFF", border = "#1F4E79", text_size = 3.2, fontface = "bold") +
  draw_box(4.5, 7, 2.4, 0.9, "select_one", fill = "#DFF5DF", border = "#2E7D32", text_size = 3.2, fontface = "bold") +
  draw_box(7.5, 7, 2.2, 0.9, "text", fill = "#FFF2CC", border = "#BF8F00", text_size = 3.2, fontface = "bold") +
  draw_box(10.5, 7, 2.2, 0.9, "integer", fill = "#E6D9F2", border = "#5B2C6F", text_size = 3.2, fontface = "bold") +

  # Flechas a descripción
  draw_arrow_v(1.5, 6.55, 5.75) +
  draw_arrow_v(4.5, 6.55, 5.75) +
  draw_arrow_v(7.5, 6.55, 5.75) +
  draw_arrow_v(10.5, 6.55, 5.75) +

  # Descripción de cada tipo
  draw_box(1.5, 5.2, 2.6, 1, "Varias respuestas\nCada opcion = columna\n1 / 0 / vacio", fill = "#F0F7FF", border = "#1F4E79", text_size = 2.6) +
  draw_box(4.5, 5.2, 2.6, 1, "Una sola respuesta\nCodigo en *_recod\nModo padre o hijo", fill = "#F0FAF0", border = "#2E7D32", text_size = 2.6) +
  draw_box(7.5, 5.2, 2.4, 1, "Texto libre\nCodigo en *_recod\nCategorizar respuestas", fill = "#FFFBEF", border = "#BF8F00", text_size = 2.6) +
  draw_box(10.5, 5.2, 2.4, 1, "Numero abierto\nCodigo en *_recod\nAgrupar en rangos", fill = "#F5EFFA", border = "#5B2C6F", text_size = 2.6) +

  # Flechas a "nuevas categorías"
  draw_arrow_v(1.5, 4.7, 4.05) +
  draw_arrow_v(4.5, 4.7, 4.05) +
  draw_arrow_v(7.5, 4.7, 4.05) +
  draw_arrow_v(10.5, 4.7, 4.05) +

  # Cómo declarar nuevas categorías
  draw_box(1.5, 3.5, 2.6, 1, "Nuevas opciones:\nColumna nueva con\n<parent>/<cod>_recod", fill = "#DCEBFF", border = "#1F4E79", text_size = 2.4) +
  draw_box(4.5, 3.5, 2.6, 1, "Nuevas categorias:\nBloque auxiliar\nnuevo_codigo +\nnueva_etiqueta", fill = "#DFF5DF", border = "#2E7D32", text_size = 2.4) +
  draw_box(7.5, 3.5, 2.4, 1, "Nuevas categorias:\nBloque auxiliar\nnuevo_codigo +\nnueva_etiqueta", fill = "#FFF2CC", border = "#BF8F00", text_size = 2.4) +
  draw_box(10.5, 3.5, 2.4, 1, "Nuevas categorias:\nBloque auxiliar\nnuevo_codigo +\nnueva_etiqueta", fill = "#E6D9F2", border = "#5B2C6F", text_size = 2.4) +

  # Nota inferior
  annotate("label", x = 6, y = 2.3,
           label = "Regla de oro: Deje vacio si no desea cambiar el valor original",
           fill = "#FFF8E1", color = "#BF8F00", size = 3.5, fontface = "bold.italic", label.size = 0.5) +

  xlim(0, 12) + ylim(1.8, 10) +
  theme_diagrama()

path_decision <- save_diagram(diag_decision, "arbol_decision_tipo")

# ── Diagrama 4: Flujo select_multiple paso a paso ────────────────────────────

diag_sm_flujo <- ggplot() +
  # Paso 1
  draw_box(2, 9, 3.2, 0.9, "Paso 1\nIdentificar opciones\noriginales del catalogo",
           fill = "#E8F0FE", border = "#4472C4", text_size = 3) +
  draw_arrow_h(y = 9, x_from = 3.6, x_to = 4.7) +

  # Paso 2
  draw_box(6.5, 9, 3.2, 0.9, "Paso 2\nRevisar cada fila\n(cada encuestado)",
           fill = "#E8F0FE", border = "#4472C4", text_size = 3) +
  draw_arrow_h(y = 9, x_from = 8.1, x_to = 9.2) +

  # Paso 3
  draw_box(11, 9, 3.2, 0.9, "Paso 3\nDecidir cambios\npor opcion",
           fill = "#DCEBFF", border = "#1F4E79", text_size = 3, fontface = "bold") +

  draw_arrow_v(11, 8.55, 7.55) +

  # Decisión
  annotate("polygon",
           x = c(11, 13, 11, 9),
           y = c(7.5, 6.5, 5.5, 6.5),
           fill = "#FFF2CC", color = "#BF8F00", linewidth = 0.8) +
  annotate("text", x = 11, y = 6.5, label = "Cambiar\nesta opcion?",
           size = 2.8, color = "#BF8F00", fontface = "bold", lineheight = 0.85) +

  # Rama SI - marcar
  annotate("segment", x = 13, xend = 13, y = 6.5, yend = 5.1,
           arrow = arrow(length = unit(0.12, "cm"), type = "closed"), color = "#2E7D32", linewidth = 0.6) +
  annotate("text", x = 13.3, y = 5.8, label = "Si", size = 2.8, color = "#2E7D32", fontface = "bold") +

  draw_box(13, 4.5, 2.2, 1, "Activar: escriba 1\nDesactivar: escriba 0",
           fill = "#DFF5DF", border = "#2E7D32", text_size = 2.8, fontface = "bold") +

  # Rama NO
  annotate("segment", x = 9, xend = 9, y = 6.5, yend = 5.1,
           arrow = arrow(length = unit(0.12, "cm"), type = "closed"), color = "#808080", linewidth = 0.6) +
  annotate("text", x = 8.6, y = 5.8, label = "No", size = 2.8, color = "#808080", fontface = "bold") +

  draw_box(9, 4.5, 2.2, 1, "Dejar vacio\n(sin cambio)",
           fill = "#F5F5F5", border = "#999999", text_size = 2.8) +

  # Paso 4: Nueva opción
  draw_box(4, 6.5, 3.5, 0.9, "Paso 4 (si aplica)\nNecesita una opcion\nque NO esta en catalogo?",
           fill = "#FFF2CC", border = "#BF8F00", text_size = 2.8) +

  draw_arrow_v(4, 6.05, 5.25) +

  draw_box(4, 4.7, 3.8, 1, "Insertar columna nueva:\nFila 1: <parent>/<cod>_recod\nFila 2: Etiqueta visible\nDatos: 1 / 0 / vacio",
           fill = "#DCEBFF", border = "#1F4E79", text_size = 2.6) +

  # Nota ejemplo
  annotate("label", x = 4, y = 3.5,
           label = "Ej: Fila 1 = p8/99_recod  |  Fila 2 = Otro servicio",
           fill = "#F0F7FF", color = "#1F4E79", size = 2.5, fontface = "italic", label.size = 0.3) +

  # Nota importante
  annotate("label", x = 8, y = 2.8,
           label = "Importante: La columna ejemplo_recod es solo referencia. El adaptador la ignora.",
           fill = "#FFF8E1", color = "#BF8F00", size = 2.8, fontface = "bold.italic", label.size = 0.4) +

  xlim(0.5, 14.5) + ylim(2.3, 9.8) +
  theme_diagrama()

path_sm_flujo <- save_diagram(diag_sm_flujo, "flujo_sm_pasos", width = 10, height = 5.5)

# ── Diagrama 5: Flujo select_one (padre vs hijo) ─────────────────────────────

diag_so_flujo <- ggplot() +
  # Inicio
  draw_box(6, 9.2, 4, 0.9, "select_one\nIdentificar el modo de la hoja",
           fill = "#DFF5DF", border = "#2E7D32", text_size = 3.3, fontface = "bold") +

  # Bifurcación
  annotate("segment", x = 4, xend = 3, y = 8.75, yend = 8.05, color = "#555555", linewidth = 0.6,
           arrow = arrow(length = unit(0.15, "cm"), type = "closed")) +
  annotate("segment", x = 8, xend = 9, y = 8.75, yend = 8.05, color = "#555555", linewidth = 0.6,
           arrow = arrow(length = unit(0.15, "cm"), type = "closed")) +

  # Modo padre
  draw_box(3, 7.5, 3.5, 0.9, "Modo PADRE\nRecodifica la variable original",
           fill = "#D5E8D4", border = "#2E7D32", text_size = 3, fontface = "bold") +
  draw_arrow_v(3, 7.05, 6.25) +

  draw_box(3, 5.8, 3.8, 0.8, "Columna editable:\nRecodificacion (codigo)",
           fill = "#F0FAF0", border = "#2E7D32", text_size = 2.8) +
  draw_arrow_v(3, 5.4, 4.65) +

  draw_box(3, 4.2, 3.8, 0.8, "Escriba el codigo final\no deje vacio (sin cambio)",
           fill = "#F0FAF0", border = "#2E7D32", text_size = 2.8) +
  draw_arrow_v(3, 3.8, 3.05) +

  draw_box(3, 2.6, 3.8, 0.8, "Ejemplo: respuesta 'Otro'\n-> recodificar a codigo 5",
           fill = "#E8F0FE", border = "#4472C4", text_size = 2.6, fontface = "italic") +

  # Modo hijo
  draw_box(9, 7.5, 3.5, 0.9, "Modo HIJO\nRecodifica el texto abierto",
           fill = "#DAE8FC", border = "#1F4E79", text_size = 3, fontface = "bold") +
  draw_arrow_v(9, 7.05, 6.25) +

  draw_box(9, 5.8, 3.8, 0.8, "Columna editable:\n<parent>_<alias>_recod",
           fill = "#F0F7FF", border = "#1F4E79", text_size = 2.8) +
  draw_arrow_v(9, 5.4, 4.65) +

  draw_box(9, 4.2, 3.8, 0.8, "Categorizar respuestas de\ntexto libre en codigos",
           fill = "#F0F7FF", border = "#1F4E79", text_size = 2.8) +
  draw_arrow_v(9, 3.8, 3.05) +

  draw_box(9, 2.6, 3.8, 0.8, "Ejemplo: 'hospital central'\n-> recodificar a codigo 3",
           fill = "#E8F0FE", border = "#4472C4", text_size = 2.6, fontface = "italic") +

  # Bloque auxiliar compartido
  annotate("segment", x = 3, xend = 6, y = 2.2, yend = 1.6, color = "#C00000", linewidth = 0.6,
           arrow = arrow(length = unit(0.12, "cm"), type = "closed")) +
  annotate("segment", x = 9, xend = 6, y = 2.2, yend = 1.6, color = "#C00000", linewidth = 0.6,
           arrow = arrow(length = unit(0.12, "cm"), type = "closed")) +

  draw_box(6, 1.1, 5, 0.9, "Ambos modos: declarar nuevas categorias\nen bloque auxiliar (nuevo_codigo + nueva_etiqueta)",
           fill = "#FCE5CD", border = "#C00000", text_size = 3, fontface = "bold") +

  xlim(0.5, 11.5) + ylim(0.4, 9.8) +
  theme_diagrama()

path_so_flujo <- save_diagram(diag_so_flujo, "flujo_so_padre_hijo")

# ── Diagrama 6: Flujo text / integer ─────────────────────────────────────────

diag_text_int <- ggplot() +
  # Títulos
  draw_box(3.5, 9, 3, 0.8, "text", fill = "#FFF2CC", border = "#BF8F00", text_size = 4, fontface = "bold") +
  draw_box(8.5, 9, 3, 0.8, "integer", fill = "#E6D9F2", border = "#5B2C6F", text_size = 4, fontface = "bold") +

  # Paso 1
  draw_arrow_v(3.5, 8.6, 7.85) +
  draw_arrow_v(8.5, 8.6, 7.85) +

  draw_box(3.5, 7.4, 3.2, 0.8, "1. Leer respuesta original\n(texto libre del encuestado)",
           fill = "#FFFBEF", border = "#BF8F00", text_size = 2.8) +
  draw_box(8.5, 7.4, 3.2, 0.8, "1. Leer valor original\n(numero abierto)",
           fill = "#F5EFFA", border = "#5B2C6F", text_size = 2.8) +

  # Paso 2
  draw_arrow_v(3.5, 7, 6.25) +
  draw_arrow_v(8.5, 7, 6.25) +

  draw_box(3.5, 5.8, 3.2, 0.8, "2. Asignar un codigo\nen <variable>_recod",
           fill = "#FFFBEF", border = "#BF8F00", text_size = 2.8) +
  draw_box(8.5, 5.8, 3.2, 0.8, "2. Asignar un codigo\nen <variable>_recod",
           fill = "#F5EFFA", border = "#5B2C6F", text_size = 2.8) +

  # Paso 3
  draw_arrow_v(3.5, 5.4, 4.65) +
  draw_arrow_v(8.5, 5.4, 4.65) +

  draw_box(3.5, 4.2, 3.2, 0.8, "3. Codigo nuevo?\nDeclare en bloque auxiliar",
           fill = "#FCE5CD", border = "#C00000", text_size = 2.8) +
  draw_box(8.5, 4.2, 3.2, 0.8, "3. Codigo nuevo?\nDeclare en bloque auxiliar",
           fill = "#FCE5CD", border = "#C00000", text_size = 2.8) +

  # Ejemplos
  draw_arrow_v(3.5, 3.8, 3.05) +
  draw_arrow_v(8.5, 3.8, 3.05) +

  draw_box(3.5, 2.5, 3.5, 1, "Ejemplo:\n'Falta de agua potable'\n-> codigo: 3\n-> etiqueta: Agua y saneamiento",
           fill = "#E8F0FE", border = "#4472C4", text_size = 2.4, fontface = "italic") +
  draw_box(8.5, 2.5, 3.5, 1, "Ejemplo:\nEdad = 47\n-> codigo: 3\n-> etiqueta: 40-49 anios",
           fill = "#E8F0FE", border = "#4472C4", text_size = 2.4, fontface = "italic") +

  # Nota especial integer
  annotate("label", x = 8.5, y = 1.3,
           label = "Si dos variables integer usan el mismo diccionario,\nel adaptador comparte la lista automaticamente.",
           fill = "#F5EFFA", color = "#5B2C6F", size = 2.5, fontface = "italic", label.size = 0.3) +

  # Marco "proceso identico"
  annotate("segment", x = 6, xend = 6, y = 8, yend = 3.3,
           color = "#CCCCCC", linewidth = 0.5, linetype = "dashed") +
  annotate("text", x = 6, y = 8.5, label = "Proceso\nidentico", size = 2.5,
           color = "#AAAAAA", fontface = "italic", lineheight = 0.85) +

  xlim(1, 11) + ylim(0.8, 9.5) +
  theme_diagrama()

path_text_int <- save_diagram(diag_text_int, "flujo_text_integer")

# ── Diagrama 7: Bloque auxiliar detallado ─────────────────────────────────────

diag_auxiliar <- ggplot() +
  # Marco de la hoja Excel simulada
  annotate("rect", xmin = 1, xmax = 11, ymin = 1, ymax = 9, fill = "#FAFAFA", color = "#CCCCCC") +

  # Separador entre zona principal y auxiliar
  annotate("rect", xmin = 5.8, xmax = 6.2, ymin = 1, ymax = 9, fill = "#FFFFFF", color = "#DDDDDD") +
  annotate("text", x = 6, y = 9.3, label = "separador", size = 2.3, color = "#AAAAAA", fontface = "italic") +

  # Zona principal (izquierda)
  annotate("text", x = 3.4, y = 9.3, label = "ZONA PRINCIPAL", size = 3, color = "#2E7D32", fontface = "bold") +

  # Headers zona principal
  annotate("rect", xmin = 1, xmax = 2.6, ymin = 8, ymax = 8.7, fill = "#DDE3EA", color = "#999999") +
  annotate("text", x = 1.8, y = 8.35, label = "_uuid", size = 2.5, color = "#1F3864") +

  annotate("rect", xmin = 2.6, xmax = 4.2, ymin = 8, ymax = 8.7, fill = "#F7F7F8", color = "#999999") +
  annotate("text", x = 3.4, y = 8.35, label = "Seleccion\n(label)", size = 2.2, color = "#555555", lineheight = 0.85) +

  annotate("rect", xmin = 4.2, xmax = 5.8, ymin = 8, ymax = 8.7, fill = "#C6EFCE", color = "#2E7D32") +
  annotate("text", x = 5, y = 8.35, label = "*_recod", size = 2.5, color = "#2E7D32", fontface = "bold") +

  # Datos de ejemplo
  annotate("rect", xmin = 1, xmax = 2.6, ymin = 6.5, ymax = 8, fill = "#EFF3F7", color = "#DDDDDD") +
  annotate("text", x = 1.8, y = 7.65, label = "abc-001", size = 2, color = "#333333") +
  annotate("text", x = 1.8, y = 7.15, label = "abc-002", size = 2, color = "#333333") +
  annotate("text", x = 1.8, y = 6.65, label = "abc-003", size = 2, color = "#333333") +

  annotate("rect", xmin = 2.6, xmax = 4.2, ymin = 6.5, ymax = 8, fill = "#F7F7F8", color = "#DDDDDD") +
  annotate("text", x = 3.4, y = 7.65, label = "Si", size = 2, color = "#333333") +
  annotate("text", x = 3.4, y = 7.15, label = "Otro", size = 2, color = "#333333") +
  annotate("text", x = 3.4, y = 6.65, label = "No", size = 2, color = "#333333") +

  annotate("rect", xmin = 4.2, xmax = 5.8, ymin = 6.5, ymax = 8, fill = "#EAF7E6", color = "#DDDDDD") +
  annotate("text", x = 5, y = 7.65, label = "", size = 2, color = "#333333") +
  annotate("text", x = 5, y = 7.15, label = "99", size = 2.5, color = "#C00000", fontface = "bold") +
  annotate("text", x = 5, y = 6.65, label = "", size = 2, color = "#333333") +

  # Zona auxiliar (derecha)
  annotate("text", x = 8.6, y = 9.3, label = "BLOQUE AUXILIAR", size = 3, color = "#C00000", fontface = "bold") +

  # Headers auxiliar
  annotate("rect", xmin = 6.2, xmax = 8, ymin = 8, ymax = 8.7, fill = "#F4CCCC", color = "#C00000") +
  annotate("text", x = 7.1, y = 8.35, label = "nuevo_codigo", size = 2.3, color = "#C00000", fontface = "bold") +

  annotate("rect", xmin = 8, xmax = 11, ymin = 8, ymax = 8.7, fill = "#F4CCCC", color = "#C00000") +
  annotate("text", x = 9.5, y = 8.35, label = "nueva_etiqueta", size = 2.3, color = "#C00000", fontface = "bold") +

  # Datos auxiliar
  annotate("rect", xmin = 6.2, xmax = 8, ymin = 7.3, ymax = 8, fill = "#FCE5CD", color = "#DDDDDD") +
  annotate("text", x = 7.1, y = 7.65, label = "99", size = 2.5, color = "#C00000", fontface = "bold") +

  annotate("rect", xmin = 8, xmax = 11, ymin = 7.3, ymax = 8, fill = "#FCE5CD", color = "#DDDDDD") +
  annotate("text", x = 9.5, y = 7.65, label = "No aplica / No sabe", size = 2.3, color = "#C00000", fontface = "bold") +

  # Fila vacía (sólo se declara una vez)
  annotate("rect", xmin = 6.2, xmax = 11, ymin = 6.5, ymax = 7.3, fill = "#FFF8F0", color = "#DDDDDD") +
  annotate("text", x = 8.6, y = 6.9, label = "(solo 1 fila por codigo nuevo)", size = 2.2, color = "#999999", fontface = "italic") +

  # Flecha de conexión
  annotate("curve", x = 5, xend = 7.1, y = 7.15, yend = 7.65,
           curvature = -0.3, color = "#C00000", linewidth = 0.8,
           arrow = arrow(length = unit(0.15, "cm"), type = "closed")) +
  annotate("label", x = 6.1, y = 6, label = "El codigo 99 usado\nen *_recod se declara\naqui con su etiqueta",
           fill = "#FFF8E1", color = "#C00000", size = 2.5, fontface = "bold", label.size = 0.3, lineheight = 0.9) +

  # Reglas
  annotate("rect", xmin = 1, xmax = 11, ymin = 1.3, ymax = 4.5, fill = "#FFFFF0", color = "#BF8F00", linewidth = 0.8) +
  annotate("text", x = 6, y = 4.1, label = "Reglas del bloque auxiliar", size = 3.2, color = "#BF8F00", fontface = "bold") +

  annotate("text", x = 1.5, y = 3.5, hjust = 0, size = 2.6, color = "#333333", lineheight = 1.1,
           label = paste0(
             "1. Declare cada codigo nuevo UNA SOLA VEZ (una fila)\n",
             "2. Cada codigo debe tener exactamente UNA etiqueta\n",
             "3. Si el codigo ya existe en el catalogo original, no lo declare\n",
             "4. Si un codigo tiene 2 etiquetas distintas, el adaptador dara error\n",
             "5. Aplica a: select_one, text e integer (NO a select_multiple)"
           )) +

  xlim(0.5, 11.5) + ylim(0.8, 9.8) +
  theme_diagrama()

path_auxiliar <- save_diagram(diag_auxiliar, "diagrama_bloque_auxiliar", width = 10, height = 6)

# ── Diagrama 8: Flujo técnico R ──────────────────────────────────────────────

diag_flujo_r <- ggplot() +
  # Fase 1: Preparar
  annotate("label", x = 1.5, y = 9.3, label = "FASE 1: Preparar", fill = "#E8F0FE",
           color = "#1F3864", size = 3, fontface = "bold", label.size = 0.5) +

  draw_box(1.5, 8.3, 2.8, 0.8, "leer_instrumento_xlsform()\nCargar XLSForm",
           fill = "#E8F0FE", border = "#4472C4", text_size = 2.6) +
  draw_box(4.5, 8.3, 2.8, 0.8, "leer_datos()\nCargar datos crudos",
           fill = "#E8F0FE", border = "#4472C4", text_size = 2.6) +
  draw_box(7.5, 8.3, 2.8, 0.8, "leer_familias_clasificar()\nCargar FAMILIAS.xlsx",
           fill = "#E8F0FE", border = "#4472C4", text_size = 2.6) +

  # Convergencia
  annotate("segment", x = 1.5, xend = 1.5, y = 7.9, yend = 7.4, color = "#555555", linewidth = 0.5) +
  annotate("segment", x = 4.5, xend = 4.5, y = 7.9, yend = 7.4, color = "#555555", linewidth = 0.5) +
  annotate("segment", x = 7.5, xend = 7.5, y = 7.9, yend = 7.4, color = "#555555", linewidth = 0.5) +
  annotate("segment", x = 1.5, xend = 7.5, y = 7.4, yend = 7.4, color = "#555555", linewidth = 0.5) +
  draw_arrow_v(4.5, 7.4, 6.85) +

  # Fase 2: Generar
  annotate("label", x = 4.5, y = 6.9, label = "FASE 2: Generar plantilla", fill = "#FFF2CC",
           color = "#BF8F00", size = 3, fontface = "bold", label.size = 0.5) +

  draw_box(4.5, 6, 4.5, 0.8, "construir_plantilla_desde_familias(inst, dat, split)\nCrea estructura multi-hoja",
           fill = "#FFF2CC", border = "#BF8F00", text_size = 2.6) +
  draw_arrow_v(4.5, 5.6, 5.15) +

  draw_box(4.5, 4.7, 4.5, 0.8, "exportar_plantilla_codificacion_xlsx(plantilla, path)\nExporta con formato y colores",
           fill = "#FFF2CC", border = "#BF8F00", text_size = 2.6) +

  # Enviar al codificador
  draw_arrow_h(y = 4.7, x_from = 6.75, x_to = 8.2) +
  draw_box(10, 4.7, 3.2, 0.8, "Enviar .xlsx\nal codificador",
           fill = "#DFF5DF", border = "#2E7D32", text_size = 2.8, fontface = "bold") +

  # Recibir de vuelta
  draw_arrow_v(10, 4.3, 3.65) +
  draw_box(10, 3.2, 3.2, 0.8, "Recibir plantilla\ncompletada",
           fill = "#DFF5DF", border = "#2E7D32", text_size = 2.8, fontface = "bold") +

  annotate("segment", x = 10, xend = 7, y = 2.8, yend = 2.45, color = "#555555", linewidth = 0.6,
           arrow = arrow(length = unit(0.15, "cm"), type = "closed")) +

  # Fase 3: Aplicar
  annotate("label", x = 4.5, y = 3.2, label = "FASE 3: Aplicar", fill = "#DCEBFF",
           color = "#1F4E79", size = 3, fontface = "bold", label.size = 0.5) +

  draw_box(4.5, 2, 4.5, 0.8, "ppra_adaptar_data(path_inst, path_dat, path_plantilla, ...)\nInserta columnas *_recod en datos",
           fill = "#DCEBFF", border = "#1F4E79", text_size = 2.4) +
  draw_arrow_v(4.5, 1.6, 1.15) +

  draw_box(4.5, 0.7, 4.5, 0.8, "ppra_adaptar_instrumento(path_inst, path_data_adaptada, ...)\nActualiza survey + choices del XLSForm",
           fill = "#E6D9F2", border = "#5B2C6F", text_size = 2.4) +

  xlim(-0.2, 12) + ylim(0.1, 9.8) +
  theme_diagrama()

path_flujo_r <- save_diagram(diag_flujo_r, "flujo_tecnico_r_diagrama", width = 10, height = 5.5)

# ── Diagrama 9: Anatomía de columna SM ────────────────────────────────────────

diag_sm_cols <- ggplot() +
  # Título del esquema
  annotate("text", x = 6, y = 9.5, label = "Anatomia de una hoja select_multiple",
           size = 4, color = "#1F4E79", fontface = "bold") +

  # Header row (Fila 1)
  annotate("rect", xmin = 0.5, xmax = 1.5, ymin = 8, ymax = 8.7, fill = "#DDE3EA", color = "#999999") +
  annotate("text", x = 1, y = 8.35, label = "_uuid", size = 2, color = "#1F3864") +

  annotate("rect", xmin = 1.5, xmax = 2.8, ymin = 8, ymax = 8.7, fill = "#F7F7F8", color = "#999999") +
  annotate("text", x = 2.15, y = 8.35, label = "p8/1", size = 2, color = "#555555") +

  annotate("rect", xmin = 2.8, xmax = 4.1, ymin = 8, ymax = 8.7, fill = "#C6EFCE", color = "#2E7D32") +
  annotate("text", x = 3.45, y = 8.35, label = "p8/1_recod", size = 2, color = "#2E7D32", fontface = "bold") +

  annotate("rect", xmin = 4.1, xmax = 5.2, ymin = 8, ymax = 8.7, fill = "#F7F7F8", color = "#999999") +
  annotate("text", x = 4.65, y = 8.35, label = "p8/2", size = 2, color = "#555555") +

  annotate("rect", xmin = 5.2, xmax = 6.5, ymin = 8, ymax = 8.7, fill = "#C6EFCE", color = "#2E7D32") +
  annotate("text", x = 5.85, y = 8.35, label = "p8/2_recod", size = 2, color = "#2E7D32", fontface = "bold") +

  annotate("rect", xmin = 6.5, xmax = 7.6, ymin = 8, ymax = 8.7, fill = "#F7F7F8", color = "#999999") +
  annotate("text", x = 7.05, y = 8.35, label = "p8/3", size = 2, color = "#555555") +

  annotate("rect", xmin = 7.6, xmax = 8.9, ymin = 8, ymax = 8.7, fill = "#C6EFCE", color = "#2E7D32") +
  annotate("text", x = 8.25, y = 8.35, label = "p8/3_recod", size = 2, color = "#2E7D32", fontface = "bold") +

  annotate("rect", xmin = 8.9, xmax = 10, ymin = 8, ymax = 8.7, fill = "#FCE4D6", color = "#C65911") +
  annotate("text", x = 9.45, y = 8.35, label = "Control", size = 2, color = "#C65911") +

  annotate("rect", xmin = 10.2, xmax = 11.5, ymin = 8, ymax = 8.7, fill = "#D9EAD3", color = "#548235") +
  annotate("text", x = 10.85, y = 8.35, label = "ejemplo_recod", size = 1.8, color = "#548235") +

  # Fila 2 (etiquetas)
  annotate("rect", xmin = 0.5, xmax = 1.5, ymin = 7.3, ymax = 8, fill = "#EFF3F7", color = "#DDDDDD") +
  annotate("text", x = 1, y = 7.65, label = "ID", size = 1.8, color = "#555555") +

  annotate("rect", xmin = 1.5, xmax = 2.8, ymin = 7.3, ymax = 8, fill = "#F7F7F8", color = "#DDDDDD") +
  annotate("text", x = 2.15, y = 7.65, label = "Salud", size = 1.8, color = "#555555") +

  annotate("rect", xmin = 2.8, xmax = 4.1, ymin = 7.3, ymax = 8, fill = "#EAF7E6", color = "#DDDDDD") +
  annotate("text", x = 3.45, y = 7.65, label = "Salud\n(recod)", size = 1.6, color = "#2E7D32", lineheight = 0.85) +

  annotate("rect", xmin = 4.1, xmax = 5.2, ymin = 7.3, ymax = 8, fill = "#F7F7F8", color = "#DDDDDD") +
  annotate("text", x = 4.65, y = 7.65, label = "Educacion", size = 1.8, color = "#555555") +

  annotate("rect", xmin = 5.2, xmax = 6.5, ymin = 7.3, ymax = 8, fill = "#EAF7E6", color = "#DDDDDD") +
  annotate("text", x = 5.85, y = 7.65, label = "Educacion\n(recod)", size = 1.6, color = "#2E7D32", lineheight = 0.85) +

  annotate("rect", xmin = 6.5, xmax = 7.6, ymin = 7.3, ymax = 8, fill = "#F7F7F8", color = "#DDDDDD") +
  annotate("text", x = 7.05, y = 7.65, label = "Agua", size = 1.8, color = "#555555") +

  annotate("rect", xmin = 7.6, xmax = 8.9, ymin = 7.3, ymax = 8, fill = "#EAF7E6", color = "#DDDDDD") +
  annotate("text", x = 8.25, y = 7.65, label = "Agua\n(recod)", size = 1.6, color = "#2E7D32", lineheight = 0.85) +

  annotate("rect", xmin = 8.9, xmax = 10, ymin = 7.3, ymax = 8, fill = "#FFF2E8", color = "#DDDDDD") +
  annotate("text", x = 9.45, y = 7.65, label = "Notas", size = 1.8, color = "#C65911") +

  annotate("rect", xmin = 10.2, xmax = 11.5, ymin = 7.3, ymax = 8, fill = "#F3F9F1", color = "#DDDDDD") +
  annotate("text", x = 10.85, y = 7.65, label = "(referencia)", size = 1.6, color = "#548235", fontface = "italic") +

  # Datos ejemplo (fila 3)
  annotate("rect", xmin = 0.5, xmax = 1.5, ymin = 6.5, ymax = 7.3, fill = "#EFF3F7", color = "#DDDDDD") +
  annotate("text", x = 1, y = 6.9, label = "abc-01", size = 1.8, color = "#333333") +

  annotate("rect", xmin = 1.5, xmax = 2.8, ymin = 6.5, ymax = 7.3, fill = "#FFFFFF", color = "#DDDDDD") +
  annotate("text", x = 2.15, y = 6.9, label = "1", size = 2, color = "#333333") +

  annotate("rect", xmin = 2.8, xmax = 4.1, ymin = 6.5, ymax = 7.3, fill = "#EAF7E6", color = "#DDDDDD") +
  annotate("text", x = 3.45, y = 6.9, label = "0", size = 2.2, color = "#C00000", fontface = "bold") +

  annotate("rect", xmin = 4.1, xmax = 5.2, ymin = 6.5, ymax = 7.3, fill = "#FFFFFF", color = "#DDDDDD") +
  annotate("text", x = 4.65, y = 6.9, label = "0", size = 2, color = "#333333") +

  annotate("rect", xmin = 5.2, xmax = 6.5, ymin = 6.5, ymax = 7.3, fill = "#EAF7E6", color = "#DDDDDD") +
  annotate("text", x = 5.85, y = 6.9, label = "", size = 2, color = "#333333") +

  annotate("rect", xmin = 6.5, xmax = 7.6, ymin = 6.5, ymax = 7.3, fill = "#FFFFFF", color = "#DDDDDD") +
  annotate("text", x = 7.05, y = 6.9, label = "1", size = 2, color = "#333333") +

  annotate("rect", xmin = 7.6, xmax = 8.9, ymin = 6.5, ymax = 7.3, fill = "#EAF7E6", color = "#DDDDDD") +
  annotate("text", x = 8.25, y = 6.9, label = "1", size = 2.2, color = "#2E7D32", fontface = "bold") +

  # Anotaciones
  annotate("curve", x = 3.45, xend = 3.45, y = 6.3, yend = 5.7,
           curvature = 0, color = "#C00000", linewidth = 0.7,
           arrow = arrow(length = unit(0.12, "cm"), type = "closed")) +
  annotate("label", x = 3.45, y = 5.3,
           label = "0 = desmarcar\n(tenia 1, se cambia a 0)",
           fill = "#FFF0F0", color = "#C00000", size = 2.5, fontface = "bold", label.size = 0.3, lineheight = 0.9) +

  annotate("curve", x = 5.85, xend = 5.85, y = 6.3, yend = 5.7,
           curvature = 0, color = "#808080", linewidth = 0.7,
           arrow = arrow(length = unit(0.12, "cm"), type = "closed")) +
  annotate("label", x = 5.85, y = 5.3,
           label = "vacio = sin cambio\n(mantiene valor original)",
           fill = "#F5F5F5", color = "#555555", size = 2.5, fontface = "bold", label.size = 0.3, lineheight = 0.9) +

  annotate("curve", x = 8.25, xend = 8.25, y = 6.3, yend = 5.7,
           curvature = 0, color = "#2E7D32", linewidth = 0.7,
           arrow = arrow(length = unit(0.12, "cm"), type = "closed")) +
  annotate("label", x = 8.25, y = 5.3,
           label = "1 = marcar\n(confirmar o activar)",
           fill = "#E8F5E8", color = "#2E7D32", size = 2.5, fontface = "bold", label.size = 0.3, lineheight = 0.9) +

  # Leyenda inferior
  annotate("rect", xmin = 1, xmax = 11, ymin = 1.5, ymax = 4, fill = "#FAFAFA", color = "#DDDDDD") +
  annotate("text", x = 6, y = 3.6, label = "Leyenda de columnas", size = 3, color = "#333333", fontface = "bold") +

  annotate("rect", xmin = 1.3, xmax = 2, ymin = 2.8, ymax = 3.2, fill = "#DDE3EA", color = "#1F3864") +
  annotate("text", x = 2.3, y = 3, label = "ID (no editar)", size = 2.3, color = "#1F3864", hjust = 0) +

  annotate("rect", xmin = 4, xmax = 4.7, ymin = 2.8, ymax = 3.2, fill = "#F7F7F8", color = "#808080") +
  annotate("text", x = 5, y = 3, label = "Original (no editar)", size = 2.3, color = "#555555", hjust = 0) +

  annotate("rect", xmin = 7.2, xmax = 7.9, ymin = 2.8, ymax = 3.2, fill = "#C6EFCE", color = "#2E7D32") +
  annotate("text", x = 8.2, y = 3, label = "*_recod (EDITAR)", size = 2.3, color = "#2E7D32", hjust = 0, fontface = "bold") +

  annotate("rect", xmin = 1.3, xmax = 2, ymin = 2.1, ymax = 2.5, fill = "#FCE4D6", color = "#C65911") +
  annotate("text", x = 2.3, y = 2.3, label = "Control (opcional)", size = 2.3, color = "#C65911", hjust = 0) +

  annotate("rect", xmin = 4, xmax = 4.7, ymin = 2.1, ymax = 2.5, fill = "#D9EAD3", color = "#548235") +
  annotate("text", x = 5, y = 2.3, label = "Ejemplo (ignorar)", size = 2.3, color = "#548235", hjust = 0) +

  xlim(0, 12) + ylim(1, 10) +
  theme_diagrama()

path_sm_cols <- save_diagram(diag_sm_cols, "diagrama_sm_columnas", width = 10, height = 6)

# ── Diagrama 10: Ciclo de manipulacion de declaracion FAMILIAS ──────────────

diag_familias_ciclo <- ggplot() +
  draw_box(6, 9.2, 6.8, 0.9, "1) R genera plantilla base de FAMILIAS",
           fill = "#E8F0FE", border = "#4472C4", text_size = 3.0, fontface = "bold") +
  draw_arrow_v(6, 8.75, 8.05) +

  draw_box(6, 7.6, 7.2, 0.95, "2) Usuario edita SOLO campos de relaci\u00f3n\nmodo_so, col_otros, campo_texto (+ incluye_no_incluye opcional)",
           fill = "#F0F7FF", border = "#1F4E79", text_size = 2.6) +

  annotate("segment", x = 6, xend = 3.3, y = 7.1, yend = 6.2, color = "#555555", linewidth = 0.6,
           arrow = arrow(length = unit(0.12, "cm"), type = "closed")) +
  annotate("segment", x = 6, xend = 8.7, y = 7.1, yend = 6.2, color = "#555555", linewidth = 0.6,
           arrow = arrow(length = unit(0.12, "cm"), type = "closed")) +

  draw_box(3.3, 5.7, 4.8, 1.0, "select_one:\nmodo_so (padre/hijo) + col_otros + campo_texto",
           fill = "#DFF5DF", border = "#2E7D32", text_size = 2.5, fontface = "bold") +
  draw_box(8.7, 5.7, 4.8, 1.0, "select_multiple:\ncol_otros (opci\u00f3n) + campo_texto",
           fill = "#FFF2CC", border = "#BF8F00", text_size = 2.5, fontface = "bold") +

  annotate("segment", x = 3.3, xend = 5.4, y = 5.2, yend = 4.35, color = "#555555", linewidth = 0.6,
           arrow = arrow(length = unit(0.12, "cm"), type = "closed")) +
  annotate("segment", x = 8.7, xend = 6.6, y = 5.2, yend = 4.35, color = "#555555", linewidth = 0.6,
           arrow = arrow(length = unit(0.12, "cm"), type = "closed")) +

  draw_box(6, 3.9, 7.4, 0.95, "3) leer_familias_clasificar() adopta las familias\ny arma relaciones padre-hijo para el proceso",
           fill = "#E8F5E8", border = "#2E7D32", text_size = 2.6) +
  draw_arrow_v(6, 3.45, 2.75) +

  draw_box(6, 2.3, 7.4, 0.95, "4) construir_plantilla_desde_familias() crea la plantilla de codificaci\u00f3n",
           fill = "#DCEBFF", border = "#1F4E79", text_size = 2.6, fontface = "bold") +
  draw_arrow_v(6, 1.85, 1.2) +

  draw_box(6, 0.8, 7.2, 0.75, "Salida: plantilla de codificaci\u00f3n lista para el codificador",
           fill = "#E2EFDA", border = "#548235", text_size = 2.5, fontface = "bold") +

  xlim(0.2, 12.2) + ylim(0.35, 9.8) +
  theme_diagrama()

path_familias_ciclo <- save_diagram(diag_familias_ciclo, "diagrama_familias_ciclo", width = 10, height = 5.8)

# ── Diagrama 11: Anatomia de declaracion FAMILIAS ────────────────────────────

diag_familias_anatomia <- ggplot() +
  annotate("text", x = 6, y = 9.5, label = "Anatomia de FAMILIAS: campos que se manipulan",
           size = 4, color = "#1F4E79", fontface = "bold") +

  annotate("rect", xmin = 0.4, xmax = 11.6, ymin = 8.2, ymax = 8.9, fill = "#EEF3FA", color = "#4472C4") +
  annotate("text", x = 1.2, y = 8.55, label = "variable", size = 2.2, color = "#1F4E79", fontface = "bold") +
  annotate("text", x = 2.8, y = 8.55, label = "tipo", size = 2.2, color = "#1F4E79", fontface = "bold") +
  annotate("text", x = 4.3, y = 8.55, label = "modo_so", size = 2.2, color = "#1F4E79", fontface = "bold") +
  annotate("text", x = 6.0, y = 8.55, label = "col_otros", size = 2.2, color = "#1F4E79", fontface = "bold") +
  annotate("text", x = 7.8, y = 8.55, label = "campo_texto", size = 2.2, color = "#1F4E79", fontface = "bold") +
  annotate("text", x = 10.0, y = 8.55, label = "incluye_no_incluye", size = 2.0, color = "#1F4E79", fontface = "bold") +

  annotate("rect", xmin = 0.4, xmax = 11.6, ymin = 5.5, ymax = 8.2, fill = "#FAFAFA", color = "#DDDDDD") +
  annotate("segment", x = c(2.0, 3.5, 5.0, 6.9, 8.8), xend = c(2.0, 3.5, 5.0, 6.9, 8.8),
           y = 5.3, yend = 8.9, color = "#DDDDDD", linewidth = 0.5) +

  annotate("text", x = 1.2, y = 7.7, label = "p12", size = 2.1, color = "#333333") +
  annotate("text", x = 2.8, y = 7.7, label = "select_one", size = 2.0, color = "#333333") +
  annotate("text", x = 4.3, y = 7.7, label = "SO_padre", size = 2.0, color = "#2E7D32", fontface = "bold") +
  annotate("text", x = 6.0, y = 7.7, label = "p12/99", size = 2.0, color = "#333333") +
  annotate("text", x = 7.8, y = 7.7, label = "p12_otro", size = 2.0, color = "#333333") +
  annotate("text", x = 10.0, y = 7.7, label = "-", size = 2.0, color = "#777777") +

  annotate("text", x = 1.2, y = 6.9, label = "p12_otro", size = 2.1, color = "#333333") +
  annotate("text", x = 2.8, y = 6.9, label = "text", size = 2.0, color = "#333333") +
  annotate("text", x = 4.3, y = 6.9, label = "SO_hijo", size = 2.0, color = "#1F4E79", fontface = "bold") +
  annotate("text", x = 6.0, y = 6.9, label = "p12/99", size = 2.0, color = "#333333") +
  annotate("text", x = 7.8, y = 6.9, label = "p12_otro", size = 2.0, color = "#333333") +
  annotate("text", x = 10.0, y = 6.9, label = "-", size = 2.0, color = "#777777") +

  annotate("text", x = 1.2, y = 6.1, label = "p20", size = 2.1, color = "#333333") +
  annotate("text", x = 2.8, y = 6.1, label = "select_multiple", size = 2.0, color = "#333333") +
  annotate("text", x = 4.3, y = 6.1, label = "-", size = 2.0, color = "#777777") +
  annotate("text", x = 6.0, y = 6.1, label = "p20/other", size = 2.0, color = "#333333") +
  annotate("text", x = 7.8, y = 6.1, label = "p20_otro", size = 2.0, color = "#333333") +
  annotate("text", x = 10.0, y = 6.1, label = "incluye", size = 2.0, color = "#BF8F00", fontface = "bold") +

  annotate("rect", xmin = 0.4, xmax = 11.6, ymin = 4.7, ymax = 5.35, fill = "#FFF7E6", color = "#BF8F00") +
  annotate("text", x = 6, y = 5.0, label = "Campos que se editan: modo_so, col_otros, campo_texto (+ incluye_no_incluye opcional)",
           size = 2.35, color = "#7A5A00", fontface = "bold") +

  annotate("rect", xmin = 0.4, xmax = 11.6, ymin = 1.3, ymax = 4.4, fill = "#F9FBFF", color = "#B7C7E3") +
  annotate("text", x = 0.9, y = 4.0, hjust = 0, label = "Reglas clave de familias:",
           size = 2.8, color = "#1F4E79", fontface = "bold") +
  annotate("text", x = 0.9, y = 3.45, hjust = 0, size = 2.4, color = "#333333",
           label = "1) FAMILIAS define relaciones padre-hijo entre seleccion y texto abierto") +
  annotate("text", x = 0.9, y = 3.0, hjust = 0, size = 2.4, color = "#333333",
           label = "2) En select_one se declara modo_so + col_otros + campo_texto") +
  annotate("text", x = 0.9, y = 2.55, hjust = 0, size = 2.4, color = "#333333",
           label = "3) En select_multiple se declara col_otros (opcion) + campo_texto") +
  annotate("text", x = 0.9, y = 2.1, hjust = 0, size = 2.4, color = "#333333",
           label = "4) Luego R adopta familias y define comportamiento de la variable padre") +

  xlim(0, 12) + ylim(1, 9.9) +
  theme_diagrama()

path_familias_anatomia <- save_diagram(diag_familias_anatomia, "diagrama_familias_anatomia", width = 10, height = 6)

message("Diagramas generados en: ", DIAGRAMAS_DIR)

# =============================================================================
# CONSTRUIR PRESENTACION
# =============================================================================

doc <- read_pptx(TEMPLATE_PPTX)
layouts_disponibles <- layout_summary(doc)$layout
message("Layouts disponibles: ", paste(layouts_disponibles, collapse = ", "))

# =============================================================================
# SECCION 0: PORTADA Y VISTA GENERAL
# =============================================================================

# Slide 1: Portada
doc <- add_title_slide(doc,
  title    = "Manual de Codificaci\u00f3n\nPlantilla PPRA",
  subtitle = PROYECTO,
  date     = FECHA
)

# Slide 2: Contenido del manual
doc <- add_section_slide(doc,
  title    = "Contenido del manual",
  subtitle = paste0(
    "Secci\u00f3n 1: Configuraci\u00f3n FAMILIAS y construcci\u00f3n de plantilla\n",
    "  - Elaborar plantilla FAMILIAS\n",
    "  - Definir relaciones padre-hijo y campos de familia\n",
    "  - Adoptar familias en R para construir plantilla de codificaci\u00f3n\n\n",
    "Secci\u00f3n 2: Gu\u00eda para Codificadores (Excel)\n",
    "  - Estructura de la plantilla de codificaci\u00f3n\n",
    "  - C\u00f3mo codificar cada tipo de variable\n\n",
    "Secci\u00f3n 3: Aplicaci\u00f3n t\u00e9cnica en R\n",
    "  - Aplicar codificaci\u00f3n a datos e instrumento\n",
    "  - Troubleshooting y cierre"
  )
)

# Slide 3: DIAGRAMA flujo general
doc <- add_full_diagram_slide(doc,
  title    = "Flujo general de codificaci\u00f3n",
  diagram_asset = path_flujo_general
)

# Slide 4: Logica general por etapas
doc <- add_text_slide(doc,
  title = "L\u00f3gica general de implementaci\u00f3n (sin c\u00f3digo literal)",
  content_block = block_list(
    fpar(ftext("Etapa A: Elaborar y completar FAMILIAS", fp_bold)),
    fpar(ftext("Etapa B: Adoptar FAMILIAS en R y construir plantilla de codificaci\u00f3n", fp_body)),
    fpar(ftext("Etapa C: Codificador completa *_recod en Excel", fp_body)),
    fpar(ftext("Etapa D: Aplicar codificaci\u00f3n a datos (R)", fp_body)),
    fpar(ftext("Etapa E: Adaptar instrumento (R)", fp_body)),
    fpar(ftext("Etapa F: Revisi\u00f3n final y salida para an\u00e1lisis", fp_body))
  )
)

# =============================================================================
# SECCION 1: CONFIGURACION FAMILIAS Y CONSTRUCCION DE PLANTILLA
# =============================================================================

# Slide 4: Seccion
doc <- add_section_slide(doc,
  title    = "Secci\u00f3n 1: Configuraci\u00f3n FAMILIAS y construcci\u00f3n de plantilla",
  subtitle = "Orden recomendado: elaborar FAMILIAS, definir familias padre-hijo y luego adoptarlas en R para construir la plantilla de codificaci\u00f3n."
)

# Slide 5: Declaracion FAMILIAS en R
doc <- add_text_slide(doc,
  title = "Paso 1: Elaborar plantilla FAMILIAS",
  content_block = block_list(
    fpar(ftext("FAMILIAS no es una validaci\u00f3n de errores.", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Su objetivo es definir relaciones entre:", fp_body)),
    fpar(ftext("  \u2022 Variables de selecci\u00f3n (padre)", fp_body)),
    fpar(ftext("  \u2022 Variables de texto abierto asociadas (hijo)", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Esa relaci\u00f3n influye en el comportamiento de la variable padre.", fp_tip)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Primero se genera/basea la plantilla FAMILIAS, luego se edita.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Despu\u00e9s R adopta esas familias para construir la plantilla de codificaci\u00f3n.", fp_body))
  )
)

# Slide 6: Sub-seccion manipulacion FAMILIAS
doc <- add_section_slide(doc,
  title    = "1.1 Manipulaci\u00f3n de la declaraci\u00f3n FAMILIAS",
  subtitle = "Se editan 3 campos principales (+1 opcional) para declarar correctamente cada familia."
)

# Slide 7: Diagrama ciclo de manipulacion
doc <- add_full_diagram_slide(doc,
  title    = "Ciclo de manipulaci\u00f3n de FAMILIAS",
  diagram_asset = path_familias_ciclo
)

# Slide 8: Diagrama anatomia de declaracion
doc <- add_full_diagram_slide(doc,
  title    = "Anatom\u00eda operativa de FAMILIAS",
  diagram_asset = path_familias_anatomia
)

# Slide 9: Ejecucion practica de edicion
doc <- add_img_text_r_slide(doc,
  title    = "Edici\u00f3n operativa de FAMILIAS (ejemplo)",
  img_path = get_screenshot("familias_edicion_real"),
  content_block = block_list(
    fpar(ftext("Campos que el usuario modifica:", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("1. ", fp_body), ftext("modo_so", fp_code_b), ftext(" (solo en select_one: padre o hijo)", fp_body)),
    fpar(ftext("2. ", fp_body), ftext("col_otros", fp_code_b), ftext(" (opci\u00f3n que dispara texto abierto)", fp_body)),
    fpar(ftext("3. ", fp_body), ftext("campo_texto", fp_code_b), ftext(" (variable abierta vinculada)", fp_body)),
    fpar(ftext("4. ", fp_body), ftext("incluye_no_incluye", fp_code_b), ftext(" (opcional)", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Regla operativa:", fp_bold)),
    fpar(ftext("  select_one: modo_so + col_otros + campo_texto", fp_body)),
    fpar(ftext("  select_multiple: col_otros + campo_texto", fp_body))
  )
)

# Slide 10: Leer FAMILIAS para construir plantilla de codificacion
doc <- add_img_text_r_slide(doc,
  title    = "Paso 3: Leer FAMILIAS en R y construir plantilla de codificaci\u00f3n",
  img_path = get_screenshot("codigo_generar_plantilla"),
  content_block = block_list(
    fpar(ftext("L\u00f3gica de etapas en R:", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("leer_familias_clasificar()", fp_code_b)),
    fpar(ftext("  Adopta las familias declaradas (padre-hijo)", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("construir_plantilla_desde_familias()", fp_code_b)),
    fpar(ftext("  Aplica familias sobre instrumento + datos", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("exportar_plantilla_codificacion_xlsx()", fp_code_b)),
    fpar(ftext("  Genera la plantilla que recibir\u00e1 el codificador", fp_body))
  )
)

# =============================================================================
# SECCION 2: GUIA PARA CODIFICADORES (EXCEL)
# =============================================================================

# Slide 4: Sección
doc <- add_section_slide(doc,
  title    = "Secci\u00f3n 2: Gu\u00eda para Codificadores",
  subtitle = "Todo lo que necesita saber para llenar la plantilla Excel de codificaci\u00f3n.\nNo se requiere conocimiento de R."
)

# Slide 5: Navegación
doc <- add_img_text_r_slide(doc,
  title    = "Hoja NAVEGACI\u00d3N",
  img_path = get_screenshot("hoja_navegacion"),
  content_block = block_list(
    fpar(ftext("La hoja NAVEGACI\u00d3N es su punto de partida.", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Contiene un \u00edndice con hipervinculos a cada hoja de variable.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Cada fila muestra:", fp_body)),
    fpar(ftext("  \u2022 Nombre de la variable", fp_body)),
    fpar(ftext("  \u2022 Tipo (select_multiple, select_one, text, integer)", fp_body)),
    fpar(ftext("  \u2022 N\u00famero de registros", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Haga clic en el nombre para ir directamente a la hoja.", fp_tip))
  )
)

# Slide 6: Instrucciones
doc <- add_img_text_r_slide(doc,
  title    = "Hoja INSTRUCCIONES",
  img_path = get_screenshot("hoja_instrucciones"),
  content_block = block_list(
    fpar(ftext("Resumen de las 10 reglas principales:", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("1. NO modifique la fila 1 ni la fila 2", fp_warning)),
    fpar(ftext("2. Edite SOLO columnas *_recod y Control", fp_body)),
    fpar(ftext("3. select_multiple: 1=marcar, 0=desmarcar, vac\u00edo=sin cambio", fp_body)),
    fpar(ftext("4. select_one/text/integer: escriba c\u00f3digo en *_recod", fp_body)),
    fpar(ftext("5. Nuevas opciones SM: columna con patr\u00f3n <parent>/<cod>_recod", fp_body)),
    fpar(ftext("6. Nuevas categor\u00edas SO/INT: bloque auxiliar derecho", fp_body)),
    fpar(ftext("7. Columna ejemplo_recod es solo referencia", fp_body))
  )
)

# ── 1.2 Estructura General ───────────────────────────────────────────────────

# Slide 7: DIAGRAMA mapa de zonas
doc <- add_full_diagram_slide(doc,
  title    = "Mapa de zonas de cada hoja",
  diagram_asset = path_zonas
)

# Slide 8: Screenshot estructura general anotada
doc <- add_full_img_slide(doc,
  title    = "Estructura general de cada hoja (ejemplo real)",
  img_path = get_screenshot("estructura_general_anotada")
)

# Slide 9: Filas 1 y 2
doc <- add_img_text_l_slide(doc,
  title    = "Fila 1 (t\u00e9cnica) vs Fila 2 (etiquetas)",
  img_path = get_screenshot("filas_1_y_2"),
  content_block = block_list(
    fpar(ftext("Fila 1: Nombre t\u00e9cnico", fp_bold)),
    fpar(ftext("Es el identificador que usa el paquete R.", fp_body)),
    fpar(ftext("NUNCA la modifique.", fp_warning)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Fila 2: Etiqueta legible", fp_bold)),
    fpar(ftext("Describe la columna en lenguaje natural.", fp_body)),
    fpar(ftext("Sirve de gu\u00eda visual. Tampoco debe editarse.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Los datos comienzan en la fila 3.", fp_tip))
  )
)

# Slide 10: Zonas de color (screenshot)
doc <- add_img_text_r_slide(doc,
  title    = "Zonas de color en cada hoja",
  img_path = get_screenshot("zonas_de_color"),
  content_block = block_list(
    fpar(ftext("C\u00f3digo de colores:", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u25a0 Azul oscuro: Columnas ID", fp_text(font.size = 14, color = "#1F3864", font.family = "Calibri"))),
    fpar(ftext("  (_uuid, _index) \u2014 NO editar", fp_small)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u25a0 Gris claro: Referencia", fp_text(font.size = 14, color = "#808080", font.family = "Calibri"))),
    fpar(ftext("  (valores originales) \u2014 NO editar", fp_small)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u25a0 Verde claro: Editables", fp_text(font.size = 14, color = "#2E7D32", font.family = "Calibri"))),
    fpar(ftext("  (*_recod) \u2014 AQU\u00cd SE CODIFICA", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u25a0 Naranja: Control / notas", fp_text(font.size = 14, color = "#C65911", font.family = "Calibri"))),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u25a0 Rojo: Bloque auxiliar", fp_text(font.size = 14, color = "#C00000", font.family = "Calibri")))
  )
)

# Slide 11: DIAGRAMA árbol de decisión
doc <- add_full_diagram_slide(doc,
  title    = "\u00bfQu\u00e9 tipo de variable estoy codificando?",
  diagram_asset = path_decision
)

# ── 1.3 select_multiple ──────────────────────────────────────────────────────

# Slide 12: Sección SM
doc <- add_section_slide(doc,
  title    = "Tipo: select_multiple",
  subtitle = "Preguntas de opci\u00f3n m\u00faltiple (marcar varias respuestas)\nColor identificador: azul (#DCEBFF)"
)

# Slide 13: DIAGRAMA flujo SM
doc <- add_full_diagram_slide(doc,
  title    = "select_multiple \u2014 Proceso paso a paso",
  diagram_asset = path_sm_flujo
)

# Slide 14: DIAGRAMA anatomía columnas SM
doc <- add_full_diagram_slide(doc,
  title    = "select_multiple \u2014 Anatom\u00eda de columnas",
  diagram_asset = path_sm_cols
)

# Slide 15: Screenshot vista SM
doc <- add_img_text_r_slide(doc,
  title    = "select_multiple \u2014 Vista general",
  img_path = get_screenshot("sm_vista_general"),
  content_block = block_list(
    fpar(ftext("Cada opci\u00f3n del cat\u00e1logo es una columna.", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("C\u00f3mo codificar:", fp_bold)),
    fpar(ftext("  \u2022 ", fp_body), ftext("1", fp_code_b), ftext(" = marcar (activar opci\u00f3n)", fp_body)),
    fpar(ftext("  \u2022 ", fp_body), ftext("0", fp_code_b), ftext(" = desmarcar (desactivar)", fp_body)),
    fpar(ftext("  \u2022 vac\u00edo = sin cambio", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Las columnas *_recod van junto a cada opci\u00f3n.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Seleccionadas y Seleccionadas_cod son referencia.", fp_small))
  )
)

# Slide 16: SM nueva opción
doc <- add_img_text_l_slide(doc,
  title    = "select_multiple \u2014 Agregar opci\u00f3n nueva",
  img_path = get_screenshot("sm_nueva_opcion"),
  content_block = block_list(
    fpar(ftext("Para agregar una opci\u00f3n nueva:", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("1. Inserte una columna nueva.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("2. Fila 1:", fp_body)),
    fpar(ftext("   <parent>/<nuevo_codigo>_recod", fp_code)),
    fpar(ftext("   Ej: ", fp_body), ftext("p8/99_recod", fp_code_b)),
    fpar(ftext("", fp_body)),
    fpar(ftext("3. Fila 2: etiqueta visible", fp_body)),
    fpar(ftext("   Ej: ", fp_body), ftext("Otro servicio", fp_code)),
    fpar(ftext("", fp_body)),
    fpar(ftext("4. Datos: 1 / 0 / vac\u00edo.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("ejemplo_recod es solo referencia.", fp_small))
  )
)

# Slide 17: SM antes y después
doc <- add_two_img_slide(doc,
  title     = "select_multiple \u2014 Antes y despu\u00e9s",
  img_left  = get_screenshot("sm_antes"),
  img_right = get_screenshot("sm_despues")
)

# ── 1.4 select_one ───────────────────────────────────────────────────────────

# Slide 18: Sección SO
doc <- add_section_slide(doc,
  title    = "Tipo: select_one",
  subtitle = "Preguntas de opci\u00f3n \u00fanica (una sola respuesta)\nColor identificador: verde (#DFF5DF)"
)

# Slide 19: DIAGRAMA flujo SO padre vs hijo
doc <- add_full_diagram_slide(doc,
  title    = "select_one \u2014 Modo padre vs modo hijo",
  diagram_asset = path_so_flujo
)

# Slide 20: Screenshot SO padre
doc <- add_img_text_r_slide(doc,
  title    = "select_one (padre) \u2014 Vista general",
  img_path = get_screenshot("so_padre_vista"),
  content_block = block_list(
    fpar(ftext("Modo padre: recodifica la variable original.", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Columna editable:", fp_body)),
    fpar(ftext("  Recodificaci\u00f3n (c\u00f3digo)", fp_code)),
    fpar(ftext("  Escriba el c\u00f3digo final deseado.", fp_body)),
    fpar(ftext("  Deje vac\u00edo si no hay cambio.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Columna de referencia:", fp_body)),
    fpar(ftext("  Selecci\u00f3n (c\u00f3digo) \u2014 valor original", fp_small)),
    fpar(ftext("  Selecci\u00f3n (label) \u2014 etiqueta original", fp_small)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Bloque auxiliar (derecha) para nuevas categor\u00edas.", fp_tip))
  )
)

# Slide 21: DIAGRAMA bloque auxiliar detallado
doc <- add_full_diagram_slide(doc,
  title    = "Bloque auxiliar \u2014 C\u00f3mo declarar categor\u00edas nuevas",
  diagram_asset = path_auxiliar
)

# Slide 22: Screenshot bloque auxiliar SO
doc <- add_img_text_l_slide(doc,
  title    = "select_one \u2014 Bloque auxiliar (ejemplo real)",
  img_path = get_screenshot("so_bloque_auxiliar"),
  content_block = block_list(
    fpar(ftext("Bloque auxiliar (derecha de la hoja):", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Dos columnas:", fp_body)),
    fpar(ftext("  \u2022 ", fp_body), ftext("nuevo_codigo", fp_code_b), ftext(" \u2014 c\u00f3digo num\u00e9rico", fp_body)),
    fpar(ftext("  \u2022 ", fp_body), ftext("nueva_etiqueta", fp_code_b), ftext(" \u2014 texto", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Declare cada c\u00f3digo UNA SOLA VEZ.", fp_warning)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Si un c\u00f3digo tiene 2 etiquetas, el adaptador dar\u00e1 error.", fp_small))
  )
)

# Slide 23: Screenshot modo hijo
doc <- add_img_text_r_slide(doc,
  title    = "select_one (hijo) \u2014 Texto asociado",
  img_path = get_screenshot("so_hijo_vista"),
  content_block = block_list(
    fpar(ftext("Modo hijo: recodifica el texto abierto.", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Se usa cuando 'Otro' tiene texto libre.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("La columna *_recod agrupa respuestas", fp_body)),
    fpar(ftext("de texto en categor\u00edas.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("El bloque auxiliar funciona igual:", fp_body)),
    fpar(ftext("nuevo_codigo + nueva_etiqueta.", fp_code)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Diferencia: aqu\u00ed se recodifica el texto,", fp_bold)),
    fpar(ftext("no la variable padre.", fp_bold))
  )
)

# ── 1.5 text & integer ───────────────────────────────────────────────────────

# Slide 24: Sección text
doc <- add_section_slide(doc,
  title    = "Tipos: text e integer",
  subtitle = "Preguntas abiertas de texto libre y num\u00e9ricas\nColores: amarillo (#FFF2CC) y morado (#E6D9F2)"
)

# Slide 25: DIAGRAMA flujo text/integer
doc <- add_full_diagram_slide(doc,
  title    = "text e integer \u2014 Proceso comparado",
  diagram_asset = path_text_int
)

# Slide 26: Screenshot text
doc <- add_img_text_r_slide(doc,
  title    = "text \u2014 C\u00f3mo codificar",
  img_path = get_screenshot("text_vista"),
  content_block = block_list(
    fpar(ftext("Hojas de texto: respuesta abierta original.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Columna editable:", fp_bold)),
    fpar(ftext("  ", fp_body), ftext("<variable>_recod", fp_code)),
    fpar(ftext("  Escriba c\u00f3digo que represente la respuesta.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Bloque auxiliar:", fp_bold)),
    fpar(ftext("  ", fp_body), ftext("nuevo_codigo", fp_code), ftext(" + ", fp_body), ftext("nueva_etiqueta", fp_code)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Deje vac\u00edo si no requiere codificaci\u00f3n.", fp_tip))
  )
)

# Slide 27: Screenshot integer
doc <- add_img_text_r_slide(doc,
  title    = "integer \u2014 C\u00f3mo codificar",
  img_path = get_screenshot("integer_vista"),
  content_block = block_list(
    fpar(ftext("Similar a text: el valor original es un n\u00famero.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Columna editable:", fp_bold)),
    fpar(ftext("  ", fp_body), ftext("<variable>_recod", fp_code)),
    fpar(ftext("  Asigne c\u00f3digo de categor\u00eda.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Bloque auxiliar:", fp_bold)),
    fpar(ftext("  ", fp_body), ftext("nuevo_codigo", fp_code), ftext(" + ", fp_body), ftext("nueva_etiqueta", fp_code)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Si 2 variables comparten diccionario,", fp_body)),
    fpar(ftext("la lista se comparte autom\u00e1ticamente.", fp_body))
  )
)

# ── 1.7 Errores Comunes y Checklist ─────────────────────────────────────────

# Slide 28: Errores comunes
doc <- add_text_slide(doc,
  title = "Errores comunes a evitar",
  content_block = block_list(
    fpar(ftext("\u2718  ", fp_warning), ftext("Editar la fila 1 (nombre t\u00e9cnico)", fp_body)),
    fpar(ftext("     Si se modifica, el adaptador no reconocer\u00e1 la columna.", fp_small)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u2718  ", fp_warning), ftext("Editar columnas de referencia (gris)", fp_body)),
    fpar(ftext("     Los valores originales son solo consulta.", fp_small)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u2718  ", fp_warning), ftext("Usar c\u00f3digo nuevo sin declarar etiqueta", fp_body)),
    fpar(ftext("     Todo c\u00f3digo nuevo DEBE tener etiqueta en el bloque auxiliar.", fp_small)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u2718  ", fp_warning), ftext("Declarar el mismo c\u00f3digo con 2 etiquetas distintas", fp_body)),
    fpar(ftext("     El adaptador devolver\u00e1 error si detecta inconsistencia.", fp_small)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u2718  ", fp_warning), ftext("Usar valores distintos de 1, 0 o vac\u00edo en select_multiple", fp_body)),
    fpar(ftext("     Solo se aceptan estos tres valores.", fp_small))
  )
)

# Slide 29: Checklist final
doc <- add_text_slide(doc,
  title = "Checklist antes de entregar",
  content_block = block_list(
    fpar(ftext("Verifique antes de enviar la plantilla completada:", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u2610  ", fp_body), ftext("Todas las columnas *_recod fueron revisadas", fp_body)),
    fpar(ftext("     (incluso si se dejan vac\u00edas intencionalmente)", fp_small)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u2610  ", fp_body), ftext("Cada c\u00f3digo nuevo tiene su etiqueta en el bloque auxiliar", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u2610  ", fp_body), ftext("No hay etiquetas duplicadas para un mismo c\u00f3digo", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u2610  ", fp_body), ftext("Columnas nuevas en SM: patr\u00f3n <parent>/<cod>_recod", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u2610  ", fp_body), ftext("Control / notas con observaciones donde corresponda", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u2610  ", fp_body), ftext("No se editaron filas 1/2 ni columnas de referencia", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("\u2610  ", fp_body), ftext("Archivo guardado como .xlsx (no .xls ni .csv)", fp_body))
  )
)

# =============================================================================
# SECCION 3: APLICACION TECNICA EN R
# =============================================================================

# Slide 30: Seccion
doc <- add_section_slide(doc,
  title    = "Secci\u00f3n 3: Aplicaci\u00f3n t\u00e9cnica en R",
  subtitle = "Con la plantilla de codificaci\u00f3n completada, aplique cambios a datos e instrumento."
)

# Slide 31: Flujo tecnico de aplicacion
doc <- add_full_diagram_slide(doc,
  title    = "Flujo t\u00e9cnico en R \u2014 Aplicaci\u00f3n y actualizaci\u00f3n",
  diagram_asset = path_flujo_r
)

# Slide 40: Aplicar datos
doc <- add_img_text_r_slide(doc,
  title    = "Fase 3a: Aplicar codificaci\u00f3n a datos",
  img_path = get_screenshot("codigo_adaptar_data"),
  content_block = block_list(
    fpar(ftext("ppra_adaptar_data()", fp_code_b)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Argumentos principales:", fp_bold)),
    fpar(ftext("  path_instrumento", fp_code), ftext(" \u2014 XLSForm original", fp_body)),
    fpar(ftext("  path_datos", fp_code), ftext(" \u2014 datos crudos", fp_body)),
    fpar(ftext("  path_plantilla", fp_code), ftext(" \u2014 plantilla completada", fp_body)),
    fpar(ftext("  sm_vars", fp_code), ftext(" \u2014 select_multiple", fp_body)),
    fpar(ftext("  so_parent_vars", fp_code), ftext(" \u2014 SO padre", fp_body)),
    fpar(ftext("  so_child_vars", fp_code), ftext(" \u2014 SO hijo", fp_body)),
    fpar(ftext("  int_vars", fp_code), ftext(" \u2014 integer", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Resultado: *_recod junto al parent.", fp_tip))
  )
)

# Slide 41: Aplicar instrumento
doc <- add_img_text_r_slide(doc,
  title    = "Fase 3b: Adaptar el instrumento",
  img_path = get_screenshot("codigo_adaptar_instrumento"),
  content_block = block_list(
    fpar(ftext("ppra_adaptar_instrumento()", fp_code_b)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Actualiza el XLSForm:", fp_bold)),
    fpar(ftext("  \u2022 Agrega preguntas *_recod al survey", fp_body)),
    fpar(ftext("  \u2022 Agrega c\u00f3digos al choices", fp_body)),
    fpar(ftext("  \u2022 Colorea por tipo:", fp_body)),
    fpar(ftext("    SM = verde, SO = azul, INT = morado", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("choices_order:", fp_bold)),
    fpar(ftext("  original_first", fp_code), ftext(" \u2014 original primero", fp_body)),
    fpar(ftext("  by_first_seen", fp_code), ftext(" \u2014 aparici\u00f3n en datos", fp_body)),
    fpar(ftext("  alphabetical", fp_code), ftext(" \u2014 alfab\u00e9tico", fp_body))
  )
)

# Slide 42: Output
doc <- add_img_text_l_slide(doc,
  title    = "Resultado: datos adaptados",
  img_path = get_screenshot("output_datos_recod"),
  content_block = block_list(
    fpar(ftext("El archivo de salida preserva:", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("  \u2022 Todas las hojas originales", fp_body)),
    fpar(ftext("  \u2022 Todas las columnas originales", fp_body)),
    fpar(ftext("  \u2022 *_recod insertadas al lado del parent", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Colores en el output:", fp_bold)),
    fpar(ftext("  Verde: columnas SM recodificadas", fp_body)),
    fpar(ftext("  Azul: columnas SO recodificadas", fp_body)),
    fpar(ftext("  Morado: columnas INT recodificadas", fp_body))
  )
)

# Slide 43: Troubleshooting
doc <- add_text_slide(doc,
  title = "Troubleshooting \u2014 Errores comunes en R",
  content_block = block_list(
    fpar(ftext("Error: etiqueta faltante para c\u00f3digo nuevo", fp_warning)),
    fpar(ftext("  Codificador us\u00f3 c\u00f3digo sin declarar en bloque auxiliar.", fp_body)),
    fpar(ftext("  Soluci\u00f3n: completar nuevo_codigo + nueva_etiqueta.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Error: c\u00f3digos duplicados con etiquetas distintas", fp_warning)),
    fpar(ftext("  Mismo c\u00f3digo con 2+ etiquetas en bloque auxiliar.", fp_body)),
    fpar(ftext("  Soluci\u00f3n: unificar a una sola etiqueta.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Error: columna no encontrada en la hoja", fp_warning)),
    fpar(ftext("  Variable en sm_vars/so_* sin hoja en plantilla.", fp_body)),
    fpar(ftext("  Soluci\u00f3n: verificar nombre exacto.", fp_body)),
    fpar(ftext("", fp_body)),
    fpar(ftext("Error: fila 1 modificada", fp_warning)),
    fpar(ftext("  Codificador alter\u00f3 nombres t\u00e9cnicos. Corregir o regenerar.", fp_body))
  )
)

# Slide 44: Resumen flujo completo
doc <- add_text_slide(doc,
  title = "Flujo completo \u2014 Resumen",
  content_block = block_list(
    fpar(ftext("Paso a paso:", fp_bold)),
    fpar(ftext("", fp_body)),
    fpar(ftext("1. ", fp_bold), ftext("Preparar FAMILIAS.xlsx", fp_body)),
    fpar(ftext("   Variables, tipos y relaciones parent-child.", fp_small)),
    fpar(ftext("", fp_body)),
    fpar(ftext("2. ", fp_bold), ftext("Generar plantilla de codificaci\u00f3n", fp_body)),
    fpar(ftext("   leer_familias \u2192 construir_plantilla \u2192 exportar", fp_code)),
    fpar(ftext("", fp_body)),
    fpar(ftext("3. ", fp_bold), ftext("Enviar plantilla al codificador", fp_body)),
    fpar(ftext("   Llena *_recod y declara nuevas categor\u00edas (no necesita R).", fp_small)),
    fpar(ftext("", fp_body)),
    fpar(ftext("4. ", fp_bold), ftext("Aplicar codificaci\u00f3n a los datos", fp_body)),
    fpar(ftext("   ppra_adaptar_data() \u2192 datos con *_recod", fp_code)),
    fpar(ftext("", fp_body)),
    fpar(ftext("5. ", fp_bold), ftext("Adaptar el instrumento", fp_body)),
    fpar(ftext("   ppra_adaptar_instrumento() \u2192 XLSForm actualizado", fp_code)),
    fpar(ftext("", fp_body)),
    fpar(ftext("6. ", fp_bold), ftext("Continuar con an\u00e1lisis", fp_body)),
    fpar(ftext("   Datos e instrumento alimentan reportes, dashboards, SPSS...", fp_small))
  )
)

# ── Exportar ──────────────────────────────────────────────────────────────────

print(doc, target = OUTPUT_PATH)
doc_out <- read_pptx(OUTPUT_PATH)
total_slides <- length(unique(pptx_summary(doc_out)$slide_id))
message("\n\u2705 Manual generado exitosamente en:\n   ", OUTPUT_PATH)
message("   Total de slides: ", total_slides)
message("\n\ud83d\udccc Diagramas generados en: ", DIAGRAMAS_DIR)
message("\ud83d\udccc Screenshots (placeholders) en: ", SCREENSHOTS_DIR)
message("\n   Para mejorar: reemplace los placeholders PNG con screenshots reales")
message("   y vuelva a ejecutar este script.\n")
