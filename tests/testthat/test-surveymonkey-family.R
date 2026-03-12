test_that("familia surveymonkey_ genera XLSForm de referencia y data compatible", {
  path_sav <- tempfile(fileext = ".sav")
  path_xlsx <- tempfile(fileext = ".xlsx")
  path_codif <- tempfile(fileext = ".xlsx")
  path_familias <- tempfile(fileext = ".xlsx")
  path_freq <- tempfile(fileext = ".xlsx")
  path_cross <- tempfile(fileext = ".xlsx")

  on.exit(unlink(c(path_sav, path_xlsx, path_codif, path_familias, path_freq, path_cross)), add = TRUE)

  likert_labs <- c(
    "Totalmente en Desacuerdo" = 1,
    "En Desacuerdo" = 2,
    "De Acuerdo" = 3,
    "Totalmente de Acuerdo" = 4,
    "SIN INF" = 99
  )

  df <- data.frame(
    CollectorNm = c("COL-01", "COL-02", "COL-03"),
    respondent_id = c("r1", "r2", "r3"),
    date_created = c(
      "10/27/2025 07:04:32 PM",
      "10/28/2025 08:10:11 PM",
      "10/29/2025 09:15:55 PM"
    ),
    email_address = c("a@pucp.edu.pe", "b@pucp.edu.pe", "c@pucp.edu.pe"),
    Sexo = c("Femenino", "Masculino", "Femenino"),
    P1 = haven::labelled(c(1, 2, 1), c("Sí" = 1, "No" = 2)),
    P4_1 = haven::labelled(c(1, 2, 4), likert_labs),
    P4_2 = haven::labelled(c(2, 3, 99), likert_labs),
    P5_1 = haven::labelled(c(1, NA, 1), c("Página web" = 1)),
    P5_2 = haven::labelled(c(NA, NA, 1), c("Redes sociales" = 1)),
    P5_3 = haven::labelled(c(NA, 1, NA), c("Otro:" = 1)),
    P5_O = c(NA, "Club de alumnos", NA),
    P6 = haven::labelled(c(1, 3, 2), c("Presencial" = 1, "Virtual" = 2, "Otro:" = 3)),
    P6_O = c(NA, "Mixta", NA),
    Edad_HO = c(20, 21, 22),
    TOTAL = c(10, 20, 30),
    stringsAsFactors = FALSE
  )

  attr(df$P1, "label") <- "¿Desea continuar?"
  attr(df$Sexo, "label") <- "Sexo"
  attr(df$P4_1, "label") <- "La misión está claramente definida"
  attr(df$P4_2, "label") <- "Los canales de difusión son adecuados"
  attr(df$P5_1, "label") <- "¿A través de qué medios se informó?"
  attr(df$P5_2, "label") <- "¿A través de qué medios se informó?"
  attr(df$P5_3, "label") <- "¿A través de qué medios se informó?"
  attr(df$P5_O, "label") <- "Other (please specify)"
  attr(df$P6, "label") <- "Modalidad del servicio"
  attr(df$P6_O, "label") <- "Other (please specify)"
  attr(df$Edad_HO, "label") <- "Edad"
  attr(df$TOTAL, "label") <- "Total"

  haven::write_sav(df, path_sav)

  sm <- prosecnur::surveymonkey_leer(path_sav)

  expect_s3_class(sm, "prosecnur_surveymonkey")
  expect_true(all(c(
    "name_raw", "label", "class", "n_value_labels", "is_labelled",
    "stem", "suffix", "kind_guess", "is_other", "group_guess", "order"
  ) %in% names(sm$vars_tbl)))

  kinds <- stats::setNames(sm$vars_tbl$kind_guess, sm$vars_tbl$name_raw)
  expect_identical(kinds[["respondent_id"]], "metadata")
  expect_identical(kinds[["Sexo"]], "select_one")
  expect_identical(kinds[["P1"]], "select_one")
  expect_identical(kinds[["P4_1"]], "battery_item")
  expect_identical(kinds[["P5_1"]], "select_multiple_dummy")
  expect_identical(kinds[["P5_O"]], "other_text")
  expect_true(sm$vars_tbl$is_auxiliary[sm$vars_tbl$name_raw == "TOTAL"])

  inst_ref <- prosecnur::surveymonkey_xlsform(sm, path = path_xlsx)
  idx_grp_p4 <- which(inst_ref$survey$name == "grp_p4")[1]
  idx_p5_other <- which(inst_ref$survey$name == "p5_other")[1]
  idx_p6_other <- which(inst_ref$survey$name == "p6_other")[1]

  expect_s3_class(inst_ref, "prosecnur_surveymonkey_xlsform")
  expect_true(file.exists(path_xlsx))
  expect_true(all(c("survey", "choices", "settings", "diagnostico") %in% names(inst_ref)))
  expect_true(any(inst_ref$survey$type == "begin_group"))
  expect_true(any(grepl("^select_multiple ", inst_ref$survey$type)))
  expect_true(any(inst_ref$survey$section == "survey_monkey_auxiliary"))
  expect_true(any(inst_ref$survey$name == "p6_other"))
  expect_true(any(inst_ref$survey$name == "sexo"))
  expect_true(any(inst_ref$survey$type == "select_one lst_sexo"))
  expect_true(is.na(inst_ref$survey$`label::es`[idx_grp_p4]))
  expect_identical(inst_ref$survey$`label::es`[idx_p5_other], "Otro:")
  expect_identical(inst_ref$survey$`label::es`[idx_p6_other], "Otro:")
  expect_true("lst_si_no" %in% inst_ref$choices$list_name)
  expect_true("lst_sexo" %in% inst_ref$choices$list_name)
  expect_true("lst_acuerdo_4" %in% inst_ref$choices$list_name)
  expect_true("lst_p6" %in% inst_ref$choices$list_name)
  expect_true("lst_p5" %in% inst_ref$choices$list_name)
  expect_true(all(c("femenino", "masculino") %in% inst_ref$choices$name[inst_ref$choices$list_name == "lst_sexo"]))
  expect_true("99" %in% as.character(inst_ref$choices$name[inst_ref$choices$list_name == "lst_acuerdo_4"]))

  rp_inst <- prosecnur::reporte_instrumento(path_xlsx, lang = "es")
  idx_rp_p1 <- which(rp_inst$survey$name == "p1")[1]
  idx_rp_sexo <- which(rp_inst$survey$name == "sexo")[1]
  idx_rp_p4_1 <- which(rp_inst$survey$name == "p4_1")[1]
  idx_rp_p4_2 <- which(rp_inst$survey$name == "p4_2")[1]
  idx_rp_p6 <- which(rp_inst$survey$name == "p6")[1]
  expect_s3_class(rp_inst, "prosecnur_instrumento")
  expect_true("p5" %in% rp_inst$survey$name)
  expect_true("sexo" %in% rp_inst$survey$name)
  expect_true("p1" %in% rp_inst$survey$name)
  expect_identical(
    rp_inst$survey$list_name[idx_rp_sexo],
    "lst_sexo"
  )
  expect_identical(
    rp_inst$survey$list_name[idx_rp_p1],
    "lst_si_no"
  )
  expect_identical(
    rp_inst$survey$list_name[idx_rp_p4_1],
    "lst_acuerdo_4"
  )
  expect_identical(
    rp_inst$survey$list_name[idx_rp_p4_1],
    rp_inst$survey$list_name[idx_rp_p4_2]
  )
  expect_identical(
    rp_inst$survey$list_name[idx_rp_p6],
    "lst_p6"
  )

  dat_ref <- prosecnur::surveymonkey_data(sm)
  expect_equal(nrow(dat_ref), nrow(df))
  expect_true(all(c(
    "p5", "p5/1", "p5/2", "p5/3", "p5/Other", "p5_other", "p6_other"
  ) %in% names(dat_ref)))
  expect_equal(dat_ref$sexo, c("femenino", "masculino", "femenino"))
  expect_false(any(c("p5_1", "p5_2", "p5_3", "p5_o", "p6_o") %in% names(dat_ref)))
  expect_equal(dat_ref$p5, c("1", "3", "1 2"))
  expect_equal(as.numeric(dat_ref$p4_2), c(2, 3, 99))
  expect_identical(dat_ref$p6_other[2], "Mixta")
  expect_identical(as.character(dat_ref[["p5/Other"]]), as.character(df$P5_3))
  expect_identical(dat_ref$p5_other[2], "Club de alumnos")

  openxlsx::write.xlsx(list(data = dat_ref), file = path_codif, overwrite = TRUE)
  inst_codif <- prosecnur::leer_instrumento_xlsform(path_xlsx)
  dat_codif_obj <- prosecnur::leer_datos(path_codif)
  prosecnur::escribir_plantilla_familias(inst_codif, dat_codif_obj, path = path_familias)
  familias <- prosecnur::leer_familias_clasificar(
    path = path_familias,
    inst = inst_codif,
    dat = dat_codif_obj,
    verbose = FALSE
  )
  row_p5 <- familias$familias_filtradas[familias$familias_filtradas$parent == "p5", , drop = FALSE]
  row_p6 <- familias$familias_filtradas[familias$familias_filtradas$parent == "p6", , drop = FALSE]
  expect_true(nrow(row_p5) == 1L)
  expect_identical(row_p5$tipo[1], "select_multiple")
  expect_identical(row_p5$other_dummy_col[1], "p5/Other")
  expect_identical(row_p5$text_col[1], "p5_other")
  expect_true(nrow(row_p6) == 1L)
  expect_identical(row_p6$tipo[1], "select_one")
  expect_identical(row_p6$text_col[1], "p6_other")

  rp_data <- prosecnur::reporte_data(dat_ref, rp_inst)
  expect_s3_class(rp_data, "prosecnur_reporte_tbl")
  expect_true(all(c("p5.1", "p5.2", "p5.3") %in% names(rp_data)))
  expect_false("p5" %in% names(rp_data))
  expect_false(is.null(attr(rp_data$p1, "labels")))

  expect_no_error(
    prosecnur::reporte_frecuencias(
      data = rp_data,
      instrumento = rp_inst,
      secciones = list(
        Principal = c("p1", "p4_1", "p5")
      ),
      path_xlsx = path_freq
    )
  )
  expect_true(file.exists(path_freq))

  expect_no_error(
    prosecnur::reporte_cruces(
      data = rp_data,
      instrumento = rp_inst,
      SECCIONES = list(
        Principal = c("p4_1", "p4_2")
      ),
      cruces = c("p1"),
      path_xlsx = path_cross
    )
  )
  expect_true(file.exists(path_cross))

  list_name_p4 <- rp_inst$survey$list_name[idx_rp_p4_1]
  orden_p4 <- as.character(rp_inst$choices$name[rp_inst$choices$list_name == list_name_p4])

  recod <- expect_no_error(
    prosecnur::reporte_recodificar_items(
      data = rp_data,
      instrumento = rp_inst,
      vars = c("p4_1"),
      orden_por_lista = stats::setNames(list(orden_p4), list_name_p4)
    )
  )
  expect_true("r100_p4_1" %in% names(recod))
})

test_that("surveymonkey_xlsform reconoce satisfaccion_4 y ordena grupos por sufijo", {
  path_sav <- tempfile(fileext = ".sav")
  path_xlsx <- tempfile(fileext = ".xlsx")
  on.exit(unlink(c(path_sav, path_xlsx)), add = TRUE)

  sat_labs <- c(
    "Muy insatisfecho" = 1,
    "Insatisfecho" = 2,
    "Satisfecho" = 3,
    "Muy satisfecho" = 4,
    "SIN INF" = 99
  )

  df <- data.frame(
    P8_1 = haven::labelled(c(1, 2), sat_labs),
    P8_2 = haven::labelled(c(2, 3), sat_labs),
    P8_6 = haven::labelled(c(3, 4), sat_labs),
    P8_3 = haven::labelled(c(4, 1), sat_labs),
    P8_4 = haven::labelled(c(1, 99), sat_labs),
    P8_5 = haven::labelled(c(2, 3), sat_labs),
    stringsAsFactors = FALSE
  )

  attr(df$P8_1, "label") <- "Servicio 1"
  attr(df$P8_2, "label") <- "Servicio 2"
  attr(df$P8_6, "label") <- "Servicio 6"
  attr(df$P8_3, "label") <- "Servicio 3"
  attr(df$P8_4, "label") <- "Servicio 4"
  attr(df$P8_5, "label") <- "Servicio 5"

  haven::write_sav(df, path_sav)

  sm <- prosecnur::surveymonkey_leer(path_sav)
  inst_ref <- prosecnur::surveymonkey_xlsform(sm, path = path_xlsx)

  p8_rows <- inst_ref$survey[grepl("^p8_", inst_ref$survey$name), , drop = FALSE]
  expect_identical(p8_rows$name, paste0("p8_", 1:6))
  expect_true(all(p8_rows$type == "select_one lst_satisfaccion_4"))
  expect_true("lst_satisfaccion_4" %in% inst_ref$choices$list_name)
})
