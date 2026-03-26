make_plan_boxplot_fixture <- function() {
  df <- data.frame(
    score = c(2.3, 3.4, 2.9, 4.1, 3.8, 2.7, 3.1, 4.2),
    region = c("N", "N", "S", "S", "N", "S", "N", "S"),
    stringsAsFactors = FALSE
  )

  attr(df$score, "label") <- "Puntaje de satisfaccion"
  attr(df$region, "label") <- "Region"

  survey <- data.frame(
    name = c("score", "region"),
    type = c("decimal", "select_one lst_region"),
    list_name = c(NA_character_, "lst_region"),
    stringsAsFactors = FALSE
  )

  choices <- data.frame(
    list_name = rep("lst_region", 2),
    name = c("N", "S"),
    label = c("Norte", "Sur"),
    stringsAsFactors = FALSE
  )

  list(
    data = df,
    instrumento = list(survey = survey, choices = choices, orders_list = NULL),
    presets = p_presets(
      boxplot = list(
        usar_canvas = FALSE,
        mostrar_puntos = FALSE,
        mostrar_leyenda = TRUE
      )
    )
  )
}

test_that("p_boxplot valida argumentos basicos", {
  expect_error(p_boxplot(var = ""), "`var`")
  expect_error(p_boxplot(var = "score", cruce = ""), "`cruce`")
  expect_error(p_boxplot(var = "score", overrides = "x"), "`overrides`")
  expect_error(p_boxplot(var = "score", base = "x"), "`base`")

  el <- p_boxplot(var = "score", cruce = "region")
  expect_identical(el$.element_type, "boxplot")
  expect_identical(el$var, "score")
  expect_identical(el$cruce, "region")
})

test_that("reporte_ppt_plan renderiza boxplot simple", {
  skip_if_not_installed("officer")
  skip_if_not_installed("rvg")

  fx <- make_plan_boxplot_fixture()
  plan <- list(
    diapo_001 = p_slide_1(
      plot = p_boxplot("score")
    )
  )

  out <- reporte_ppt_plan(
    data = fx$data,
    instrumento = fx$instrumento,
    plan = plan,
    presets = fx$presets,
    solo_lista = TRUE,
    mensajes_progreso = FALSE
  )

  expect_length(out$rendered, 1L)
  expect_s3_class(out$rendered[[1]], "ggplot")
  expect_true(any(vapply(
    out$rendered[[1]]$layers,
    function(l) inherits(l$geom, "GeomBoxplot"),
    logical(1)
  )))
})

test_that("reporte_ppt_plan boxplot aplica etiquetas de cruce y render_meta", {
  skip_if_not_installed("officer")
  skip_if_not_installed("rvg")

  fx <- make_plan_boxplot_fixture()
  plan <- list(
    diapo_001 = p_slide_1(
      plot = p_boxplot("score", cruce = "region")
    )
  )

  out <- reporte_ppt_plan(
    data = fx$data,
    instrumento = fx$instrumento,
    plan = plan,
    presets = fx$presets,
    solo_lista = TRUE,
    mensajes_progreso = FALSE,
    build_render_meta = TRUE
  )

  cats <- as.character(out$rendered[[1]]$data$categoria)
  expect_true(all(c("Norte", "Sur") %in% cats))
  expect_equal(out$render_meta[[1]]$etype, "boxplot")
  expect_equal(out$render_meta[[1]]$kind, "chart")
})
