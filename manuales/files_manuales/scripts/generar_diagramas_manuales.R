args <- commandArgs(trailingOnly = FALSE)
file_arg <- grep("^--file=", args, value = TRUE)[1]

if (is.na(file_arg)) {
  stop("No se pudo inferir la ruta del wrapper de diagramas.", call. = FALSE)
}

wrapper_path <- normalizePath(sub("^--file=", "", file_arg), winslash = "/", mustWork = TRUE)
repo_root <- normalizePath(file.path(dirname(wrapper_path), "..", "..", "..", "..", ".."), winslash = "/", mustWork = TRUE)

source(file.path(repo_root, "scripts", "generar_diagramas_manuales_qmd.R"), local = TRUE)
generate_diagramas_manuales_qmd(repo_root = repo_root)
