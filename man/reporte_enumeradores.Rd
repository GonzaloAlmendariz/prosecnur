\name{reporte_enumeradores}
\alias{reporte_enumeradores}
\title{Generar reporte PDF de enumeradores por modalidad y cortes de cuota}
\usage{
reporte_enumeradores(
  data,
  col_enumerador,
  cols_corte = NULL,
  col_modalidad = NULL,
  modalidad_reglas = NULL,
  modalidad_fn = NULL,
  modalidad_default = "Presencial",
  modalidades_esperadas = c("Presencial", "Telefónica"),
  mostrar_modalidades_vacias = FALSE,
  output_file = "reporte_enumeradores.pdf",
  titulo = "Producción de Enumeradores",
  subtitulo = NULL,
  min_encuestas = 0,
  ordenar_por = c("total", "nombre"),
  quiet = FALSE
)
}
\arguments{
\item{data}{\code{data.frame} o \code{tibble} con los datos crudos.}

\item{col_enumerador}{Nombre de la columna de enumerador.}

\item{cols_corte}{Vector opcional de columnas para desagregar el reporte.}

\item{col_modalidad}{Columna existente de modalidad (opcional).}

\item{modalidad_reglas}{Reglas opcionales para clasificar modalidad. Puede ser
lista de reglas o \code{data.frame} con columna \code{modalidad}.}

\item{modalidad_fn}{Función opcional \code{f(data) -> character(nrow(data))} para
clasificar modalidad.}

\item{modalidad_default}{Etiqueta por defecto para filas sin modalidad.}

\item{modalidades_esperadas}{Modalidades esperadas para ordenar secciones.}

\item{mostrar_modalidades_vacias}{Si \code{TRUE}, incluye secciones vacías para
\code{modalidades_esperadas}.}

\item{output_file}{Ruta de salida del PDF.}

\item{titulo}{Título del reporte.}

\item{subtitulo}{Subtítulo opcional.}

\item{min_encuestas}{Filtro mínimo de producción por enumerador.}

\item{ordenar_por}{Criterio de orden para filas de enumerador.}

\item{quiet}{Si \code{TRUE}, reduce salida de \code{quarto_render()}.}
}
\value{
Invisiblemente una lista con \code{output_file} (archivo general),
\code{output_files} (todos los PDFs generados), \code{bundle} y \code{bundles}.
}
\description{
\code{reporte_enumeradores()} construye un reporte en PDF (Quarto + Typst) con
producción observada por enumerador. El desglose es genérico: se puede pasar
cualquier conjunto de columnas de corte (por ejemplo \code{sexo},
\code{aula}, \code{turno}) mediante \code{cols_corte}.
}
\details{
La modalidad se resuelve con precedencia fija:
\code{modalidad_fn > col_modalidad > modalidad_reglas > modalidad_default}.
}
