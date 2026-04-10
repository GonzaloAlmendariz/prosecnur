#set text(font: "Arial", size: 10pt, lang: "es")
#set par(justify: false, leading: 0.75em)
#set page(numbering: "1", number-align: right)

#show heading.where(level: 1): it => {
  v(0.9em)
  block(
    width: 100%,
    fill: rgb("#1A4A7A"),
    stroke: (paint: rgb("#163C63"), thickness: 0.8pt),
    inset: (x: 12pt, y: 8pt),
    radius: 6pt
  )[
    #text(fill: white, weight: "bold", size: 13pt)[#it.body]
  ]
  v(0.5em)
}

#show heading.where(level: 2): it => {
  v(0.7em)
  block(
    width: 100%,
    fill: rgb("#2C5A97"),
    stroke: (paint: rgb("#23497A"), thickness: 0.8pt),
    inset: (x: 10pt, y: 6pt),
    radius: 5pt
  )[
    #text(fill: white, weight: "bold", size: 11.2pt)[#it.body]
  ]
  v(0.3em)
}

#show heading.where(level: 3): it => {
  v(0.5em)
  text(weight: "bold", fill: rgb("#1A4A7A"), size: 10.5pt)[#it.body]
  v(0.15em)
}

#show strong: it => text(weight: "bold", fill: rgb("#1A4A7A"))[#it.body]
