
# statR

Mit dem `statR`-Package können Visualisierungen und Excel-Tabellen in `R` erstellt werden, die mit dem Corporate Design des Kantons Zürich konform sind. Das Package beinhaltet:


- Funktionen, mit denen Datensätze aus R als XLSX-Datei abgespeichert werden können inklusive Quellenangaben und weitere Zusatzinformationen.

- generische Farbpaletten für Datenvisualisierungen

- `ggplot2` themes



Weitere Informationen zum Package und Beispiele sind auf der pkgdown-Seite zu finden:  https://statistikzh.github.io/statR/.


# Installation

Die development-Version des Package liegt auf GitHub. Sie kann auf zwei Arten installiert werden:

```
# Variante 1

library(downloader)
download("https://github.com/statistikZH/statR/archive/dev.tar.gz", "statR.tar.gz")
install.packages("statR.tar.gz", repos = NULL, type = "source")


# Variante 2

library(devtools)
devtools::install_github("statistikZH/statR",ref="dev")
```






