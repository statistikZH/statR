# statR 

Mit dem `statR`-Package können mit dem Corporate Design des Kantons Zürich konforme Visualisierungen und Excel-Tabellen in `R` erstellt werden. Das Package beinhaltet:


- Funktionen, mit denen Datensätze aus R als XLSX-Datei abgespeichert werden können (inkl. Quellenangaben und weiteren Zusatzinformationen).

- generische Farbpaletten für Datenvisualisierungen

- ein `ggplot2` theme

- ein Template für Html-Berichte


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






