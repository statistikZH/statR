# statR 

  <!-- badges: start -->
  [![R-CMD-check](https://github.com/statistikZH/statR/workflows/R-CMD-check/badge.svg)](https://github.com/statistikZH/statR/actions)
  <!-- badges: end -->

<p align="left">
<img src="https://raw.githubusercontent.com/statistikZH/statR/dev/inst/extdata/Stempel_STAT-01.png" alt="" width="220"/>

Mit dem `statR`-Package können mit dem Corporate Design des Kantons Zürich konforme Visualisierungen, Excel-Tabellen und Html-Berichte in `R` erstellt werden. Das Package beinhaltet:


- Funktionen, mit denen Datensätze aus R als XLSX-Datei abgespeichert werden können (inkl. Quellenangaben und weiteren Zusatzinformationen).

- generische Farbpaletten für Datenvisualisierungen

- ein `ggplot2` theme

- ein Template für Html-Berichte


# Installation

Das Package liegt auf GitHub. Es kann auf zwei Arten installiert werden:

```
# Variante 1

library(devtools)
devtools::install_github("statistikZH/statR")

# Variante 2

library(downloader)
download("https://github.com/statistikZH/statR/archive/refs/heads/master.tar.gz", "statR.tar.gz")
install.packages("statR.tar.gz", repos = NULL, type = "source")


```

Die Development-Version kann folgendermassen installiert werden:

```
# Variante 1

library(devtools)
devtools::install_github("statistikZH/statR", ref = "dev")

# Variante 2

library(downloader)
download("https://github.com/statistikZH/statR/archive/refs/heads/dev.tar.gz", "statR.tar.gz")
install.packages("statR.tar.gz", repos = NULL, type = "source")

```

