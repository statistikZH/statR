
# statR

The **statR** package is a toolbox for corporate design conforming graphics and spreadsheets of the statistical office of the canton of Zurich . It contains:

- *generic colour palettes for any kind of data-visualization*

- *ggplot2-theme*

- *templates for saving data in CD-conforming formated xlsx-spreadsheets*

For examples and instructions please consult the *NEW* pkgdown-page (under development):

https://statistikzh.github.io/statR/

Das statR-package enthält:

- Das Stat ZH Theme (stat_theme()) für ggplot2 Grafiken
- Farbpaletten (zhpal)
- Farbpaletten-Funktionen für ggplot2
- Funktionen um Datensätze aus R in formattierte und Kantons CD-konforme XLSX-files abzuspeichern.

Zu jeder Funktion ist mittels Hilfe-Befehl (?) eine Beschreibung sowie ein Verwendungsbeispiel aufrufbar (z.B.: ?splitXLSX )

# Installation

Die Development Version des Package liegt auf github und kann auf zwei Varianten installiert werden. 

```
# Variante 1
library(downloader)
download("https://github.com/statistikZH/statR/archive/dev.tar.gz", "statR.tar.gz")
install.packages("statR.tar.gz", repos = NULL, type = "source")

# Variante 2
library(devtools)

devtools::install_github("statistikZH/statR",ref="dev")
```






