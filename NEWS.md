# statR 2.1.2

* implemented new color palettes
* removes old and not accessible palettes
* the documentation has been update 
  * bugs in the documentation fixed

# statR 2.1.1

* fix `datasetXLSX()` which generated broken Links in the index-sheet

# statR 2.1.0

* New function `aXLSX()`: exports data from R to a formatted excel file optimized for accessbility, separating contact info and additional metadata (other than the source and title) from the data
* Updated `theme_stat()`: `theme_stat(map = TRUE)` optimizes the theme for maps
* Updated `theme_stat()`: additional function arguments to add or remove major and minor grid lines
* Extended documentation, (e.g., added `flush_left()` to the vignette on visualizations, a function that allows one to align the title, subtitle, and caption with the y-axis of a plot)
* New function `insert_metadata_sheet()`: adds formatted metadata to an existing .xlsx-workbook

# statR 2.0.0

* An R Markdown template `ZH report` for .html-reports is now available.  
* New function `datasetsXLSX()`: several datasets and/or figures can now be exported from R to a single .xlsx-file. The function creates an overview sheet and separate sheets for each dataset/figure.
* Additional function arguments for `quickXLSX()`, `splitXLSX()`, and `datasetsXLSX()` have been implemented, for example the author's initials and individual contact details can be specified and the logo of the canton of Zurich can be added. 
* The column widths in the output of `quickXLSX()`, `splitXLSX()`, and `datasetsXLSX()` are set automatically and the first row no longer freezes. 
* The documentation has been extended.
* `theme_stat()` has been revised and a couple of theme arguments have been added for increased flexibility (e.g., axis ticks, position of axis labels)
