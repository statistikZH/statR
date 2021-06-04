# statR 2.0.0

* An R Markdown template `ZH report` for .html-reports is now available.  
* New function `datasetsXLSX()`: several datasets and/or figures can now be exported from R to a single .xlsx-file. The function creates an overview sheet and separate sheets for each dataset/figure.
* Additional function arguments for `quickXLSX()`, `splitXLSX()`, and `datasetsXLSX()` have been implemented, for example the author's initials and individual contact details can be specified and the logo of the canton of Zurich can be added. 
* The column widths in the output of `quickXLSX()`, `splitXLSX()`, and `datasetsXLSX()` are set automatically and the first row no longer freezes. 
* The documentation has been extended.
* `theme_stat()` has been revised and a couple of theme arguments have been added for increased flexibility (e.g., axis ticks, position of axis labels)
