# statR 2.4.0 

* Fixed an issue where Excel Worksheets created using datasetsXLSX could become corrupted.
* When defining a second header (either by numeric index or variable name at the beginning of a block) 
  are now properly merged and centered. Last block terminates at the final column of the dataset.
* splitXLSX is now based on datasetsXLSX.
* When including images with datasetsXLSX, users can now specify a title, source, and metadata.
* Users can now attach an additional metadata sheet in datasetsXLSX.
* insert_worksheet allows users to leave the source and metadata arguments at NULL. In these instances,
  the function no longer includes an empty row.
* Reworked datasetsXLSX to allow a tidyverse-like workflow where titles, sources, metadata, etc. are
  attached as attributes to objects (see examples).
* As part of this syntax, users can more easily control how sources and metadata are displayed. add_source()
  and add_metadata() both take prefix and collapse as input.
* insert_hyperlinks now (formally) allows users to point to external files.
* 


# statR 2.3.2

## Bugfix
* Fixed a problem where user configurations were not being applied properly

# statR 2.3.1

## Bugfixes
* Fixed a problem with user configurations on Windows systems

# statR 2.3.0

## New Features
* Users can now create, load, and switch between configurations, allowing for more personalized output files.
* grouplines and second header functionality is now available in all functions which generate worksheets.
* base::histogram objects can now be passed to datasetsXLSX
* Prefixes for fields like auftrags_id are now customizable
* Positioning of content in worksheets is now relative to last inserted content

## Bugfixes
* Fixed undesired behavior where inserting an image by path resulted in the source file being deleted

# statR 2.2.0

* implemented new color palettes according to the ZH.CH Design System 
* outdated not accessible palettes have been discarded ("zhextralight" and "zhultralight" among others)
* errors in the documentation have been corrected

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
