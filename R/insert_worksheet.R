#' insert_worksheet()
#'
#' @description Function to  add formatted worksheets to an existing
#'  .xlsx-workbook.
#'
#' @note Marked for deprecation in upcoming version. The function does not write
#'  the result into a .xlsx file. A separate call to openxlsx::saveWorkbook() is
#'  required.
#'
#' @param data data to be included.
#'
#' @param workbook workbook object to add new worksheet to.
#'
#' @param title title to be put above the data.
#'
#' @param sheetname name of the sheet tab.
#'
#' @param source source of the data. Defaults to "statzh".
#'
#' @param metadata metadata information to be included. Defaults to NA.
#'
#' @param logo path of the file to be included as logo (png / jpeg / svg).
#'  Defaults to "statzh"
#'
#' @param contactdetails contact details of the data publisher. Defaults to
#'  "statzh".
#'
#' @param grouplines defaults to FALSE. Can be used to separate grouped
#'  variables visually.
#'
#' @param author defaults to the last two letters (initials) or numbers of the
#'  internal user name.
#'
#' @importFrom dplyr "%>%"
#' @keywords insert_worksheet
#' @export
#' @examples
#'
#' # Generation of a spreadsheet
#' wb <- openxlsx::createWorkbook("hello")
#'
#' insert_worksheet(data = head(mtcars), workbook = wb, title = "mtcars",
#'  sheetname = "carb")
#'
insert_worksheet <- function(data, workbook, sheetname = "data", title = "Title",
                             source = "statzh", metadata = NA, logo = "statzh",
                             grouplines = FALSE, contactdetails = "statzh",
                             author = "user"){

  # Check length of sheetname
  sheetname <- check_sheetname(sheetname)

  # Add worksheet
  openxlsx::addWorksheet(workbook, sheetname)


  # Process input ----------

  ## Contactdetails & position
  contactdetails <- prep_contact(contactdetails, compact = TRUE)
  contact_pos <- max(ncol(data) - 2, 4)

  ## Source information
  source <- prep_source(source)

  ## Metadata

  if (any(is.na(metadata))) {
    remarks <- "Bemerkungen:"
  } else if (any(metadata == "HAE")) {
    remarks <- "Die Zahlen der letzten drei Jahre sind provisorisch."
  } else {
    remarks <- metadata
  }

  #--------
  # Insert content

  ## Logo
  if (!is.null(logo)){
    logo <- prep_logo(logo)

    if (file.exists(logo)){
      openxlsx::insertImage(workbook, sheet = "Inhalt", file = logo, width = 2.145,
                            height = 0.7865, units = "in")
    } else {
      message("no logo found.")
    }
  }

  ## Titel
  openxlsx::writeData(workbook, sheet = sheetname, x = title, startRow = 7,
                      headerStyle = style_title())

  n_metadata <- length(metadata)
  datenbereich <- 9 + n_metadata + 3
  openxlsx::writeData(workbook, sheet = sheetname, x = metadata, startRow = 8,
                      headerStyle = style_subtitle3())

  ### Quelle
  openxlsx::writeData(workbook, sheet = sheetname, x = source, startRow = 8 + n_metadata,
                      headerStyle = style_subtitle3())

  # Metadaten zusammenmergen
  purrr::walk(7:(7 + length(metadata) + length(source)),
              ~openxlsx::mergeCells(wb, sheet = sheetname, cols = 1:26, rows = .))

  ### Contact information
  purrr::walk(2:5, ~openxlsx::mergeCells(wb, sheet = sheetname,
                                         cols = contact_pos:26, rows = .))

  openxlsx::writeData(wb, sheet = sheetname, x = contactdetails,
                      startCol = contact_pos, startRow = 2, headerStyle = style_wrap())

  ### Aktualisierungsdatum

  # User-K\u00fcrzel f\u00fcr Kontaktinformationen
  if(author == "user"){
    # f\u00fcr das lokale R
    if(Sys.getenv("USERNAME")!="") {
      contactperson <- stringr::str_sub(Sys.getenv("USERNAME"), start = 6, end = 7)
    } else {
      # f\u00fcr den R-server
      contactperson <- stringr::str_sub(Sys.getenv("USER"), start = 6, end = 7)
    }
  } else {
    contactperson <- author
  }

  openxlsx::writeData(wb, sheet = sheetname,
                      x = paste(prep_creationdate("Aktualisiert am"),
                                           "durch:", contactperson),
                      startCol = contact_pos, startRow = 5,
                      headerStyle = style_subtitle3())

  # Insert data --------

  ## increase width of colnames for better auto-fitting of column width
  colnames(data) <- paste0(colnames(data), "  ", sep = "")

  openxlsx::writeData(wb, sheet = sheetname,
                      x = as.data.frame(data %>% dplyr::ungroup()),
                      startRow = datenbereich, rowNames = FALSE,
                      withFilter = FALSE)

  #-----------------------
  # Apply styles

  ## Title
  openxlsx::addStyle(wb, sheet = sheetname, style = style_title(),
                     rows = 7, cols = 1, gridExpand = TRUE)

  ## Data header
  openxlsx::addStyle(wb, sheet = sheetname, style = style_header(),
                     rows = datenbereich, cols = 1:ncol(data),
                     gridExpand = TRUE, stack = TRUE)

  ## Headerline
  openxlsx::addStyle(wb, sheet = sheetname, style = style_headerline(),
                     rows = 5, cols = 1:ncol(data), gridExpand = TRUE,
                     stack = TRUE)

  ## Grouplines
  if (!is.null(grouplines)){
    datenbereich_end <- nrow(data) + datenbereich

    openxlsx::addStyle(wb, sheet = sheetname, style = style_leftline(),
                       rows = datenbereich:datenbereich_end,
                       cols = grouplines, gridExpand = TRUE, stack = TRUE)
  }

  ## minmale Spaltenbreite definieren
  options("openxlsx.minWidth" = 5)

  ## automatische Zellenspalten
  openxlsx::setColWidths(wb, sheet = sheetname, cols = 1:ncol(data),
                         widths = "auto", ignoreMergedCells = TRUE)
}
