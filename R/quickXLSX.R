# quickXLSX: Function to create formatted spreadsheet with a single worksheet automatically

#' quickXLSX
#'
#' Function to create a formated single-worksheet XLSX automatically
#' @param data data to be included in the XLSX-table.
#' @param name title of the table / file.
#' @metadata metadata-information to be included.
#' @logo path of the file to be included as logo (png / jpeg / svg)
#' @param grouplines insert vertical lines to separate columns visually.
#' @keywords quickXLSX
#' @export
#' @examples
#' quickXLSX(data, name, filename="file", grouplines = 0, metadata = NA, logo = statzh,
#' contactdetails = statzhcontact)

# Function


quickXLSX <-function (data, name,filename="file", grouplines = 0, metadata = NA, logo = NA,
                      contactdetails = statzhcontact) {

  remarks <- if (is.na(metadata)) {
    "Bemerkungen:"
  }
  else if (metadata == "HAE") {
    "Die Zahlen der letzten drei Jahre sind provisorisch."
  }
  else {
    metadata
  }

  wb <- openxlsx::createWorkbook(paste(name))
  options(openxlsx.numFmt = "#,###0")
  datenbereich = 14
  spalten = ncol(data)
  contact = if (spalten > 4) {
    spalten - 2
  }
  else {
    3
  }
  title <- openxlsx::createStyle(fontSize = 14, textDecoration = "bold",
                                 fontName = "Arial")
  subtitle <- openxlsx::createStyle(fontSize = 12, textDecoration = "italic",
                                    fontName = "Arial")
  header <- openxlsx::createStyle(fontSize = 12, fontColour = "#000000",
                                  halign = "left", border = "Bottom", borderColour = "#009ee0",
                                  textDecoration = "bold")
  headerline <- openxlsx::createStyle(border = "Bottom", borderColour = "#009ee0",
                                      borderStyle = getOption("openxlsx.borderStyle", "thick"))
  bottomline <- openxlsx::createStyle(border = "Bottom", borderColour = "#009ee0")
  leftline <- openxlsx::createStyle(border = "Left", borderColour = "#4F81BD")
  wrap <- openxlsx::createStyle(wrapText = TRUE)
  openxlsx::addWorksheet(wb, "data")
  statzh <- "Stempel_STAT-01.png"

  if (is.na(logo)) { logo <- statzh }

  if (file.exists(logo)) {

    openxlsx::insertImage(wb, 1, logo, width = 2.145, height = 0.7865,
                          units = "in")
  }

  openxlsx::writeData(wb, sheet = 1, substr(name,0,6) , headerStyle = title,
                      startRow = 7)
  openxlsx::writeData(wb, sheet = 1, "Quelle: Statistisches Amt des Kantons Zürich",
                      headerStyle = subtitle, startRow = 8)
  openxlsx::writeData(wb, sheet = 1, remarks, headerStyle = subtitle,
                      startRow = 9)
  statzhcontact <- c("Datashop, Tel: 0432597500", "datashop@statistik.ji.zh.ch",
                     "http://www.statistik.zh.ch")
  if (length(contactdetails) > 3) {
    warning("Contactdetails may overlap with other elements. To avoid this issue please do not include more than three elements in the contactdetails vector.")
  }
  openxlsx::writeData(wb, sheet = 1, contactdetails, headerStyle = wrap,
                      startRow = 2, startCol = contact)
  openxlsx::writeData(wb, sheet = 1, paste("Aktualisiert am ",
                                           format(Sys.Date(), format = "%d.%m.%Y"), " durch: ",
                                           stringr::str_sub(Sys.getenv("USERNAME"), -2)), headerStyle = subtitle,
                      startRow = 5, startCol = contact)
  openxlsx::writeData(wb, sheet = 1, data, rowNames = FALSE,
                      startRow = datenbereich, withFilter = FALSE)
  openxlsx::addStyle(wb, sheet = 1, title, rows = 7, cols = 1,
                     gridExpand = TRUE)
  openxlsx::addStyle(wb, sheet = 1, headerline, rows = 5, cols = 1:spalten,
                     gridExpand = TRUE, stack = TRUE)
  openxlsx::addStyle(wb, sheet = 1, header, rows = datenbereich,
                     cols = 1:spalten, gridExpand = TRUE, stack = TRUE)
  if (grouplines != 0) {
    openxlsx::addStyle(wb, sheet = 1, leftline, rows = (datenbereich -
                                                          1):nrow(data), cols = grouplines, gridExpand = TRUE,
                       stack = TRUE)
  }
  openxlsx::freezePane(wb, sheet = 1, firstActiveRow = datenbereich +2)
  openxlsx::setColWidths(wb, 1, cols = 1:spalten, widths = 20)
  openxlsx::saveWorkbook(wb, paste(substr(name,0,20), ".xlsx", sep = ""),
                         overwrite = TRUE)
}


