#' insert_worksheet()
#'
#' Function to  add formatted worksheets to an existing .xlsx-workbook.
#' @note
#' Marked for deprecation in upcoming version. The function does not write the
#' result into a .xlsx file. A separate call to openxlsx::saveWorkbook() is
#' required.
#' @param data data to be included.
#' @param workbook workbook object to add new worksheet to.
#' @param title title to be put above the data.
#' @param sheetname name of the sheet tab.
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata information to be included. Defaults to NA.
#' @param logo path of the file to be included as logo (png / jpeg / svg). Defaults to "statzh"
#' @param contactdetails contact details of the data publisher. Defaults to "statzh".
#' @param grouplines defaults to FALSE. Can be used to separate grouped variables visually.
#' @param author defaults to the last two letters (initials) or numbers of the internal user name.
#' @importFrom dplyr "%>%"
#' @keywords insert_worksheet
#' @export
#' @examples
#'
#' # Generation of a spreadsheet
#' wb <- openxlsx::createWorkbook("hello")
#'
#' insert_worksheet(data = mtcars[c(1:10),], workbook = wb, title = "mtcars", sheetname = "carb")


insert_worksheet <- function(data, workbook, sheetname="data",title="Title",
                             source="statzh", metadata = NA, logo= "statzh",
                             grouplines = FALSE, contactdetails="statzh",
                             author = "user"){

  # Metadata
  remarks <- if (any(is.na(metadata))) {
    "Bemerkungen:"
  }
  else if (any(metadata == "HAE")) {
    "Die Zahlen der letzten drei Jahre sind provisorisch."
  }
  else {
    metadata
  }


  #Zahlenformat: Tausendertrennzeichen
  # options("openxlsx.numFmt" = "#,###0")

  #extrahiere colname
  # col_name <- rlang::enquo(sheetvar)

  wb<-workbook

  # data-container from row 5
  n_metadata <- length(metadata)

  datenbereich = 9 + n_metadata + 3

  #define width of the area in which data is contained for formatting
  spalten = ncol(data)

  # increase width of colnames for better auto-fitting of column width
  colnames(data) <- paste0(colnames(data), "  ", sep = "")

  #position of contact details
  contact = if(spalten>=6){
    spalten-2
  } else {
    4
  }

  #styles
  titleStyle <- openxlsx::createStyle(fontSize=14, textDecoration="bold",fontName="Arial")

  subtitle <- openxlsx::createStyle(fontSize=12, textDecoration="italic",fontName="Arial")

  header <- openxlsx::createStyle(fontSize = 12, fontColour = "#000000",  halign = "left", border="Bottom",  borderColour = "#009ee0",textDecoration = "bold")

  # header1 <- createStyle(fontSize = 14, fontColour = "#FFFFFF", halign = "center",
  #                        fgFill = "#3F98CC", border="TopBottom", borderColour = "#4F81BD")

  # header2 <- createStyle(fontSize = 12, fontColour = "#FFFFFF", halign = "center",
  #                        fgFill = "#407B9F", border="TopBottom", borderColour = "#4F81BD")

  headerline <- openxlsx::createStyle(border="Bottom", borderColour = "#009ee0",borderStyle = getOption("openxlsx.borderStyle", "thick"))

  #Linien
  bottomline <- openxlsx::createStyle(border="Bottom", borderColour = "#009ee0")

  leftline <- openxlsx::createStyle(border="Left", borderColour = "#4F81BD")

  #wrap
  wrap <- openxlsx::createStyle(wrapText = TRUE)


  ### Loop for multiple years / worksheets ------------

  # for (year in sheets){
  #
  #   #get index
  #   i<- which(sheets==year)

  # warning if sheetname is longer than the limit imposed by excel (31 characters)
  if(nchar(sheetname)>31){warning("sheetname is cut to 31 characters (limit imposed by MS-Excel)")}

  ## Add worksheet
  # openxlsx::addWorksheet(wb,paste(substr(sheetname,0,31)))

  suppressWarnings(openxlsx::addWorksheet(wb,paste(substr(sheetname,0,31))))

  i <- paste(substr(sheetname,0,31))


  # Style ---------------------

  #title,subtitle & Header


  # Titel & Untertitel -----------------

  #Logo

  if(!is.null(logo)){

    if(logo == "statzh") {

      statzh <- paste0(.libPaths(),"/statR/extdata/Stempel_STAT-01.png")
      openxlsx::insertImage(wb, i, statzh[1], width = 2.145, height = 0.7865,
                            units = "in")
    } else if(logo == "zh"){

      zh <- paste0(.libPaths(),"/statR/extdata/Stempel_Kanton_ZH.png")

      openxlsx::insertImage(wb, i, zh[1], width = 2.145, height = 0.7865,
                            units = "in")
    } else if((logo != "statzh" | logo != "zh") & file.exists(logo)) {

      openxlsx::insertImage(wb, i, logo, width = 2.145, height = 0.7865,
                            units = "in")
    } else {

      message("no logo found and / or added")
    }
  }




  #standard contactdetails
  if(contactdetails=="statzh"){

    contactdetails <- c("Datashop, Tel: 0432597500",
                        "datashop@statistik.ji.zh.ch",
                        "http://www.statistik.zh.ch")

  }else {contactdetails}

  if (length(contactdetails) > 3) {
    warning("Contactdetails may overlap with other elements. To avoid this issue please do not include more than three elements in the contactdetails vector.")
  }


  if(source=="statzh"){

    source="Quelle: Statistisches Amt des Kantons Z\u00fcrich"

  }else {source}



  #Titel
  openxlsx::writeData(wb, sheet = i,title, headerStyle=titleStyle,startRow = 7)


  ##Metadata
  openxlsx::writeData(wb, sheet = i, metadata, headerStyle=subtitle, startRow = 8)


  ##Quelle
  openxlsx::writeData(wb, sheet = i, source, headerStyle=subtitle, startRow = 8+n_metadata)

  # Metadaten zusammenmergen
  purrr::walk(7:(7+length(metadata)+length(source)), ~openxlsx::mergeCells(wb, sheet = i, cols = 1:26, rows = .))
  # Kontaktdaten zusammenmergen
  purrr::walk(2:5, ~openxlsx::mergeCells(wb, sheet = i, cols = contact:26, rows = .))


  #Kontakt
  openxlsx::writeData(wb, sheet = i,
                      contactdetails,
                      headerStyle = wrap,
                      startRow = 2,
                      startCol = contact)

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

  #Aktualisierungsdatum
  openxlsx::writeData(wb, sheet = i, paste("Aktualisiert am ",
                                           format(Sys.Date(), format="%d.%m.%Y"),
                                           " durch: ", contactperson),
                      headerStyle=subtitle, startRow = 5, startCol=contact)

  # Daten abf\u00fcllen
  openxlsx::writeData(wb, sheet = i, as.data.frame(data%>%dplyr::ungroup()), rowNames = FALSE, startRow = datenbereich, withFilter = FALSE)

  #F\u00fcge Formatierungen ein

  openxlsx::addStyle(wb, sheet = i, headerline, rows = 5, cols = 1:spalten, gridExpand = TRUE,stack = TRUE)

  openxlsx::addStyle(wb, sheet = i, titleStyle, rows = 7, cols = 1, gridExpand = TRUE)

  # addStyle(wb, sheet = i, header1, rows = datenbereich, cols = 1:spalten, gridExpand = TRUE,stack = TRUE)

  openxlsx::addStyle(wb, sheet = i, header, rows = datenbereich, cols = 1:spalten, gridExpand = TRUE,stack = TRUE)

  if (!is.null(grouplines)){

    datenbereich_end <-nrow(data)+datenbereich

    openxlsx::addStyle(wb, sheet = i, leftline, rows=datenbereich:datenbereich_end, cols = grouplines, gridExpand = TRUE,stack = TRUE)

  }


  # bodyStyle <- createStyle(border="TopBottom", borderColour = "#4F81BD")
  # addStyle(wb, sheet = 1, bodyStyle, rows = 2:6, cols = 1:11, gridExpand = TRUE)

  # minmale Spaltenbreite definieren
  options("openxlsx.minWidth" = 5)

  # automatische Zellenspalten
  openxlsx::setColWidths(wb, sheet = i, cols=1:spalten, widths = "auto", ignoreMergedCells = TRUE) ## set column width for row names column

  # newworkbook<<-wb

}
