#' insert_metadata_sheet()
#'
#' Function to  add formatted worksheets to an existing .xlsx-workbook.
#' @param wb workbook object to add new worksheet to.
#' @param title title to be put above the data.
#' @param sheetname name of the sheet tab.
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata information to be included. Defaults to NA.
#' @param logo path of the file to be included as logo (png / jpeg / svg). Defaults to "statzh"
#' @param contactdetails contact details of the data publisher. Defaults to "statzh".
#' @param author defaults to the last two letters (initials) or numbers of the internal user name.
#' @importFrom dplyr "%>%"
#' @keywords insert_metadata_sheet
#' @export
#' @examples
#'
#' # Generation of a spreadsheet
#' wb <- openxlsx::createWorkbook("hello")
#'
#' insert_metadata_sheet(workbook = wb, title = "mtcars", sheetname = "Metadaten", source = "statzh")


insert_metadata_sheet <- function(wb, sheetname="Metadaten",title="Title",
                             source="statzh", metadata = NA, logo= "statzh",
                             contactdetails="statzh", author = "user") {

  n_metadata <- length(metadata)

  # warning if sheetname is longer than the limit imposed by excel (31 characters)
  if(nchar(sheetname)>31){
    warning("sheetname is cut to 31 characters (limit imposed by MS-Excel)")
    }

  ## Add metadata sheet

  suppressWarnings(openxlsx::addWorksheet(wb,paste(substr(sheetname,0,31))))
  # sheet name
  i <- paste(substr(sheetname, 0, 31))


  # Style definitions ---------------------

  ## title, subtitle & Header
  style_title <- openxlsx::createStyle(fontSize=14, textDecoration="bold",fontName="Arial")

  style_subtitle <- openxlsx::createStyle(fontSize=12, textDecoration="bold",fontName="Arial")

  style_header <- openxlsx::createStyle(border="Bottom", borderColour = "#009ee0",
                                      borderStyle = getOption("openxlsx.borderStyle", "thick"))

  style_wrap <- openxlsx::createStyle(wrapText = TRUE)


  # Fill Excel -----------------

  ## Title
  openxlsx::writeData(wb, sheet = i,title, headerStyle=style_title,startRow = 7)

  ## Logo

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


  ## Contact details
  if(contactdetails=="statzh"){

    contactdetails <- c("Datashop, Tel: 0432597500",
                        "datashop@statistik.ji.zh.ch",
                        "http://www.statistik.zh.ch")

  }

  openxlsx::writeData(wb, sheet = i,
                      contactdetails,
                      headerStyle = style_wrap,
                      startRow = 2,
                      startCol = 12)

  if (length(contactdetails) > 3) {
    warning("Contactdetails may overlap with other elements. To avoid this issue please do not include more than three elements in the contactdetails vector.")
  }

  ## Source
  if(source == "statzh"){
    source <- "Statistisches Amt des Kantons Zürich"
  }
  openxlsx::writeData(wb, sheet = i, "Datenquelle:", headerStyle=style_subtitle, startRow = 9)
  openxlsx::writeData(wb, sheet = i, source, startRow = 10)

  ## Metadata
  openxlsx::writeData(wb, sheet = i, "Hinweise:", headerStyle=style_subtitle, startRow = 12)
  openxlsx::writeData(wb, sheet = i, metadata, startRow =13)


  ## User
  if(author == "user"){
    # für das lokale R
    if(Sys.getenv("USERNAME")!="") {
      contactperson <- stringr::str_sub(Sys.getenv("USERNAME"), start = 6, end = 7)
    } else {
      # für den R-server
      contactperson <- stringr::str_sub(Sys.getenv("USER"), start = 6, end = 7)
    }
  } else {
    contactperson <- author
  }

  #Aktualisierungsdatum
  openxlsx::writeData(wb, sheet = i, paste("Aktualisiert am ",
                                           format(Sys.Date(), format="%d.%m.%Y"),
                                           " durch: ", contactperson),
                      startRow = 5, startCol=12)


  # Füge Formatierungen ein
  openxlsx::addStyle(wb, sheet = i, style_header, rows = 5, cols = 1:26, gridExpand = TRUE, stack = TRUE)

  openxlsx::addStyle(wb, sheet = i, style_title, rows = 7, cols = 1, gridExpand = TRUE)

  openxlsx::addStyle(wb, sheet = i, style_subtitle, rows = 9, cols = 1, gridExpand = TRUE)

  openxlsx::addStyle(wb, sheet = i, style_subtitle, rows = 12, cols = 1, gridExpand = TRUE)

}
