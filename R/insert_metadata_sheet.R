#' insert_metadata_sheet()
#'
#' Function to  add formatted metadata information to an existing .xlsx-workbook.
#' @param wb workbook object to add new worksheet to.
#' @param title title to be put above the data.
#' @param sheetname name of the sheet tab.
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata information to be included. Defaults to NA.
#' @param logo path of the file to be included as logo (png / jpeg / svg). Defaults to "statzh"
#' @param contactdetails contact details of the data publisher. Defaults to "statzh".
#' @param author defaults to the last two letters (initials) or numbers of the internal user name.
#' @importFrom dplyr "%>%"
#' @keywords insert_metadata_sheet, openxlsx
#' @export
#' @examples
#' # Generation of a spreadsheet
#' wb <- openxlsx::createWorkbook()
#' insert_metadata_sheet(wb, title = "Title of mtcars",
#'   metadata = c("Meta data information."))
#' \dontrun{
#' openxlsx::saveWorkbook(wb,"insert_metadata_sheet.xlsx")
#' }
insert_metadata_sheet <- function(wb, sheetname = "Metadaten", title = "Title",
  source = "statzh", metadata = NA, logo = "statzh", contactdetails = "statzh",
  author = "user"){

  sheetname <- check_sheetname(sheetname)
  openxlsx::addWorksheet(wb, sheetName = sheetname)


  ## Content

  ### Logo ------------
  if(!is.null(logo)){
    logo <- prep_logo(logo)

    if (file.exists(logo)){
      openxlsx::insertImage(wb, sheet = sheetname, file = logo,
                            width = 2.145, height = 0.7865, units = "in")
    } else {
      message("no logo found.")
    }
  }

  ### Contact details --------------
  openxlsx::writeData(wb, sheet = sheetname, x = prep_contact(contactdetails),
                      headerStyle = style_wrap(), startRow = 2, startCol = 12)

  ### Title -----------------
  openxlsx::writeData(wb, sheet = sheetname, x = title, headerStyle = style_title(), startRow = 7)

  ### Source ------------
  openxlsx::writeData(wb, sheet = sheetname, x = "Datenquelle:", headerStyle = style_subtitle2(), startRow = 9)
  openxlsx::writeData(wb, sheet = sheetname, x = prep_source(source, prefix = NULL), startRow = 10)

  ### Metadata ----------
  openxlsx::writeData(wb, sheet = sheetname, x = "Hinweise:", headerStyle = style_subtitle2(), startRow = 12)
  openxlsx::writeData(wb, sheet = sheetname, x = prep_metadata(metadata, prefix = NULL), startRow = 13)


  ### User -----------
  if (author == "user"){
    # for the local R setup
    if (Sys.getenv("USERNAME") != "") {
      contactperson <- stringr::str_sub(Sys.getenv("USERNAME"), start = 6, end = 7)
    } else {
      # for the R server setup
      contactperson <- stringr::str_sub(Sys.getenv("USER"), start = 6, end = 7)
    }
  } else {
    contactperson <- author
  }


  ### Aktualisierungsdatum --------
  openxlsx::writeData(wb, sheet = sheetname,
                      x = paste(prep_creationdate(prefix = "Aktualisiert am:"),
                                "durch:", contactperson), startRow = 5, startCol = 12)

  ## Format

  ### add formatting ---------
  openxlsx::addStyle(wb, sheet = sheetname, style = style_headerline(),
                     rows = 5, cols = 1:26, gridExpand = TRUE, stack = TRUE)
  openxlsx::addStyle(wb, sheet = sheetname, style = style_title(),
                     rows = 7, cols = 1, gridExpand = TRUE)
  openxlsx::addStyle(wb, sheet = sheetname, style = style_subtitle2(),
                     rows = c(9, 12), cols = 1, gridExpand = TRUE)

  ## Remove gridlines ----------
  openxlsx::showGridLines(wb, sheet = sheetname, showGridLines = FALSE)
}
