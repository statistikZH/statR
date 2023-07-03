#' insert_worksheet()
#'
#' @description Function to insert a formatted worksheet into an existing
#'  Workbook object.
#' @note The function does not write the result into a .xlsx file.
#'  A separate call to openxlsx::saveWorkbook() is required.
#' @param data data to be included.
#' @param wb workbook object to add new worksheet to.
#' @param title title to be put above the data.
#' @param sheetname name of the sheet tab.
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata information to be included. Defaults to NA.
#' @param logo path of the file to be included as logo (png / jpeg / svg).
#'  Defaults to "statzh"
#' @param contactdetails contact details of the data publisher. Defaults to
#'  "statzh".
#' @param homepage Homepage of data publisher. Defaults to "statzh".
#' @param grouplines defaults to NA. Can be used to separate grouped
#'  variables visually.
#' @param author defaults to the last two letters (initials) or numbers of the
#'  internal user name.
#' @param group_names Names for groupings in secondary header. Format:
#'   List e.g list(c("title 1", "title 2", "title 3"))
#' @examples
#' # Initialize Workbook
#' wb <- openxlsx::createWorkbook()
#'
#'# Insert mtcars dataset with STATZH design
#' insert_worksheet(data = mtcars,
#'                  wb = wb,
#'                  title = "mtcars dataset",
#'                  sheetname = "carb")
#'
#' @keywords insert_worksheet
#' @export
insert_worksheet <- function(wb,
                             sheetname = "Daten",
                             data,
                             title = "Title",
                             source = "statzh",
                             metadata = NA,
                             logo = "statzh",
                             contactdetails = "statzh",
                             homepage = "statzh",
                             author = "user",
                             grouplines = NA,
                             group_names = NA){

  # Initialize new worksheet ------
  sheetname <- verifyInputSheetname(sheetname)
  openxlsx::addWorksheet(wb, sheetname)


  # Insert logo ------
  insert_worksheet_image(wb = wb, sheetname = sheetname,
                         image = inputHelperLogoPath(logo),
                         startrow = 1, startcol = 1,
                         width = 2.145, height = 0.7865)


  # Insert contact info, date created, and author -----
  ### Contact info
  contact_start_col <- max(ncol(data) - 2, 4)

  openxlsx::writeData(wb, sheetname,
                      x = inputHelperContactInfo(contactdetails),
                      contact_start_col, 2,
                      name = paste(sheetname, "contact", sep = "_"))

  openxlsx::writeData(wb, sheetname,
                      x = inputHelperHomepage(homepage),
                      contact_start_col,
                      namedRegionLastRow(wb, sheetname, "contact") + 1,
                      name = paste(sheetname, "homepage", sep = "_"))

  ### Information string about time of generation and responsible user
  openxlsx::writeData(wb, sheetname,
                      paste(
                        inputHelperDateCreated(),
                        inputHelperAuthorName(author)
                      ),
                      contact_start_col,
                      startRow = namedRegionLastRow(wb, sheetname, "homepage") + 1,
                      name = paste(sheetname, "info", sep = "_"))


  ### Horizontally merge cells to ensure that contact entries are displayed properly
  purrr::walk(namedRegionRowExtent(wb, sheetname, c("contact", "homepage", "info")),
              ~openxlsx::mergeCells(wb, sheetname,
                                    cols = contact_start_col:26,
                                    rows = .))


  ### Insert headerline after contacts ------
  openxlsx::addStyle(wb, sheetname, style_headerline(),
                     rows = namedRegionLastRow(wb, sheetname, "info"),
                     cols = 1:ncol(data), gridExpand = TRUE, stack = TRUE)

  insert_worksheet_nh(wb,
                      sheetname = sheetname,
                      data = data,
                      title = title,
                      source = source,
                      metadata = metadata,
                      grouplines = grouplines,
                      group_names = group_names)
}
