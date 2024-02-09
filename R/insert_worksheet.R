#' Functions for generating formatted worksheets for data
#'
#' @description Inserts a data.frame into a new formatted worksheet.
#'   The distinction between \code{insert_worksheet} and
#'   \code{insert_worksheet_nh} is that the former also generates a header with
#'   contact information using the function \code{insert_header}.
#'   Some arguments have default values which are pulled from the active user
#'   configuration. When using the default configuration, the logo and contact
#'   information of the Statistical Office of Kanton Zurich will be inserted.
#'
#' @note This function does not output an .xlsx file on its own. A separate call
#'   to \code{openxlsx::saveWorkbook()} is required.
#' @param data data to be included.
#' @param wb workbook object to add new worksheet to.
#' @param title title to be put above the data.
#' @param sheetname Names of the worksheet in output file. Note that this name
#'   will be truncated to 31 characters, must be unique, and cannot contain
#'   some special characters (namely the following: /, \, ?, *, :, [, ]).
#' @param source source of the data. Default can be adjusted via user profiles
#' @param metadata metadata information to be included. Defaults to NA, meaning
#'   no metadata are attached.
#' @param grouplines Can be used to visually group variables. Values should
#'   either correspond to numeric column indices or column names, denoting the
#'   first variable in a group. Defaults to NA, meaning no lines are added.
#' @param group_names A vector of names for groups to be displayed in a
#'   secondary header. Should be of the same length as grouplines, and cannot
#'   be used unless these are set. Defaults to NA, meaning no secondary header
#'   is created.
#' @param logo File path to the logo to be included in the index-sheet.
#'   Defaults to the logo of the Statistical Office of Kanton Zurich. This can
#'   either be overridden with a path to an image file, or configured in a user
#'   profile.
#' @param contactdetails Character vector with contact information to be displayed
#'  on the title sheet. By default uses \code{inputHelperContactInfo()} to
#'  construct it based on values defined in the active user configuration.
#' @param homepage Homepage of data publisher. Default can be adjusted via user
#'   configuration.
#' @param author defaults to the last two letters (initials) or numbers of the
#'  internal user name.
#' @param auftrag_id An ID associated with the Excel file. Defaults to NULL (
#'   no output).
#' @param openinghours A character vector with office hours. Defaults to NULL (
#'   no output).
#' @param contact_col Column number at which the contact information should be
#'   inserted.
#' @examples
#' \dontrun{
#' # Initialize Workbook
#' wb <- openxlsx::createWorkbook()
#'
#' # Insert mtcars dataset with STATZH design
#' insert_worksheet(
#'   wb, sheetname = "cars1", data = mtcars, title = "mtcars dataset",
#'   source = "Source: ...", metadata = "Note: ...",
#'   grouplines = c(5,8), group_names = c("First group", "Second group")
#' )
#'
#' # The same, but without header
#' insert_worksheet_nh(
#'   wb, sheetname = "cars2", data = mtcars, title = "mtcars dataset (no header)",
#'   source = "Source: ...", metadata = "Note: ...",
#'   grouplines = c(5,8), group_names = c("First group", "Second group")
#' )
#' }
#'
#' @keywords insert_worksheet
#' @export
insert_worksheet <- function(wb, sheetname, data, title = NULL,
                             source = NULL, metadata = NULL,
                             grouplines = NULL, group_names = NULL,
                             logo = getOption("statR_logo"),
                             contactdetails = inputHelperContactInfo(),
                             homepage = getOption("statR_homepage"),
                             author = "user") {

  sheetname <- verifyInputSheetname(sheetname)
  insert_header(wb, sheetname, logo, contactdetails, homepage, NULL, author,
                NULL, contact_col = max(ncol(data) - 2, 4))
  insert_worksheet_nh(wb, sheetname, data, title = title, source = source,
                      metadata = metadata, grouplines = grouplines,
                      group_names = group_names)
}

#' @rdname insert_worksheet
#' @export
insert_worksheet_nh <- function(wb, sheetname, data, title = NULL, source = NULL,
                                metadata = NULL, grouplines = NULL,
                                group_names = NULL) {

  for (value in c("title", "source", "metadata", "grouplines", "group_names")) {
    if (is.null(eval(as.name(value)))) {
      assign(value, extract_attribute(data, value))
    }
  }

  sheetname <- verifyInputSheetname(sheetname)

  if (!(sheetname %in% names(wb))) {
    openxlsx::addWorksheet(wb, sheetname)
    start_row <- 1

  } else {
    start_row <- namedRegionLastRow(wb, sheetname) + 3
  }

  openxlsx::createNamedRegion(wb, sheetname, 1, start_row, paste0(sheetname, "_content_start"))

  # Insert title, metadata, and sources into worksheet --------
  if (is.character(title)) {
    writeText(wb, sheetname, title, start_row, 1:18, style_title(), "title")
    start_row <- namedRegionLastRow(wb, sheetname, "title") + 1
  }

  if (is.character(source)) {
    writeText(wb, sheetname, source, start_row, 1:18, style_subtitle(), "source")
    start_row <- namedRegionLastRow(wb, sheetname, "source") + 1
  }

  if (is.character(metadata)) {
    writeText(wb, sheetname, metadata, start_row, 1:18, style_subtitle(), "metadata")
    start_row <- namedRegionLastRow(wb, sheetname, "metadata") + 1
  }

  data_start_row <- max(namedRegionLastRow(wb, sheetname, c("content_start, title", "source", "metadata")) + 2,
                        start_row)

  # Grouplines ---------
  if (!any(is.null(grouplines)) & !any(is.na(grouplines))) {
    if (is.numeric(grouplines)) {
      groupline_numbers <- grouplines

    } else if (is.character(grouplines)) {
      groupline_numbers <- match(grouplines, colnames(data))
    }

    ### Insert second header
    if (!any(is.null(group_names)) & !any(is.na(group_names))) {
      insert_second_header(wb, sheetname, data_start_row, group_names, grouplines, data)
      data_start_row <- data_start_row + 1
    }

    data_row_extent <- data_start_row + 0:nrow(data)
    openxlsx::addStyle(wb, sheetname, style_leftline(),
                       data_row_extent, groupline_numbers,
                       gridExpand = TRUE, stack = TRUE)
  }

  ### Pad colnames using whitespaces for better auto-fitting of column width
  colnames(data) <- paste0(colnames(data), "  ", sep = "")

  openxlsx::writeData(wb, sheetname, verifyDataUngrouped(data),
                      startRow = data_start_row, rowNames = FALSE,
                      withFilter = FALSE,
                      name = paste(sheetname, "data", sep = "_"))
  openxlsx::addStyle(wb, sheetname, style_header(), data_start_row,
                     1:ncol(data), gridExpand = TRUE, stack = TRUE)

  ### Define minimum column width
  options("openxlsx.minWidth" = 5)

  ### Use automatic column width for columns with data
  openxlsx::setColWidths(wb, sheetname, 1:ncol(data), "auto",
                         ignoreMergedCells = TRUE)
}

#' @rdname insert_worksheet
#' @export
insert_header <- function(wb, sheetname, logo = getOption("statR_logo"),
                          contactdetails = inputHelperContactInfo(),
                          homepage = getOption("statR_homepage"),
                          auftrag_id = NULL, author = "user",
                          openinghours = NULL, contact_col = 13) {

  logo <- inputHelperLogoPath(getOption("statR_logo"))


  sheetname <- verifyInputSheetname(sheetname)
  if (!(sheetname %in% names(wb))) openxlsx::addWorksheet(wb, sheetname)

  # Insert logo ------
  insert_worksheet_image(wb, sheetname, image = logo,
                         startrow = 1, startcol = 1)

  start_row <- 2
  openxlsx::createNamedRegion(wb, sheetname, contact_col + 0:4, start_row,
                              paste0(sheetname, "_header_start"))

  fields <- list(
    contact = unlist(contactdetails),
    homepage = inputHelperHomepage(homepage),
    info = c(paste(inputHelperDateCreated(), inputHelperAuthorName(author)),
             inputHelperOrderNumber(auftrag_id))
  )

  fields <- fields[!is.na(fields)]

  # Insert contact info, date created, and author -----
  for (field_name in names(fields)){
    if (is.character(fields[[field_name]])) {
      writeText(wb, sheetname, fields[[field_name]], start_row, contact_col + 0:4,
                NULL, field_name)
      start_row <- namedRegionLastRow(wb, sheetname, field_name) + 1
    }
  }

  # Needs to be handled separately
  if (is.character(openinghours)) {
    writeText(wb, sheetname, openinghours,
              namedRegionFirstRow(wb, sheetname, "header_start"),
              contact_col + 5:7, NULL, "openinghours")
  }

  header_entries <- c("header_start", "contact", "homepage", "info",
                      "openinghours")
  openxlsx::createNamedRegion(wb, sheetname,
                    namedRegionColumnExtent(wb, sheetname, header_entries),
                    namedRegionRowExtent(wb, sheetname, header_entries),
                    paste0(sheetname, "_header_body"))

  openxlsx::addStyle(wb, sheetname, style_headerline(), start_row,
                     1:namedRegionLastCol(wb, sheetname, "header_body"),
                     gridExpand = TRUE, stack = TRUE)
}
