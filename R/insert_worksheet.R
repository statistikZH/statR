#' Functions for generating formatted worksheets for data
#'
#' @description Inserts a data.frame into a new formatted worksheet.
#'   The distinction between \code{insert_worksheet} and
#'   \code{insert_worksheet_nh} is that the former generates a header with
#'   contact information. By default, contact information is imported from the
#'   user configuration, but all fields can be overridden if needed.
#' @note The function does not write the result into a .xlsx file.
#'  A separate call to openxlsx::saveWorkbook() is required.
#' @param data data to be included.
#' @param wb workbook object to add new worksheet to.
#' @param title title to be put above the data.
#' @param sheetname name of the sheet tab.
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
#' @param logo path of the file to be included as logo (png / jpeg / svg).
#'   Default can be adjusted via user profiles.
#' @param contactdetails contact details of the data publisher. Default can be
#'   adjusted via user profiles.
#' @param homepage Homepage of data publisher. Default can be adjusted via user
#'   profiles.
#' @param author defaults to the last two letters (initials) or numbers of the
#'  internal user name.
#' @examples
#' # Initialize Workbook
#' wb <- openxlsx::createWorkbook()
#'
#'# Insert mtcars dataset with STATZH design
#' insert_worksheet(
#'   wb = wb, sheetname = "cars1", data = mtcars, title = "mtcars dataset",
#'   source = "Source: ...", metadata = "Note: ...",
#'   grouplines = c(5,8), group_names = c("First group", "Second group")
#'
#' # The same, but without header
#' insert_worksheet_nh(
#'   wb, sheetname = "cars2", data = mtcars, title = "mtcars dataset (no header)",
#'   source = "Source: ...", metadata = "Note: ...",
#'   grouplines = c(5,8), group_names = c("First group", "Second group"))
#'
#' @keywords insert_worksheet
#' @export

insert_worksheet <- function(wb, sheetname, data, title, source,
                             metadata, grouplines = NA, group_names = NA,
                             logo = getOption("statR_logo"),
                             contactdetails = inputHelperContactInfo(),
                             homepage = getOption("statR_homepage"),
                             author = "user") {
  UseMethod("insert_worksheet", data)
}

#' @rdname insert_worksheet
#' @keywords internal
insert_worksheet.default <- function(wb, sheetname, data, title, source,
                             metadata, grouplines = NA, group_names = NA,
                             logo = getOption("statR_logo"),
                             contactdetails = inputHelperContactInfo(),
                             homepage = getOption("statR_homepage"),
                             author = "user") {
  # Check that sheetname satisfies the character limit
  sheetname <- verifyInputSheetname(sheetname)

  insert_header(wb, sheetname, logo, contactdetails, homepage, author,
                contact_col = max(ncol(data) - 2, 4))

  insert_worksheet_nh(
    wb, sheetname, data, title, source, metadata, grouplines, group_names)
}

#' @keywords internal
#' @rdname insert_worksheet
insert_worksheet.Content <- function(wb, sheetname, data, title, source,
                                     metadata, grouplines = NA, group_names = NA,
                                     logo = getOption("statR_logo"),
                                     contactdetails = inputHelperContactInfo(),
                                     homepage = getOption("statR_homepage"),
                                     author = "user") {

  # Try to extract sheetname from data object if missing
  if (missing(sheetname) || is.null(sheetname))
    sheetname <- extract_attribute(data, "sheetname", TRUE)


  insert_worksheet.default(wb, sheetname, data, title, source, metadata,
                           grouplines, group_names, logo, contactdetails,
                           homepage, author = "user")
}


#' @keywords insert_worksheet_nh
#' @rdname insert_worksheet
#' @export
insert_worksheet_nh <- function(wb, sheetname, data, title, source, metadata,
                                grouplines = NULL, group_names = NULL) {
  UseMethod("insert_worksheet_nh", data)
}

#' @keywords internal
#' @rdname insert_worksheet
insert_worksheet_nh.Content <- function(wb, sheetname, data, title, source,
                                        metadata, grouplines = NULL,
                                        group_names = NULL) {

  if (missing(title))
    title <- extract_attribute(data, "title")

  if (missing(source))
    source <- extract_attribute(data, "source")

  if (missing(metadata))
    metadata <- extract_attribute(data, "metadata")

  if (missing(grouplines))
    grouplines <- extract_attribute(data, "grouplines")

  if (missing(group_names))
    group_names <- extract_attribute(data, "group_names")

  insert_worksheet_nh.default(wb, sheetname, data, title, source, metadata,
                              grouplines, group_names)
}


#' @keywords internal
#' @rdname insert_worksheet
insert_worksheet_nh.default <- function(
    wb, sheetname, data, title, source, metadata, grouplines = NULL, group_names = NULL) {

  sheetname <- verifyInputSheetname(sheetname)

  if (!(sheetname %in% names(wb))) {
    openxlsx::addWorksheet(wb, sheetname)
    start_row <- 1

  } else {
    start_row <- namedRegionLastRow(wb, sheetname) + 3
  }

  # Insert title, metadata, and sources into worksheet --------
  if (is.character(title)) {
    writeText(wb, sheetname, title, start_row, 1, style_title(), "title")
    start_row <- namedRegionLastRow(wb, sheetname, "title") + 1
  }

  if (is.character(source)) {
    writeText(wb, sheetname, source, start_row, 1, style_subtitle(), "source")
    start_row <- namedRegionLastRow(wb, sheetname, "source") + 1
  }

  if (is.character(metadata)) {
    writeText(wb, sheetname, metadata, start_row, 1, style_subtitle(), "metadata")
    start_row <- namedRegionLastRow(wb, sheetname, "metadata") + 1
  }

  if (is.character(title) || is.character(source) || is.character(metadata)) {

    row_extent <- namedRegionRowExtent(wb, sheetname, c("title", "source", "metadata"))

    ### Merge cells with title, metadata, and sources to ensure that they're displayed properly
    purrr::walk(row_extent, ~openxlsx::mergeCells(wb, sheetname, cols = 1:18, rows = .))

    ### Add Line wrapping
    openxlsx::addStyle(wb, sheetname, style_wrap(), row_extent, 1, stack = TRUE, gridExpand = TRUE)

    # Insert data --------
    data_start_row <- namedRegionLastRow(wb, sheetname, c("title", "source", "metadata")) + 2

  } else {
    data_start_row <- start_row
  }

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

  ### Write data after checking for leftover grouping
  openxlsx::writeData(wb, sheetname, verifyDataUngrouped(data),
                      startRow = data_start_row, rowNames = FALSE,
                      withFilter = FALSE,
                      name = paste(sheetname, "data", sep = "_"))
  openxlsx::addStyle(wb, sheetname, style_header(),
                     data_start_row, 1:ncol(data),
                     gridExpand = TRUE, stack = TRUE)

  # Format --------
  ### Define minimum column width
  options("openxlsx.minWidth" = 5)

  ### Use automatic column width for columns with data
  openxlsx::setColWidths(wb, sheetname, 1:ncol(data), "auto", ignoreMergedCells = TRUE)
}


#' @keywords insert_header
#' @rdname insert_worksheet
insert_header <- function(wb, sheetname, logo = getOption("statR_logo"),
                          contactdetails = inputHelperContactInfo(),
                          homepage = getOption("statR_homepage"),
                          author = "user", contact_col = 13) {

  sheetname <- verifyInputSheetname(sheetname)

  # Add sheet if not existing
  if (!(sheetname %in% names(wb))) {
    openxlsx::addWorksheet(wb, sheetname)
  }

  # Insert logo ------
  insert_worksheet_image(wb = wb, sheetname = sheetname,
                         image = inputHelperLogoPath(logo),
                         startrow = 1, startcol = 1)

  # Insert contact info, date created, and author -----
  ### Contact info
  writeText(wb, sheetname, contactdetails, 2, contact_col, NULL, "contact")

  writeText(wb, sheetname, inputHelperHomepage(homepage),
            namedRegionLastRow(wb, sheetname, "contact") + 1, contact_col,
            NULL, "homepage")

  ### Information string about time of generation and responsible user
  writeText(wb, sheetname, paste(inputHelperDateCreated(), inputHelperAuthorName(author)),
            namedRegionLastRow(wb, sheetname, "homepage") + 1, contact_col,
            NULL, "info")

  ### Horizontally merge cells to ensure that contact entries are displayed properly
  row_extent <- namedRegionRowExtent(wb, sheetname, c("contact", "homepage", "info"))
  purrr::walk(row_extent, ~openxlsx::mergeCells(wb, sheetname, cols = contact_col + 0:3, rows = .))


  ### Insert headerline after contacts ------
  openxlsx::addStyle(wb, sheetname, style_headerline(), namedRegionLastRow(wb, sheetname, "info"),
                     1:contact_col, gridExpand = TRUE, stack = TRUE)
}
