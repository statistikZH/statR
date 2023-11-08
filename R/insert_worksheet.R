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
#' @param source source of the data. Default can be adjusted via user profiles
#' @param metadata metadata information to be included. Defaults to NA, meaning
#'   no metadata are attached.
#' @param logo path of the file to be included as logo (png / jpeg / svg).
#'   Default can be adjusted via user profiles.
#' @param contactdetails contact details of the data publisher. Default can be
#'   adjusted via user profiles.
#' @param homepage Homepage of data publisher. Default can be adjusted via user
#'   profiles.
#' @param author defaults to the last two letters (initials) or numbers of the
#'  internal user name.
#' @param grouplines Can be used to visually group variables. Values should
#'   either correspond to numeric column indices or column names, denoting the
#'   first variable in a group. Defaults to NA, meaning no lines are added.
#' @param group_names A vector of names for groups to be displayed in a
#'   secondary header. Should be of the same length as grouplines, and cannot
#'   be used unless these are set. Defaults to NA, meaning no secondary header
#'   is created.
#' @examples
#' # Initialize Workbook
#' wb <- openxlsx::createWorkbook()
#'
#'# Insert mtcars dataset with STATZH design
#' insert_worksheet(
#'   wb = wb, sheetname = "carb", data = mtcars, title = "mtcars dataset")
#'
#' @keywords insert_worksheet
#' @export
insert_worksheet <- function(wb,
                             sheetname = "Daten",
                             data,
                             title,
                             source,
                             metadata = NA,
                             logo = getOption("statR_logo"),
                             contactdetails = inputHelperContactInfo(),
                             homepage = getOption("statR_homepage"),
                             author = "user",
                             grouplines = NA,
                             group_names = NA) {


  # Try to fill in values if not provided
  if (missing(sheetname) || is.null(sheetname))
    sheetname <- extract_attribute(data, "sheetname", TRUE)



  # Initialize new worksheet ------
  sheetname <- verifyInputSheetname(sheetname)
  openxlsx::addWorksheet(wb, sheetname)


  # Insert logo ------
  insert_worksheet_image(wb = wb, sheetname = sheetname,
                         image = inputHelperLogoPath(logo),
                         startrow = 1, startcol = 1)


  # Insert contact info, date created, and author -----
  ### Contact info
  contact_start_col <- max(ncol(data) - 2, 4)

  openxlsx::writeData(wb, sheetname, contactdetails,
                      contact_start_col, 2,
                      name = paste(sheetname, "contact", sep = "_"))

  openxlsx::writeData(wb, sheetname, inputHelperHomepage(homepage),
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

  insert_worksheet_nh(
    wb, sheetname = sheetname, data = data, title = title, source = source,
    metadata = metadata, grouplines = grouplines, group_names = group_names)
}


#' insert_worksheet_nh
#'
#' @description Function to add formatted worksheets to an existing Workbook.
#'   The worksheets do not include contact information or logos.
#' @note This function doesn't include any contact information. The function
#'  does not write the result into a .xlsx file. A separate
#'  call to `openxlsx::saveWorkbook()` is required.
#' @inheritParams insert_worksheet
#' @examples
#' ## Create workbook
#' wb <- openxlsx::createWorkbook()
#'
#' ## insert a new worksheet
#' insert_worksheet_nh(wb, sheetname = "data1", data = head(mtcars),
#'   title = "Title", source = "Source: ...", metadata = "Note: ...")
#'
#' ## insert a further worksheet
#' insert_worksheet_nh(wb, sheetname = "data2", data = tail(mtcars),
#'   title = "Title", source = "Source: ...", metadata = "Note: ...")
#'
#' ## insert a worksheet with group lines and second header
#' insert_worksheet_nh(
#'   wb, sheetname = "data3", data = head(mtcars), title = "grouplines",
#'   source = "Source: ...", metadata = "Note: ...",
#'   grouplines = c(5,8), group_names = c("First group", "Second group"))
#' @keywords insert_worksheet_nh
#' @export
insert_worksheet_nh <- function(
    wb, sheetname, data, title, source,
    metadata = NA, grouplines = NA, group_names = NA) {

  sheetname <- verifyInputSheetname(sheetname)

  if (!(sheetname %in% names(wb))) {
    openxlsx::addWorksheet(wb, sheetname)
    start_row <- 1

  } else {
    start_row <- namedRegionLastRow(wb, sheetname) + 3
  }

  # Try to fill in values if not provided

  if (missing(title) || is.null(title)){
    title <- extract_attribute(data, "title")
  }

  if (missing(source) || is.null(source)){
    source <- extract_attribute(data, "source")
  }

  if (missing(metadata) || is.null(metadata)){
    metadata <- extract_attribute(data, "metadata")
  }

  if (missing(grouplines) || is.null(grouplines)){
    grouplines <- extract_attribute(data, "grouplines")
  }

  if (missing(group_names) || is.null(group_names)){
    group_names <- extract_attribute(data, "group_names")
  }


  # Insert title, metadata, and sources into worksheet --------
  ### Title
  if (!is.null(title) && !all(is.na(title))) {
    openxlsx::writeData(wb, sheetname, title, startCol = 1, startRow = start_row,
                        name = paste(sheetname, "title", sep = "_"))
    openxlsx::addStyle(wb, sheetname, style_title(), start_row, 1)
    start_row <- namedRegionLastRow(wb, sheetname, "title") + 1
  }

  ### Source
  if (!is.null(source) && !all(is.na(source))) {
    openxlsx::writeData(
      wb, sheetname, inputHelperSource(source),
      startRow = start_row, name = paste(sheetname, "source", sep = "_"))
    openxlsx::addStyle(
      wb, sheetname, style_subtitle(), namedRegionRowExtent(wb, sheetname, "source"),
      1, stack = TRUE, gridExpand = TRUE)
    start_row <- namedRegionLastRow(wb, sheetname, "source") + 1
  }

  if (!is.null(metadata) && !all(is.na(metadata))) {
    ### Metadata
    openxlsx::writeData(
      wb, sheetname, inputHelperMetadata(metadata),
      startRow = start_row,
      name = paste(sheetname, "metadata", sep = "_"))
    openxlsx::addStyle(
      wb, sheetname, style_subtitle(), namedRegionRowExtent(wb, sheetname, "metadata"),
      1, stack = TRUE, gridExpand = TRUE)
    start_row <- namedRegionLastRow(wb, sheetname, "metadata") + 1
  }

  if (any(c(!is.na(title), !is.na(source), !is.na(metadata)))) {

    ### Merge cells with title, metadata, and sources to ensure that they're displayed properly
    purrr::walk(namedRegionRowExtent(wb, sheetname, c("title", "source", "metadata")),
                ~openxlsx::mergeCells(wb, sheetname, cols = 1:18, rows = .))

    ### Add Line wrapping
    openxlsx::addStyle(wb, sheetname, style_wrap(),
                       namedRegionRowExtent(wb, sheetname, c("title", "source", "metadata")), 1,
                       stack = TRUE, gridExpand = TRUE)

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
