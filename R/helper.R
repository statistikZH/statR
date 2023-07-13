#' verifyInputSheetname()
#'
#' @description Function which truncates a sheetname to 31 characters
#' @details MS Excel imposes a character limit of 31 characters for names of
#'  worksheets. This function truncates sheetnames accordingly and notify the
#'  user via a message.
#' @param sheetname A character string with the name for an XLSX worksheet
#' @returns A character string
#' @keywords internal
#' @noRd
verifyInputSheetname <- function(sheetname) {
  if (nchar(sheetname) > 31) {
    message("sheetname truncated at 31 characters to satisfy MS-Excel limit.")
  }

  return(substr(sheetname, 0, 31))
}


#' verifyInputFilename()
#'
#' @description Function which adds a file extension to a filename if missing .
#' @param filename A character string with the filename
#' @param extension A character string with the file extension. Defaults to ".xlsx"
#' @returns A character string
#' @keywords internal
#' @noRd
verifyInputFilename <- function(filename, extension = ".xlsx") {
  regex_pattern <- paste0(extension, "$")
  paste0(gsub(regex_pattern, "", filename), extension)
}


#' verifyDataUngrouped()
#'
#' @description Function which checks if a data.frame is a grouped_df, in which
#'  case it calls dplyr::ungroup().
#' @param data A character string with the filename
#' @returns A data.frame
#' @keywords internal
#' @noRd
verifyDataUngrouped <- function(data) {
  if (!dplyr::is_grouped_df(data)) {
    return(data)
  }

  return(dplyr::ungroup(data))
}

#' inputHelperSource()
#'
#' @description Substitute default value "statzh" with official title
#' @param source A character vector with source information
#' @param prefix A character string giving prepended to the sources
#' @param collapse Separator for collapsing multiple sources with
#' @returns A character string
#' @keywords internal
#' @noRd
inputHelperSource <- function(source, prefix = getOption("statR_prefix_source"),
                              collapse = getOption("statR_collapse")) {

  if (all(is.na(source))) {
    return("")
  }

  if (!is.null(collapse)) {
    return(paste(prefix, paste0(source, collapse = collapse)))
  }

  return(c(prefix, source))
}


#' inputHelperMetadata()
#'
#' @description Concatenate metadata into a formatted string.
#' @inheritParams inputHelperSource
#' @param metadata A character vector with metadata information
#' @returns A character vector
#' @keywords internal
#' @noRd
inputHelperMetadata <- function(metadata, prefix = getOption("statR_prefix_metadata"),
                                collapse = getOption("statR_collapse")) {

  if (all(is.na(metadata))) {
    return("")
  }

  if (!is.null(collapse)) {
    return(paste(prefix, paste0(metadata, collapse = collapse)))
  }

  return(c(prefix, metadata))
}


#' inputHelperLogoPath()
#'
#' @description Replace default values "zh" and "statzh" with the file path to
#'  the respective logo (included in /extdata), otherwise returns the input value.
#' @param logo A character string with the file path
#' @returns A character string
#' @keywords internal
#' @noRd
inputHelperLogoPath <- function(logo) {
  if (is.null(logo)) {
    message("No logo added.")

  } else if (logo == "statzh") {
    logo <- paste0(find.package("statR"), "/extdata/", statzh_logo)

  } else if (logo == "zh") {
    logo <- paste0(find.package("statR"), "/extdata/", zh_logo)
  }

  return(logo)
}


#' inputHelperContactInfo()
#'
#' @description Replaces default value "statzh" with the contact information of
#'  the Statistics Office of Canton Zurich, otherwise returns the input value.
#' @param compact A boolean which controls the format of the contact information. Default: FALSE
#' @returns A character vector
#' @keywords internal
#' @noRd
inputHelperContactInfo <- function(compact = FALSE) {

  phone <- inputHelperPhone(getOption("statR_phone"))

  if (compact) {
    return(c(paste(getOption("statR_name"), phone, sep = ", "),
             getOption("statR_email")))
  }

  return(c(getOption("statR_organization"), getOption("statR_name"),
           phone, getOption("statR_email")))
}


#' inputHelperOfficeHours()
#'
#' @description Replace default value "statzh" with the office hours
#'  of the Statistics Office of Canton Zurich, otherwise returns
#'  the input value.
#' @param openinghours A character vector with opening hours.
#' @returns A character vector
#' @keywords internal
#' @noRd
inputHelperOfficeHours <- function(openinghours) {
  if (openinghours == "statzh") {
    openinghours <- statzh_openinghours
  }

  return(openinghours)
}


#' inputHelperHomepage()
#'
#' @description Formats homepage as a hyperlink object.
#' @param homepage A character string
#' @returns A character string or a 'hyperlink' object
#' @keywords internal
#' @noRd
inputHelperHomepage <- function(homepage) {
  if (!is.null(homepage)) {
    class(homepage) <- "hyperlink"
  }

  return(homepage)
}

#' inputHelperPhone()
#'
#' @description Formats phone number
#' @param homepage A character string
#' @returns A character string or a 'hyperlink' object
#' @keywords internal
#' @noRd
inputHelperPhone <- function(phone, prefix = getOption("statR_prefix_phone")) {
  if (!is.null(phone)) {
    phone <- paste(prefix, phone)
  }

  return(phone)
}


#' inputHelperDateCreated()
#'
#' @description Returns current date as a string with format specified by date_format
#' @param prefix A Prefix to prepend to the date. Default to NULL
#' @param date_format A character string for the date_format. Default: "%d.%m.%Y".
#' @returns A character vector
#' @keywords internal
#' @seealso format
#' @noRd
inputHelperDateCreated <- function(prefix = getOption("statR_prefix_date"),
                                   date_format = getOption("date_format")) {
  paste(prefix, format(Sys.Date(), format = date_format))
}


#' inputHelperOrderNumber()
#'
#' @description Returns current date as a string with format specified by date_format
#' @param prefix A Prefix to prepend to the date. Default to NULL
#' @returns A character vector
#' @keywords internal
#' @seealso format
#' @noRd
inputHelperOrderNumber <- function(order_num,
                                   prefix = getOption("statR_prefix_order_id")) {
  if (!is.null(order_num)) {
    order_num <- paste(prefix, order_num)
  }

  return(order_num)
}


#' inputHelperAuthorName()
#'
#' @description Function extracts initials from global environment if
#'  input is 'user', otherwise passes on the input.
#' @param author Name of author
#' @param prefix A character string to prepend to the result
#' @returns A character string with the username
#' @keywords internal
#' @noRd
#'
inputHelperAuthorName <- function(author,
                                  prefix = getOption("statR_prefix_author")) {
  if (author == "user") {
    sys_vals <- c(Sys.getenv("USERNAME"), Sys.getenv("USER"))
    author_name <- sys_vals[which(sys_vals != "")[1]]
    author <- stringr::str_sub(author_name, start = 6, end = 7)
  }

  return(paste(prefix, author))
}



#' excelIndexToRowCol()
#'
#' @description Converts an Excel style index (e.g. A1) into numeric row and
#'   column indices. Handles cells (A1) as well as matrices (A1:B2)
#' @param index A Microsoft Excel style index (see details)
#' @returns A list containing row indices and column indices
#' @keywords internal
#' @importFrom utils stack
excelIndexToRowCol <- function(index) {

  splitIndex <- function(x, split = "") unlist(strsplit(x, split))

  excelColumnLetterToNumeric <- function(x) {
    chars <- splitIndex(x, "")
    offsets <- 26^(length(chars):1 - 1)
    sum(offsets * match(chars, LETTERS))
  }

  # Return extent of combined region
  if (length(index) > 1) {
    extents <- lapply(index, excelIndexToRowCol)
    extents <- unique(do.call(rbind, lapply(extents, stack)))
    rows <- extents[extents$ind == "row", "values"]
    cols <- extents[extents$ind == "col", "values"]
    return(list(row = seq(min(rows), max(rows)),
                col = seq(min(cols), max(cols))))
  }

  # Single index
  column_index <- sapply(splitIndex(gsub("([0-9]+)", "", index), ":"),
                         excelColumnLetterToNumeric)

  rows <- eval(parse(text = gsub("([A-Z]{1,3})", "", index)))
  cols <- seq(min(column_index), max(column_index))
  return(list(row = rows, col = cols))
}


#' namedRegionExtent()
#'
#' @description Get extent of a named region in a workbook object
#' @details If a single name is provided, returns the extent of the associated
#'   named region. If a character vector is provided, the combined extent is
#'   returned instead. If left at default (NULL), the combined extent of all
#'   regions is returned.
#' @param wb A workbook object
#' @param sheet Name of a worksheet
#' @param region_name names of regions in Workbook.
#' @param which either "row", "col", or "both" (default).
#' @returns A list with two numeric vectors row and col, containing
#'   row and column indices.
#' @keywords internal
namedRegionExtent <- function(wb, sheetname, region_name = NULL,
                              which = "both") {
  named_regions <- openxlsx::getNamedRegions(wb)

  if (!(sheetname %in% names(wb))) {
    stop("Sheetname does not exist in Workbook")
  }

  sheet_ind <- which(sheetname == attr(named_regions, "sheet"))
  region_names <- named_regions[sheet_ind]


  positions <- attr(named_regions, "position")[sheet_ind]

  if (which == "both") {
    dimensions <- c("row", "col")
  } else {
    dimensions <- which
  }

  # If name is null, return full extent for sheet
  if (all(is.null(region_name))) {
    return(excelIndexToRowCol(positions)[dimensions])
  }

  region_ind <- unlist(sapply(region_name, function(region) {
    which(grepl(paste0(region, "$"), region_names))
  }))

  # Return extent for combined region
  if (length(region_ind) > 0) {
    return(excelIndexToRowCol(positions[region_ind])[dimensions])
  }

  # Default to A1
  return(excelIndexToRowCol("A1")[dimensions])
}


#' namedRegionRowExtent()
#'
#' @description Get row extent of a named region in a workbook object
#' @inheritParams namedRegionExtent
#' @returns A numeric vector of row indices
#' @keywords internal
#'
namedRegionRowExtent <- function(wb, sheetname, region_name = NULL) {
  unlist(namedRegionExtent(wb, sheetname, region_name, "row"))
}


#' namedRegionColumnExtent()
#'
#' @description Get column extent of a named region in a Workbook object
#' @inheritParams namedRegionExtent
#' @returns A numeric vector of column indices
#' @keywords internal
#'
namedRegionColumnExtent <- function(wb, sheetname, name = NULL) {
  unlist(namedRegionExtent(wb, sheetname, name, "col"))
}


#' namedRegionFirstRow()
#'
#' @description Get first row number of named region
#' @inheritParams namedRegionExtent
#' @returns Numeric value corresponding to first row of named region
#' @keywords internal
namedRegionFirstRow <- function(wb, sheet, region_name = NULL) {
  min(namedRegionRowExtent(wb, sheet, region_name))
}


#' namedRegionLastRow()
#'
#' @description Get last row number of named region
#' @inheritParams namedRegionExtent
#' @returns Numeric value corresponding to last row of named region
#' @keywords internal
namedRegionLastRow <- function(wb, sheet, region_name = NULL) {
  max(namedRegionRowExtent(wb, sheet, region_name))
}


#' namedRegionFirstCol()
#'
#' @description Get first col number of named region
#' @inheritParams namedRegionExtent
#' @returns Numeric value corresponding to first column of named region
#' @keywords internal
namedRegionFirstCol <- function(wb, sheet, region_name = NULL) {
  min(namedRegionColumnExtent(wb, sheet, region_name))
}


#' namedRegionLastCol()
#'
#' @description Get last col number of named region
#' @inheritParams namedRegionExtent
#' @returns Numeric value corresponding to last column of named region
#' @keywords internal
namedRegionLastCol <- function(wb, sheet, region_name = NULL) {
  max(namedRegionColumnExtent(wb, sheet, region_name))
}
