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

  forbidden_chars <- c("/", "\\", "?", "*", ":", "[", "]")
  pattern <- paste0("(", paste0(
    paste0("\\", forbidden_chars), collapse = "|"), ")")

  if (nchar(sheetname) > 31) {
    message("sheetname ", sheetname, "truncated at 31 characters (Excel limit).")
    sheetname <- substr(sheetname, 0, 31)
  }

  if (grepl(pattern, sheetname)) {
    message("Found (and replaced with '_') forbidden characters: ",
            paste(forbidden_chars, collapse = " "))
    sheetname <- gsub(pattern, "_", sheetname)
  }

  return(sheetname)
}

#' verifyInputSheetnames()
#'
#' @description Function which truncates multiple sheetnames to 31 characters and
#'  checks for uniqueness.
#' @details MS Excel imposes a character limit of 31 characters for names of
#'  worksheets. This function truncates sheetnames accordingly and notify the
#'  user via a message. If the truncated sheetnames aren't unique, the function
#'  raises an error.
#' @param sheetname A character string with the name for an XLSX worksheet
#' @returns A character string
#' @keywords internal
#' @noRd
verifyInputSheetnames <- function(sheetnames) {

  output_sheetnames <- lapply(sheetnames, verifyInputSheetname)
  counts <- table(unlist(output_sheetnames))

  if (any(counts > 1)) {
    duplicates <- names(counts)[counts > 1]
    stop("Duplicate sheetnames after truncation for datasets ",
         paste(which(output_sheetnames %in% duplicates), collapse = ", "))
  }

  return(output_sheetnames)
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


#' Set up logo path and attach settings for width and height
#'
#' @description Replace default values "zh" and "statzh" with the file path to
#'  the respective logo (included in /extdata), otherwise returns the input value.
#' @param logo A character string with the file path
#' @returns A character string
#' @keywords internal
#' @noRd
inputHelperLogoPath <- function(
    logo, width = getOption("statR_logo_width"),
    height = getOption("statR_logo_height")) {

  if (is.null(logo)) {
    message("No logo added.")

  } else {

    if (logo == "statzh") {
      logo <- paste0(find.package("statR"), "/extdata/Stempel_Kanton_ZH.png")

    } else if (logo == "zh") {
      logo <- paste0(find.package("statR"), "/extdata/Stempel_STAT-01.png")
    }

    logo <- add_plot_size(logo, c(width, height))
  }

  return(logo)
}


#' Construct a vector with contact information from defaults
#'
#' A helper function which constructs a character vector from the contact
#' information defined in the user profile. If \code{compact = TRUE}, the
#' organization is dropped, and the name and phone number are displayed in the
#' same row.
#' @param compact A boolean. If TRUE, a shortened version is displayed
#' @returns A character vector
#' @export
inputHelperContactInfo <- function(compact = FALSE) {

  phone <- inputHelperPhone(getOption("statR_phone"))

  if (compact) {
    return(c(paste(getOption("statR_name"), phone, sep = ", "),
             getOption("statR_email")))
  }

  return(c(getOption("statR_organization"), getOption("statR_name"),
           phone, getOption("statR_email")))
}

#' Set homepage to a hyperlink
#'
#' Formats homepage as a hyperlink object.
#' @param homepage A character string
#' @returns A character string or a 'hyperlink' object
#' @export
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


#' Determine the row and column extent of named regions
#'
#' @description Get extent of a named region in a workbook object
#' @details If a single name is provided, returns the extent of the associated
#'   named region. If a character vector is provided, the combined extent is
#'   returned instead. If left at default (NULL), the combined extent of all
#'   regions is returned.
#' @param wb A workbook object
#' @param sheetname Name of a worksheet
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

  } else if (all(is.null(named_regions))) {
    stop("No named regions defined.")
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


#' Determine the extent of a named region in a particular direction
#'
#' @description Get row or column extent of a named region in a workbook object.
#' @inheritParams namedRegionExtent
#' @returns A numeric vector of indices
#' @rdname namedRegionExtent1D
#' @keywords internal
namedRegionRowExtent <- function(wb, sheetname, region_name = NULL) {
  unlist(namedRegionExtent(wb, sheetname, region_name, "row"))
}
#' @rdname namedRegionExtent1D
#' @keywords internal
namedRegionColumnExtent <- function(wb, sheetname, region_name = NULL) {
  unlist(namedRegionExtent(wb, sheetname, region_name, "col"))
}


#' Functions to determine the outer boundaries of named regions
#'
#' @description Methods for determining the first and last index of
#'   a namedRegion in row and column directions.
#' @inheritParams namedRegionExtent
#' @returns Numeric value
#' @keywords internal
#' @rdname namedRegionBoundary
namedRegionFirstRow <- function(wb, sheetname, region_name = NULL) {
  min(namedRegionRowExtent(wb, sheetname, region_name))
}
#' @rdname namedRegionBoundary
#' @keywords internal
namedRegionLastRow <- function(wb, sheetname, region_name = NULL) {
  max(namedRegionRowExtent(wb, sheetname, region_name))
}
#' @rdname namedRegionBoundary
#' @keywords internal
namedRegionFirstCol <- function(wb, sheetname, region_name = NULL) {
  min(namedRegionColumnExtent(wb, sheetname, region_name))
}
#' @rdname namedRegionBoundary
#' @keywords internal
namedRegionLastCol <- function(wb, sheetname, region_name = NULL) {
  max(namedRegionColumnExtent(wb, sheetname, region_name))
}


#' cleanNamedRegions()
#'
#' @description Function to clean up unneeded named regions.
#' @details Named regions are used as a convenient tool during the construction
#'   of the output workbook. This function can be used to remove some or all
#'   named regions. Note: this doesn't extend to the data contained in a named
#'   region.
#' @note When working with insert-like functions to construct a custom output,
#'   this function should only be called just before the conversion of the workbook into an .xlsx file.
#' @param wb A workbook object
#' @param which Either "keep_data" (to keep any named regions pertaining to tables),
#'   or "all".
#' @keywords internal
#' @noRd
cleanNamedRegions <- function(wb, which = c("keep_data", "all")) {
  named_regions <- openxlsx::getNamedRegions(wb)
  delete_regions <- c()

  if (which == "keep_data") {
    delete_regions <- named_regions[!grepl("_data", named_regions)]
  } else if (which == "all") {
    delete_regions <- named_regions
  }

  if (length(delete_regions > 0)) {
    purrr::walk(delete_regions, ~openxlsx::deleteNamedRegion(wb, .))
  }
}



#' Convert missing input to NULL
#'
#' @keywords internal
missingToNull <- function(input_value) {
  if (missing(input_value)) {
    return(NULL)
  } else {
    return(input_value)
  }
}
