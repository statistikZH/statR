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

  output_sheetnames <- sapply(sheetnames, verifyInputSheetname)
  counts <- table(output_sheetnames)

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

#' inputHelperSource()
#'
#' @description Substitute default value "statzh" with official title
#' @param source A character vector with source information
#' @param prefix A character string giving prepended to the sources
#' @param collapse Separator for collapsing multiple sources with
#' @returns A character string
#' @keywords internal
#' @noRd
inputHelperSource <- function(source, prefix = NULL, collapse = NULL) {

  collapse <- c(collapse, attr(source, "collapse"), NA)[1]
  prefix <- c(prefix, attr(source, "prefix"))[1]

  if (all(is.na(source))) {
    return(NULL)
  }

  if (!is.null(collapse) && !is.na(collapse)) {
    return(paste0(prefix, paste0(source, collapse = collapse)))
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
inputHelperMetadata <- function(metadata, prefix = NULL, collapse = NULL) {

  collapse <- c(collapse, attr(metadata, "collapse"), NA)[1]
  prefix <- c(prefix, attr(metadata, "prefix"))[1]

  if (all(is.na(metadata))) {
    return(NULL)
  }

  if (!is.null(collapse) && !is.na(collapse)) {
    return(paste0(prefix, paste0(metadata, collapse = collapse)))
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

    attr(logo, "plot_width") <- width
    attr(logo, "plot_height") <- height

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


#' Helper functions for adding attributes to input objects
#'
#' @description A set of helper functions which serve the purpose of allowing
#'   users to assign information like sheetnames, titles, sources, metadata, and
#'   more to an object. These attributes are then used in place of the function
#'   arguments.
#'
#' @details
#'   \code{add_sheetname}: Adds a sheetname.
#'
#'   \code{add_title}: Adds a title
#'
#'   \code{add_source}: Adds a source. Behavior can be modified further via the
#'     prefix and collapse argument. Setting collapse to a value other than NULL
#'     will result in the input provided in value being concatenated.
#'
#'   \code{add_metadata}: Adds metadata. Behavior can be modified further via the
#'     prefix and collapse argument, just like with \code{add_source}.
#'
#'   \code{add_grouplines}: Adds group lines. Throws an error if input object is
#'     not a data.frame.
#'
#'   \code{add_group_names}: Adds group names. Throws an error if group lines
#'     have not been set.
#'
#'   \code{add_plot_height}: Adds plot height. Throws an error if input object is
#'     not a valid plot type.
#'
#'   \code{add_plot_width}: Adds plot width. Throws an error if input object is
#'     not a valid plot type.
#'
#'   \code{add_plot_size}: Assigns both height and width in one call. Expects a
#'     vector of length 2, with the first element corresponding to width, and the
#'     second to height.
#' @param object The object to add an attribute to
#' @param value A value
#' @param prefix A prefix, default to NULL
#' @param collapse Separator to collapse character vectors on, defaults to NULL
#' @rdname add_attribute
#' @examples
#' library(dplyr)
#' library(ggplot2)
#'
#' # Load data and assign attributes in one pipeline
#' df <- mtcars %>%
#'   add_sheetname("Cars") %>%
#'   add_title("Motor Trend Car Road Tests") %>%
#'   add_source(
#'     c("Henderson and Velleman (1981),",
#'       "Building multiple regression models interactively.",
#'       "Biometrics, 37, 391–411."),
#'     prefix = "Source:",
#'     collapse = " ") %>%
#'   add_metadata(
#'     c("The data was extracted from the 1974 Motor",
#'       "Trend US magazine and comprises fuel consumption",
#'       "and 10 aspects of automobile design and",
#'       "performance for 32 automobiles (1973–74 models)."),
#'     collapse = " ")
#'
#' # Create a plot and assign attributes in one pipeline
#' plt <- (ggplot(mtcars) +
#'   geom_histogram(aes(x = hp))) %>%
#'     add_sheetname("PS") %>%
#'     add_title("Histogram of horsepower") %>%
#'     add_plot_size(c(6, 3))
#'
#' \dontrun{
#' # Generate outputfile using minimal call
#' datasetsXLSX(
#'   file = tempfile(fileext = ".xlsx"),
#'   datasets = list(df, plt))
#' }
#' @export
add_sheetname <- function(object, value) {
  attr(object, "sheetname") <- value
  return(object)
}
#' @rdname add_attribute
#' @export
add_title <- function(object, value) {
  attr(object, "title") <- value
  return(object)
}
#' @rdname add_attribute
#' @export
add_source <- function(object, value, prefix = NULL, collapse = NULL) {

  attr(value, "prefix") <- prefix
  attr(value, "collapse") <- collapse
  attr(object, "source") <- value

  return(object)
}
#' @rdname add_attribute
#' @export
add_metadata <- function(object, value, prefix = NULL, collapse = NULL) {

  attr(value, "prefix") <- prefix
  attr(value, "collapse") <- collapse
  attr(object, "metadata") <- value

  return(object)
}
#' @rdname add_attribute
#' @export
add_grouplines <- function(object, value) {
  if (!is.data.frame(object)) {
    stop("Cannot only assign grouplines for data")
  }

  attr(object, "grouplines") <- value
  return(object)
}
#' @rdname add_attribute
#' @export
add_group_names <- function(object, value) {
  if (is.null(extract_attribute(object, "grouplines"))) {
    stop("Group names can only be used when grouplines have been defined.")
  }

  attr(object, "group_names") <- value
  return(object)
}
#' @rdname add_attribute
#' @export
add_plot_height <- function(object, value) {

  if (!(is.character(object) | inherits(object, c("gg", "ggplot")))) {
    stop("Can only assign plot dimension to valid plot types")
  }

  attr(object, "plot_height") <- value
  return(object)
}
#' @rdname add_attribute
#' @export
add_plot_width <- function(object, value) {

  if (!(is.character(object) | inherits(object, c("gg", "ggplot")))) {
    stop("Can only assign plot dimension to valid plot types")
  }

  attr(object, "plot_width") <- value
  return(object)
}
#' @rdname add_attribute
#' @export
add_plot_size <- function(object, value) {

  if (!(is.character(object) | inherits(object, c("gg", "ggplot")))) {
    stop("Can only assign plot dimension to valid plot types")
  }
  if (length(value) != 2) {
    stop("Expected 2 values but got ", length(value))
  }

  attr(object, "plot_width") <- value[1]
  attr(object, "plot_height") <- value[2]
  return(object)
}

#' Extract attributes from target object
#'
#' @description Function to extract attributes from objects.
#'   \code{extract_attribute} expects the input object to be the target. When
#'   the target object is nested in a list (as is the case in
#'   \code{datasetsXLSX}), \code{extract_attributes} should be used instead.
#' @param object Object to extract attribute from. In practice a data.frame,
#'   ggplot object, or a character string providing a path to an image.
#' @param object_list A list of objects of arbitrary types.
#' @param which Name of the attribute to extract.
#' @param required_val Boolean, if TRUE tries to look up a default in global
#'   options if attribute not found, and raises an error if none was defined.
#' @rdname extract_attribute
#' @keywords internal
extract_attribute <- function(object, which, required_val = FALSE) {

  value <- attr(object, which)

  if (all(is.null(value))) {

    if (required_val) {
      value <- getOption(paste0("statR_default_", which))

      if (all(is.null(value))) {
        stop("No default value found for required argument ", which)
      }

    } else {
      value <- NA
    }
  }

  return(value)
}
#' @rdname extract_attribute
#' @keywords internal
extract_attributes <- function(object_list, which, required_val = FALSE) {
  values <- list()

  for (i in seq_along(object_list)) {
    values[[i]] <- tryCatch(
      extract_attribute(object_list[[i]], which, required_val),
      error = function(err) {
        stop("missing value in required field '", which, "' of dataset ", i)
      })
  }

  return(values)
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
