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
verifyInputSheetname <- function(sheetname){
  if (nchar(sheetname) > 31){
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
verifyInputFilename <- function(filename, extension = ".xlsx"){
  regex_pattern <- paste0(extension,"$")
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
verifyDataUngrouped <- function(data){
  if (!dplyr::is_grouped_df(data)){
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
#'
inputHelperSource <- function(source, prefix = NULL, collapse = NULL){
  source <- sub("statzh", statzh_name, source)

  if (all(is.na(source))){
    return(source)
  }

  if (!is.null(collapse)){
    return(paste0(prefix, paste0(source, collapse = collapse)))
  }

  return(c(prefix, source))
}


#' inputHelperMetadata()
#'
#' @description Concatenate metadata into a formatted string.
#' @param metadata A character vector with metadata information
#' @param extension A character string with the prefix. Default: "Metadaten: ".
#' @param collapse A character string with the separator for metadata. Default: ";"
#' @returns A character vector if
#' @keywords internal
#' @noRd
inputHelperMetadata <- function(metadata, prefix = NULL, collapse = NULL){

  if (all(is.na(metadata))){
    return(NULL)
  }

  if (!is.null(collapse)){
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
inputHelperLogoPath <- function(logo){
  if (is.null(logo)){
    message("No logo added.")

  } else if (logo == "statzh"){
    logo <- paste0(find.package("statR"), "/extdata/", statzh_logo)

  } else if (logo == "zh"){
    logo <- paste0(find.package("statR"), "/extdata/", zh_logo)
  }

  return(logo)
}


#' inputHelperContactInfo()
#'
#' @description Replaces default value "statzh" with the contact information of
#'  the Statistics Office of Canton Zurich, otherwise returns the input value.
#' @param contact A character vector with contact information
#' @param compact A boolean which controls the format of the contact information. Default: FALSE
#' @returns A character vector
#' @keywords internal
#' @noRd
inputHelperContactInfo <- function(contact, compact = FALSE){

  if (contact == "statzh"){
    if (compact){
      contact <- statzh_contact_compact
    } else {
      contact <- statzh_contact
    }
  }

  return(contact)
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
#'
inputHelperOfficeHours <- function(openinghours){
  if (openinghours == "statzh"){
    openinghours <- statzh_openinghours
  }

  return(openinghours)
}


#' inputHelperHomepage()
#'
#' @description Replace default values "zh" and "statzh" with the homepage of
#'  the Statistics Office of Canton Zurich, otherwise returns the input value.
#' @note Converts input to 'hyperlink' object.
#' @param homepage A character string
#' @returns A character string or a 'hyperlink' object
#' @keywords internal
#' @noRd
#'
inputHelperHomepage <- function(homepage){
  homepage <- sub("statzh", statzh_homepage, homepage)
  # class(homepage) <- 'hyperlink'

  return(homepage)
}


#' înputHelperDateCreated()
#'
#' @description Returns current date as a string with format specified by date_format
#' @param prefix A Prefix to prepend to the date. Default to NULL
#' @param date_format A character string for the date_format. Default: "%d.%m.%Y".
#' @returns A character vector
#' @keywords internal
#' @seealso format
#' @noRd
inputHelperDateCreated <- function(prefix = NULL, date_format = "%d.%m.%Y"){
  paste0(prefix, format(Sys.Date(), format = date_format))
}


#' înputHelperOrderNumber()
#'
#' @description Returns current date as a string with format specified by date_format
#' @param prefix A Prefix to prepend to the date. Default to NULL
#' @returns A character vector
#' @keywords internal
#' @seealso format
#' @noRd
inputHelperOrderNumber <- function(order_num, prefix = NULL){
  paste0(prefix, order_num)
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
inputHelperAuthorName <- function(author, prefix = NULL){
  if (author == "user"){
    sys_vals <- c(Sys.getenv("USERNAME"), Sys.getenv("USER"))
    author_name <- sys_vals[which(sys_vals != "")[1]]
    author <- stringr::str_sub(author_name, start = 6, end = 7)
  }

  return(paste0(prefix, author))
}



#' excelIndexToRowCol()
#'
#' @description Function extracts initials from global environment if
#'  input is 'user', otherwise passes on the input.
#' @param author Name of author
#' @param prefix A character string to prepend to the result
#' @returns A character string with the username
#' @keywords internal
#' @noRd
#'
excelIndexToRowCol <- function(index){

  splitIndex <- function(x, split = "") unlist(strsplit(x, split))

  excelColumnLetterToNumeric <- function(x){
    chars <- splitIndex(x, "")
    offsets <- 26^(length(chars):1 - 1)
    sum(offsets * match(chars, LETTERS))
  }

  # Return extent of occupied region
  if (length(index) > 1){
    extents <- lapply(index, excelIndexToRowCol)
    extents <- unique(do.call(rbind,lapply(extents, stack)))
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

#' getNamedRegionExtent()
#'
#' @description Function extracts initials from global environment if
#'  input is 'user', otherwise passes on the input.
#' @param author Name of author
#' @param prefix A character string to prepend to the result
#' @returns A character string with the username
#' @keywords internal
#' @noRd
#'
getNamedRegionExtent <- function(wb, sheet, name = NULL){
  named_regions <- openxlsx::getNamedRegions(wb)
  named_regions_attr <- as.data.frame(attributes(named_regions))

  if (!is.null(name)){
    match(name, named_regions)
    ind <- which(name == named_regions)
  } else {
    ind <- 1:length(named_regions)
  }

  # print(named_regions)
  if (length(ind) != 0){
    return(excelIndexToRowCol(named_regions_attr[ind, "position"]))
  }
}

#' getNamedRegionExtent()
#'
#' @description Function extracts initials from global environment if
#'  input is 'user', otherwise passes on the input.
#' @param author Name of author
#' @param prefix A character string to prepend to the result
#' @returns A character string with the username
#' @keywords internal
#' @noRd
#'
getNamedRegionFirstRow <- function(wb, sheet, name = NULL){
  region_extent <- getNamedRegionExtent(wb, sheet, name)
  return(min(region_extent[["row"]]))
}

#' getNamedRegionExtent()
#'
#' @description Function extracts initials from global environment if
#'  input is 'user', otherwise passes on the input.
#' @param author Name of author
#' @param prefix A character string to prepend to the result
#' @returns A character string with the username
#' @keywords internal
#' @noRd
#'
getNamedRegionLastRow <- function(wb, sheet, name = NULL){
  region_extent <- getNamedRegionExtent(wb, sheet, name)
  return(max(region_extent[["row"]]))
}

