#-----------------
#' check_sheetname()
#'
#' @description Function which truncates a sheet name to 31 characters
#'
#' @details MS Excel imposes a character limit of 31 characters for sheetnames.
#'  By default, this function truncates sheetnames accordingly and raises a
#'  warning. If shorten is set to FALSE, an error is raised instead.
#'
#' @param sheetname A character string with the name for an XLSX worksheet
#'
#' @param shorten A boolean whether to return shortened string or raise an error.
#'  Default: TRUE
#'
#' @returns A character string
#' @keywords internal
#' @noRd
#'
check_sheetname <- function(sheetname, shorten = TRUE){

  short_name <- paste(substr(sheetname, 0, 31))

  if (short_name != sheetname){
    if (shorten) {
      warning("sheetname is truncated to 31 characters")
    } else {
      stop("sheetname exceeds character limit of MS-Excel)")
    }
  }

  return(short_name)
}

#-----------------
#' writeDataIf()
#'
#' @description A wrapper around openxlsx::writeData conditioned on the result of
#'  condition_fun(x). By default, `condition_fun <- function(x) !is.null(x)` is
#'  used to verify that the input is not NULL.
#'
#' @param wb Workbook object
#'
#' @param sheet Sheetname
#'
#' @param x Input to be written to worksheet
#'
#' @param xy Position where input should be written to
#'
#' @param condition_fun A function which returns a boolean and takes x as input.
#'  Default: function(x) !is.null(x)
#'
#' @keywords internal
#' @noRd
#'
writeDataIf <- function(wb, sheet, x, xy, condition_fun = function(x) !is.null(x)){

  if (condition_fun(x)){
    openxlsx::writeData(wb, sheet, x, xy = xy)
  } else {
    message("Skipped write because x is NULL")
  }
}

#-----------------
#' prep_filename()
#'
#' Ensure that file extension is added to filename.
#' @param filename A character string with the filename
#' @param extension A character string with the file extension. Defaults to ".xlsx"
#' @returns A character string
#' @keywords internal
#' @noRd
#'
prep_filename <- function(filename, extension = ".xlsx"){

  if (!grepl(".xlsx", filename)){
    filename <- paste0(filename, extension)
  }

  return(filename)
}

#-----------------
#' prep_source()
#'
#' @description Concatenate vector of sources into a formatted string
#'
#' @param source A character vector with source information
#'
#' @param prefix A character string with the prefix. Default: "Quelle: "
#'
#' @param sep A character string with the separator for sources. Default: ";"
#'
#' @returns A character string
#'
#' @keywords internal
#' @noRd
#'
prep_source <- function(source){

  # Replace "statzh" entries
  source <- sub("statzh", statzh_name, source)

  cat_source <- paste0(source, collapse = sep)

  if (!is.null(prefix)){
    cat_source <- paste0(prefix, cat_source)
  }

  return(cat_source)
}

#-------------------
#' prep_metadata()
#'
#' @description Concatenate metadata into a formatted string.
#'
#' @param metadata A character vector with metadata information
#'
#' @param extension A character string with the prefix. Default: "Metadaten: ".
#'
#' @param sep A character string with the separator for metadata. Default: ";"
#'
#' @returns A character string
#'
#' @keywords internal
#' @noRd
#'
prep_metadata <- function(metadata, prefix = "Metadaten: ", sep = ";"){

  cat_metadata <- paste0(metadata, collapse = sep)

  if (!is.null(prefix)){
    cat_metadata <- paste0(prefix, cat_metadata)
  }

  return(cat_metadata)
}

#-----------------
#' prep_logo()
#'
#' @description Replace default values "zh" and "statzh" with the file path to
#'  the respective logo (included in /extdata), otherwise returns the input value.
#'
#' @param logo A character string with the file path
#'
#' @returns A character string
#' @keywords internal
#' @noRd
#'
prep_logo <- function(logo){

  if (logo == "statzh"){
    logo <- statzh_logo
  } else if (logo == "zh"){
    logo <- zh_logo
  }

  return(logo)
}

#-----------------
#' prep_contact()
#'
#' Replace default values "zh" and "statzh" with the contact information of
#' the Statistics Office of Canton Zurich, otherwise returns the input value.
#' @note
#' Raises a warning if the input has more than 3 elements.
#' @param contact A character vector with contact information
#' @param compact A boolean which controls the format of the contact information. Default: FALSE
#' @returns A character vector
#' @keywords internal
#' @noRd
#'
prep_contact <- function(contact, compact = FALSE){

  if (contact == "statzh"){
    contact <- ifelse(compact, statzh_contact_compact, statzh_contact)
  } else if (length(contact) > 3){
    warning("More than 3 elements in contactdetails, may overlap with other elements.")
  }

  return(contact)
}

#-----------------
#' prep_openinghours()
#'
#' @description Replace default values "zh" and "statzh" with the opening hours
#'  of the Statistics Office of Canton Zurich, otherwise returns
#'  the input value.
#'
#' @param openinghours A character vector with opening hours.
#'
#' @returns A character vector
#' @keywords internal
#' @noRd
#'
prep_openinghours <- function(openinghours){

  if (openinghours == "statzh"){
    openinghours <- statzh_openinghours
  }

  return(openinghours)
}

#-----------------
#' prep_homepage()
#'
#' @description Replace default values "zh" and "statzh" with the homepage of
#'  the Statistics Office of Canton Zurich, otherwise returns the input value.
#'
#' @note Converts input to 'hyperlink' object if as_hyperlink is TRUE (default).
#'
#' @param homepage A character string
#'
#' @param as_hyperlink A boolean, controls whether input is converted to
#'  'hyperlink'. Default: TRUE
#'
#' @returns A character string or a 'hyperlink' object depending on the value of
#'  as_hyperlink.
#' @keywords internal
#' @noRd
#'
prep_homepage <- function(homepage, as_hyperlink = TRUE){

  if (homepage == "statzh"){
    homepage <- statzh_homepage
  }

  if (as_hyperlink){
    class(homepage) <- 'hyperlink'
  }

  return(homepage)
}

#-----------------
#' prep_creationdate()
#'
#' @param prefix A character string, defaults to "Erstellt am:"
#' @param date_format A character string for the date_format. Default: "%d.%m.%Y".
#' @returns A character vector
#' @keywords internal
#' @seealso format
#' @noRd
#'
prep_creationdate <- function(prefix = "Erstellt am:", date_format = "%d.%m.%Y"){

  creation_date <- format(Sys.Date(), format = date_format)
  return(paste(prefix, creation_date))
}

#-----------------
#' prep_orderid()
#'
#' @param order_id An integer
#' @param prefix A character string, defaults to "Auftragsnr.:"
#' @returns A character vector
#' @keywords internal
#' @noRd
#'
prep_orderid <- function(order_id, prefix = "Auftragsnr.:"){

  return(paste(prefix, order_id))
}


#-----------------
#' prep_username()
#'
#' @param author Name of author
#' @returns A character string with the username
#' @keywords internal
#' @noRd
#'
prep_username <- function(author){

  if (author != "user"){
    return(author)
  }

  if(Sys.getenv("USERNAME") != ""){
    author <- stringr::str_sub(Sys.getenv("USERNAME"), start = 6, end = 7)

  } else {
    author <- stringr::str_sub(Sys.getenv("USER"), start = 6, end = 7)
  }

  return(author)
}
