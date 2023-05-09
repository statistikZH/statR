#' Functions for checking and transforming user input
#'
#' @description
#' Helper functions not intended to be called directly by users.
#' @keywords internal
#'
check_sheetname <- function(sheetname){

  if (nchar(sheetname) > 31){
    warning("sheetname is cut to 31 characters (limit imposed by MS-Excel)")
    sheetname <- paste(substr(sheetname, 0, 31))
  }

  return(sheetname)
}

prep_filename <- function(filename){

  if (!grepl(".xlsx", filename)){
    filename <- paste0(filename, ".xlsx")
  }

  return(filename)
}

prep_source <- function(source, prefix = "Quelle:"){

  source <- sub("statzh", "Statistisches Amt des Kantons Z\u00fcrich", source)
  return(paste("Quelle:", paste0(source, collapse = ";")))
}


prep_metadata <- function(metadata){

  return(paste0("Metadaten: ", paste0(metadata, collapse = "; ")))
}

prep_logo <- function(logo){

  if(logo == "statzh") {
    logo <- paste0(path.package("statR"),"/extdata/Stempel_STAT-01.png")

  } else if(logo == "zh"){
    logo <- paste0(path.package("statR"),"/extdata/Stempel_Kanton_ZH.png")
  }

  return(logo)
}

prep_contact <- function(contact){
  if (contact == "statzh"){
    contact <- c("Datashop", "Tel.:  +41 43 259 75 00",
                 "datashop@statistik.zh.ch")
  } else if (length(contact) > 3){
    warning("More than 3 elements in contactdetails, may overlap with other elements.")
  }

  return(contact)
}

prep_openinghours <- function(openinghours){
  if (openinghours == "statzh"){
    openinghours <- c("B\u00fcrozeiten",
                      "Montag bis Freitag",
                      "09:00 bis 12:00",
                      "13:00 bis 16:00")
  }

  return(openinghours)
}

prep_homepage <- function(homepage){
  if (homepage == "statzh"){
    homepage <- "http://www.statistik.zh.ch"
  }

  class(homepage) <- 'hyperlink'
  return(homepage)
}

prep_creationdate <- function(prefix = "Erstellt am:", date_format = "%d.%m.%Y"){
  return(paste(prefix, format(Sys.Date(), format = date_format)))
}

prep_orderid <- function(order_id, prefix = "Auftragsnr.:"){
  return(paste(prefix, order_id))
}
