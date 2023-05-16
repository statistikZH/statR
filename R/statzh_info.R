#################### STAT ZH Farbpaletten
#' statzh_info
#'
#' Commonly used info for the Statistical Office of Canton Zurich
#'
#' @format A \code{list}.
#' @noRd
#'
statzh_info <- {
  # Contact info
  #---------------
  statzh_name <- "Statistisches Amt des Kantons Z\u00fcrich"
  statzh_department <- "Datashop"
  statzh_phone <- "Tel.:  +41 43 259 75 00"
  statzh_email <- "datashop@statistik.zh.ch"
  statzh_homepage <- "http://www.statistik.zh.ch"

  statzh_contact <- c(statzh_department, statzh_phone, statzh_email)
  statzh_contact_compact <- c(paste(statzh_department, statzh_phone, sep = ", "),
                              statzh_email, statzh_homepage)
  statzh_openinghours <- c("B\u00fcrozeiten", "Montag bis Freitag", "09:00 bis 12:00", "13:00 bis 16:00")

  # Logo filenames (located in /extdata/ folder)
  #-------------
  statzh_logo <- "Stempel_STAT-01.png"
  zh_logo <- "Stempel_Kanton_ZH.png"

}
