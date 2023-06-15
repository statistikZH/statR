#' .onLoad()
#'
#' @keywords internal
.onLoad <- function(libname, pkgname) {
  defaults <- list(
    date_format = "%d.%m.%Y",
    prefix_date = "Aktualisiert am:",
    prefix_author = "durch:",
    prefix_source = "Quelle:",
    prefix_metadata = "Hinweise:",
    prefix_order_id = "Auftragsnr.:",
    collapse_source = ";",
    toc_title = "Inhalt"
  )

  options(defaults)
}
