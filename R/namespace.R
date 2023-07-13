#' .onLoad()
#'
#' @keywords internal
.onLoad <- function(libname, pkgname) {

  # Set default values for some format options
  defaults <- list(
    date_format = "%d.%m.%Y",
    statR_prefix_date = "Aktualisiert am:",
    statR_prefix_author = "durch:",
    statR_prefix_phone = "Tel.",
    statR_prefix_source = "Quelle:",
    statR_prefix_metadata = "Hinweise:",
    statR_prefix_order_id = "Auftragsnr.:",
    statR_collapse = ";",
    statR_toc_title = "Inhalt"
  )

  options(defaults)

  if ("persistent" %in% getUserConfigs()) {
    readUserConfig("persistent")

  } else {
    readUserConfig("default")
  }
}
