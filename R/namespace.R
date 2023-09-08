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
    statR_toc_title = "Inhalt",
    statR_default_title = "Title",
    statR_default_source = "Source",
    statR_default_metadata = NA,
    statR_default_grouplines = NA,
    statR_default_group_names = NA
  )

  options(defaults)

  if ("persistent" %in% getUserConfigs()) {
    readUserConfig("persistent")

  } else {
    readUserConfig("default")
  }
}
