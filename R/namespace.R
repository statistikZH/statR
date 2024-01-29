#' @keywords internal
.onLoad <- function(libname, pkgname) {
  message(paste(list.files(system.file("extdata/config", package = "statR"))))

  initUserConfigStore()
}

#' @keywords internal
.onUnload <- function(libname, pkgname) {

}
