#' @keywords internal
.onLoad <- function(libname, pkgname) {

  # TODO: do we need to give this a non-default path?
  #       Note: this will set up the profile store on first startup,
  #       adding the default profile from the package
  if ((path <- Sys.getenv("STATR_CONFIG_PATH")) == "") {
    path <- c(
      getOption("statR_config_path"),
      "~/.config/R/statR")[1]
  }


  initUserConfigStore(path)
}

#' @keywords internal
.onUnload <- function(libname, pkgname) {

}
