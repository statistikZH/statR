
#' setStatROpts()
#' @description Sets options
#' @param opt_list A named list
#' @export
setStatROpts <- function(opt_list) {
  opt_list <- opt_list[which(names(opt_list) %in% names(options()))]
  options(opt_list)
}

#' getUserConfigs()
#' @description Returns a character vector of all user configurations
#' @export
getUserConfigs <- function() {
  config_path <- system.file("extdata/config/", package = "statR")
  list.files(config_path)
}


#' readUserConfig()
#' @description Reads a user config. By default loads the default config
#' @param name The name of the configuration
#' @export
readUserConfig <- function(name = "default") {

  # config_path <- system.file("extdata/config/", package = "statR")
  # config_file <- paste0(config_path, name)

  config_path <- system.file("extdata/config", package = "statR")
  config_file <- paste0(c(config_path, name), collapse = "/")


  if (file.exists(config_file)) {
    config <- yaml::read_yaml(config_file)
    config[["statR_config_name"]] <- name
    options(config)
  }
}


#' writeUserConfig()
#' @description Writes a user config into a YAML file
#' @param name Name of the configuration
#' @param config_list List of options set by user
#' @examples
#' \dontrun{
#' # statzh config list
#' config_list <- list(
#'   statR_config_name = "statzh",
#'   statR_organization = "Statistisches Amt des Kantons Zürich",
#'   statR_name = "Datashop",
#'   statR_phone =  "+41 43 259 75 00",
#'   statR_email = "datashop@statistik.zh.ch",
#'   statR_homepage = "http://www.statistik.zh.ch",
#'   statR_openinghours = c("Bürozeiten",
#'                          "Montag bis Freitag",
#'                          "09:00 bis 12:00",
#'                          "13:00 bis 16:00"),
#'   statR_logo = "statzh",
#'   statR_source = "Statistisches Amt des Kantons Zürich"
#' )
#'
#' # Write config list to disk
#' writeUserConfig("statzh", config_list)
#' }
#' @export
writeUserConfig <- function(name, config_list) {
  # config_path <- system.file("extdata/config/", package = "statR")
  config_path <- system.file("extdata/config", package = "statR")
  config_file <- paste0(c(config_path, name), collapse = "/")


  config_list[["statR_config_name"]] <- name
  # yaml::write_yaml(config_list, paste0(config_path, name))
  yaml::write_yaml(config_list, config_file)
}


#' getActiveConfigName()
#' @description Returns the name of the active config
#' @export
getActiveConfigName <- function() {
  getOption("statR_config_name")
}


#' setActiveConfig()
#' @description Sets a configuration to active and optionally makes it
#'   persistent
#' @param name Name of the configuration
#' @param persistent Whether to load the configuration by default on next
#'   startup
#' @export
setActiveConfig <- function(name, persistent = FALSE) {
  # config_path <- system.file("extdata/config/", package = "statR")
  # config_file <- paste0(config_path, name)
  config_path <- system.file("extdata/config", package = "statR")
  config_file <- paste0(c(config_path, name), collapse = "/")



  if (getActiveConfigName() != name & file.exists(config_file)) {
    config <- yaml::read_yaml(config_file)

    if (persistent) {
      writeUserConfig("persistent", config)
    }

    options(config)
  }
}
