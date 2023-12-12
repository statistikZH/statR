
#' getUserConfigs()
#' @description Returns a character vector of all user configurations
#' @export
getUserConfigs <- function() {
  config_path <- system.file("extdata/config/", package = "statR")
  list.files(config_path)
}

#' exportUserConfig()
#' @description Reads a user config. By default loads the default config
#' @param name The name of the configuration
#' @export
exportUserConfig <- function(name = "default") {
  config_path <- system.file("extdata/config/", package = "statR")
  config_file <- file.path(config_path, name)

  if (file.exists(config_file)){
    return(yaml::read_yaml(config_file))
  }
}


#' readUserConfig()
#' @description Reads a user config. By default loads the default config
#' @param name The name of the configuration
#' @param persistent Whether to load the configuration by default on next
#' @export
readUserConfig <- function(name = "default", persistent = FALSE) {

  config_path <- system.file("extdata/config", package = "statR")
  config_file <- file.path(config_path, name)

  if (file.exists(config_file)) {
    config <- yaml::read_yaml(config_file)
    config[["statR_config_name"]] <- name

    if (persistent) {
      writeUserConfig("persistent", config)
    }

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
#' # or alternatively
#' config_list <- exportUserConfig(name = "statzh")
#'
#' # Modify list:
#' config_list[["statR_name"]] <- "Data Management"
#'
#' # Write config list to disk
#' writeUserConfig("statzh", config_list)
#' }
#' @export
writeUserConfig <- function(name, config_list) {
  config_path <- system.file("extdata/config", package = "statR")
  config_file <- file.path(config_path, name)

  config_list[["statR_config_name"]] <- name
  yaml::write_yaml(config_list, config_file)
}


#' getActiveConfigName()
#' @description Returns the name of the active config
#' @export
getActiveConfigName <- function() {
  getOption("statR_config_name")
}


#TODO Tool to export configs from old R installation
function(old_version) {
  current_path <- system.file("extdata/config", package = "statR")

  old_path <- stringr::str_replace(
    current_path, "([0-9\\.]+)/statR", paste0(4.2, "/statR")
  )

  configs <- list.files(old_path)

  file.copy(file.path(old_path, configs),
            file.path(new_path, configs),
            overwrite = TRUE)
}
