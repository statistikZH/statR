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
  config_path <- system.file("extdata/config/", package = "statR")
  config_file <- paste0(config_path, name)

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
#' @export
writeUserConfig <- function(name, config_list) {
  config_path <- system.file("extdata/config/", package = "statR")
  config_list[["statR_config_name"]] <- name
  yaml::write_yaml(config_list, paste0(config_path, name))
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
  config_path <- system.file("extdata/config/", package = "statR")
  config_file <- paste0(config_path, name)

  if (getActiveConfigName() != name & file.exists(config_file)) {
    config <- yaml::read_yaml(config_file)

    if (persistent) {
      writeUserConfig("persistent", config)
    }

    options(config)
  }
}
