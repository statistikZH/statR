# TODO: breaks down. After some remove/add operations, the store is now
# of dim = c(8,1) instead of c(2,2). Somewhere along the path, an rbind
# or similar operation is going awry.
#' Intended to be ran once after setup
#'
#' @param path the path to the file which will store names and paths of
#'   different statR configs
initUserConfigStore <- function(path = "~/.config/R/statR") {

  if (!dir.exists(path)) {
    dir.create(path, recursive = TRUE)
  }

  store_file <- file.path(path, "statR_profile.csv")

  if (!file.exists(store_file)) {
    file.create(store_file)
    writeLines("config_name,config_path", store_file)

  }
}

readUserConfigStore <- function(path = "~/.config/R/statR") {
  store_file <- file.path(path, "statR_profile.csv")
  read.table(store_file, header = TRUE, sep = ",")
}


#' Add a new configuration using a name and path
addUserConfig <- function(name = "default", path = NULL,
                          store_path = "~/.config/R/statR") {

  store_file <- file.path(store_path, "statR_profile.csv")

  if(!file.exists(store_file)){
    initUserConfigStore(store_path)
  }

  configs <- readUserConfigStore(store_path)
  if (is.null(path) && name != "default") {
    stop("Must provide a path")
  }

  if (!is.null(path) && !file.exists(path)) {
    stop("No config file found at ", path)
  }

  if (name == "default" && is.null(path)) {
    path <- system.file("extdata/config/default", package = "statR")
  }


  if (name %in% configs$config_name &&
      rstudioapi::showQuestion("Overwrite config?",
                               "Overwrite the configuration file?")) {
      removeUserConfig(name, store_path)
  }

  out <- rbind(configs, data.frame(config_name = name, config_path = path))

  write.table(out, store_file, row.names = FALSE, sep = ",")
}


#' Remove a configuration using a name
#'
#'
removeUserConfig <- function(name, store_path = "~/.config/R/statR") {
  store_file <- file.path(store_path, "statR_profile.csv")
  configs <- readUserConfigStore(store_path)
  write.table(subset(configs, configs$config_name != name),
              store_file, row.names = FALSE, sep = ",")

  # When default config is deleted, replace it with the package default
  if (name == "default") {
    addUserConfig(store_path = store_path)
  }
}

readUserConfig <- function(name = "default", store_path = "~/.config/R/statR") {
  all_configs <- readUserConfigStore(store_path)
  path <- subset(all_configs, name == all_configs$config_name)$config_path

  if (!file.exists(path)) {
    stop("Configuration ", name, " not found.")
  }

  yaml::read_yaml(path)
}

loadUserConfig <- function(name = "default", store_path = "~/.config/R/statR") {
  options(readUserConfig(name, store_path))
}
