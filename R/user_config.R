#' Initialisiert die Excel-Header-Konfiguration
#'
#' Die Header-Konfiguration wird ueber globale Options geloest, welche in
#' einem Yaml file definiert werden. Per default wird die Konfiguration des
#' Packages angezogen. Es besteht die Moeglichkeit, eine eigene Konfiguration
#' zu hinterlegen. Siehe dazu die Funktion: addUserConfig()
#' @param store_path Pfad unter welchem das Konfigurationsfile fuer das registrieren
#'  von Header-Config-Yaml-Files liegen soll
#' @export
initUserConfigStore <- function(store_path = "~/.config/R/statR") {

  if (!dir.exists(store_path)) {
    dir.create(store_path, recursive = TRUE)
  }

  store_file <- file.path(store_path, "statR_profile.csv")

  if (!file.exists(store_file)) {
    file.create(store_file)
    writeLines("config_name,config_path", store_file)

  }


  addUserConfig(store_path = store_path)

}


#' Liest das Konfigurationsfile in welchem die Pfade zum Header-Config file
#' hinterlegt sind
#'
#' @inheritParams initUserConfigStore
readUserConfigStore <- function(store_path = "~/.config/R/statR") {
  store_file <- file.path(store_path, "statR_profile.csv")
  read.table(store_file, header = TRUE, sep = ",")
}


#' Registrieren eines neuen Header-Config-Yaml Files
#'
#' @param name unter welchem namen soll die Header-Konfiguration abrufbar sein.
#' @param path Pfad zum Header-Konfigurations-Yaml-File
#' @inheritParams initUserConfigStore
#' @export
addUserConfig <- function(name = "default", path = NULL,
                          store_path = "~/.config/R/statR") {

  store_file <- file.path(store_path, "statR_profile.csv")

  if(!file.exists(store_file)){
    initUserConfigStore(store_path)
  }

  configs <- readUserConfigStore(store_path)
  if (is.null(path) && name != "default") {
    stop("Kein Pfad zu config file angegeben")
  }

  if (!is.null(path) && !file.exists(path)) {
    stop("Kein config file gefunden bei Pfad ", path)
  }

  if (name == "default" && is.null(path)) {
    path <- system.file("extdata/config/default.yaml", package = "statR")
  }

  if (name == "default" & "default" %in% configs$config_name){
    return("Alles bereit")
  }

  if (name %in% configs$config_name){

    if (configs[configs$config_name == name, "config_path"] == path){
      stop("Diese Konfiguration existiert bereits! Verwende die ",
           "updateUserConfig()-Funktion um den Pfad zu aendern.")

    } else {
      stop("Der Konfigurationsname: ",
           configs[configs$config_name == name, "config_name"],
           " existiert bereits. Setze einen neuen Pfad mit der ",
           "updateUserConfig()-Funktion")
    }
  }

  out <- rbind(configs, data.frame(config_name = name, config_path = path))
  write.table(out, store_file, row.names = FALSE, sep = ",")
}


#' Anpassen eines Header-Config-Yaml-Files-Pfades
#'
#' @param name Zu welchem Eintrag moechtest du den Pfad anpassen
#' @param path Pfad zum Header-Konfigurations-Yaml-File
#' @inheritParams initUserConfigStore
#' @export
updateUserConfig <- function(name, path, store_path = "~/.config/R/statR"){
  if (!is.null(path) && !file.exists(path)) {
    stop("No config file found at ", path)
  }

  configs <- readUserConfigStore(store_path)

  configs[configs$config_name == name, "config_path"] <- path

  out <- configs

  store_file <- file.path(store_path, "statR_profile.csv")
  write.table(out, store_file, row.names = FALSE, sep = ",")
}

#' Loeschen eines Konfigurations-Eintrages
#'
#' @param name Welcher registrierte Eintrag soll geloescht werden
#' @inheritParams initUserConfigStore
#' @export
removeUserConfig <- function(name, store_path = "~/.config/R/statR") {

  if (name == "default") {
    stop("Der Default-Wert kann nicht geloescht werden. Wenn du den Pfad ",
         "anpassen moechtest, verwende die updateUserConfig()-Funktion")
  }

  store_file <- file.path(store_path, "statR_profile.csv")
  configs <- readUserConfigStore(store_path)
  write.table(subset(configs, configs$config_name != name),
              store_file, row.names = FALSE, sep = ",")
}


#' Liest das Header-Konfigurations-Yaml-File
#'
#' @param name welches file soll angezogen werden
#' @inheritParams initUserConfigStore
#' @export
readUserConfig <- function(name = "default", store_path = "~/.config/R/statR") {
  all_configs <- readUserConfigStore(store_path)
  path <- subset(all_configs, name == all_configs$config_name)$config_path

  if (!file.exists(path)) {
    stop("Header-Konfigurations-YAML-File: ", path, " existiert nicht.")
  }

  yaml::read_yaml(path)
}






get_user_config <- function(config, params_to_check){

  initUserConfigStore()

  user_config <- readUserConfig(config)

  out <- unlist(user_config, recursive = FALSE)

  names(out) <- gsub(".*\\.", "", names(out))

  config_name <- tail(paste0("statR_", substitute(params_to_check)), -1)


  user_config <- purrr::reduce2(params_to_check, config_name, ~ replace_by_parameter(..1, ..2, ..3), .init = out)

  options(user_config)


  if(!("statR_contactdetails" %in% names(user_config))){
    user_config$statR_contactdetails <- inputHelperContactInfo()

    options(user_config)
  }



}


replace_by_parameter <- function(yaml_file, parameter, config_param_name) {

  if (!is.na(parameter)) {
    if(config_param_name %in% names(yaml_file)){
      yaml_file[config_param_name] <- parameter
    } else {
      yaml_file$add <- parameter

      new_names <- c(head(names(yaml_file),-1), config_param_name)

      names(yaml_file) <- new_names
    }

  }

  return(yaml_file)
}


