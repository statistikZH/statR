#' my_sum: Funktion für einfache (gruppierte) Standradauswertungen
#' my_sum
#'
#' Function to create a formated single-worksheet XLSX automatically
#' @param df dataframe
#' @param var variables to be aggregated
#' @param ... Grouping Variables
#' @keywords my_sum
#' @export
#' @examples
#'
#' my_sum(df, var, grp1, grp2, ...)
#' df -> data.frame
#' var -> Auswertungsvariable
#' grp1 -> Groupingvariable 1 (optional)
#' grp2 -> Groupingvariable 2 (optional)
#'
#'


my_sum <- function(df, var, ...) {

  dplyr::group_by <- quos(...)
  var <- enquo(var)
  mean_name <- paste0("mean_", quo_name(var))
  sd_name <- paste0("sd_", quo_name(var))
  #q10_name <- paste0("q10_", quo_name(var))
  q25_name <- paste0("q25_", quo_name(var))
  med_name <- paste0("med_", quo_name(var))
  q75_name <- paste0("q75_", quo_name(var))
 #q90_name <- paste0("q90_", quo_name(var))

  df %>%
    group_by(!!!group_by) %>%
    summarise(
      Anzahl=n(),
      !!mean_name := ifelse(Anzahl<=3,NA,mean(!!var, na.rm = T)),
      !!sd_name := ifelse(Anzahl<=5,NA,sd(!!var, na.rm = T)),
      #!!q10_name := ifelse(Anzahl<=5,NA,quantile(!!var, probs=0.1, na.rm = T)),
      !!q25_name := ifelse(Anzahl<=5,NA,quantile(!!var, probs=0.25, na.rm = T)),
      !!med_name := ifelse(Anzahl<=3,NA,median(!!var, probs=0.5, na.rm = T)),
      !!q75_name := ifelse(Anzahl<=5,NA,quantile(!!var, probs=0.75, na.rm = T))
      #!!q90_name := ifelse(Anzahl<=5,NA,quantile(!!var, probs=0.9, na.rm = T))
    ) %>%
    ungroup()
}
