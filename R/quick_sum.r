# quick_sum: Funktion für einfache (gruppierte) Standradauswertungen

#' quick_sum
#'
#' Function to create a formated single-worksheet XLSX automatically
#' @param df dataframe
#' @param var variables to be aggregated
#' @param ... grouping variables
#' @param stats Which evaluation statistics should be displayed (all, base, mean)?
#' @param protect Apply data protection T/F
#' @keywords quick_sum
#' @importFrom dplyr ungroup summarize group_by n
#' @importFrom stats median quantile sd
#' @export
#' @examples
#'
#' quick_sum(df=mtcars, var=cyl, mpg, vs, stats="all", protect=FALSE)


quick_sum <- function(df, var, ..., stats="all", protect=FALSE) {
  grps <- rlang::quos(...)
  var <- rlang::enquo(var)

  mean_name <- paste0("mean_", rlang::quo_name(var))
  sd_name <- paste0("sd_", rlang::quo_name(var))
  q10_name <- paste0("q10_", rlang::quo_name(var))
  q25_name <- paste0("q25_", rlang::quo_name(var))
  med_name <- paste0("med_", rlang::quo_name(var))
  q75_name <- paste0("q75_", rlang::quo_name(var))
  q90_name <- paste0("q90_", rlang::quo_name(var))

  if(protect==TRUE){
    n1 <- 3
    n2 <- 5
  }else{
    n1 <- 0
    n2 <- 0
  }

  if(stats%in%c("base")){df %>%
    dplyr::group_by(!!!grps) %>%
    dplyr::summarise(
      Anzahl=dplyr::n(),
      !!mean_name := ifelse(Anzahl<=n1,NA,mean(!!var, na.rm = T)),
      !!q25_name := ifelse(Anzahl<=n2,NA,quantile(!!var, probs=0.25, na.rm = T)),
      !!med_name := ifelse(Anzahl<=n1,NA,median(!!var, probs=0.5, na.rm = T)),
      !!q75_name := ifelse(Anzahl<=n2,NA,quantile(!!var, probs=0.75, na.rm = T))
    )%>%
    dplyr::ungroup()}
  else if(stats%in%c("all")){df %>%
      dplyr::group_by(!!!grps) %>%
      dplyr::summarise(
        Anzahl=dplyr::n(),
        !!mean_name := ifelse(Anzahl<=n1,NA,mean(!!var, na.rm = T)),
        !!sd_name := ifelse(Anzahl<=n2,NA,sd(!!var, na.rm = T)),
        !!q10_name := ifelse(Anzahl<=n2,NA,quantile(!!var, probs=0.1, na.rm = T)),
        !!q25_name := ifelse(Anzahl<=n2,NA,quantile(!!var, probs=0.25, na.rm = T)),
        !!med_name := ifelse(Anzahl<=n1,NA,median(!!var, probs=0.5, na.rm = T)),
        !!q75_name := ifelse(Anzahl<=n2,NA,quantile(!!var, probs=0.75, na.rm = T)),
        !!q90_name := ifelse(Anzahl<=n2,NA,quantile(!!var, probs=0.9, na.rm = T))
      )%>%
      dplyr::ungroup()}
  else if(stats%in%c("mean")){df %>%
      dplyr::group_by(!!!grps) %>%
      dplyr::summarise(
        Anzahl=dplyr::n(),
        !!mean_name := ifelse(Anzahl<=n1,NA,mean(!!var, na.rm = T)),
        !!sd_name := ifelse(Anzahl<=n2,NA,sd(!!var, na.rm = T))
      )%>%
      dplyr::ungroup()}
  else {print("Please choose one of the following stats arguments: all, base, mean")}
}
