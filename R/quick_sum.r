#' quick_sum()
#'
#' Function for simple (grouped) descriptive analyses.
#'
#' If stats is set to "all" (the default), the following descriptive statistics are computed per category of the grouping variable(s): mean, standard deviation,
#' 10\%-quantile, 25\%-quantile, median, 75\%-quantile, 90\%-quantile.
#' If stats is set to "base", the mean, 25\%-quantile, median, and 75\%-quantile are returned.
#' stats = "mean" computes the mean and the standard deviation.
#'
#' If protect is set to TRUE, the mean and the median is only computed for variables/variable categories with four or more observations. For <4 observations, the function returns NA for the respective variable/category. The other descriptives are only computed if there are at least 6 observations per variable (category).
#'
#' @param df dataframe
#' @param var variables to be aggregated
#' @param ... grouping variables
#' @param stats Which descriptive statistics should be computed? Can be "all", "base", or "mean". See details.
#' @param protect Apply data protection, can be TRUE or FALSE. See details.
#' @keywords quick_sum
#' @importFrom dplyr ungroup summarize group_by n
#' @importFrom stats median quantile sd
#' @importFrom rlang :=
#' @export
#' @examples
#'
#' quick_sum(df=mtcars, var=cyl, mpg, vs, stats="all", protect=FALSE)


quick_sum <- function(df, var, ..., stats = "all", protect = FALSE){
  if (!stats %in% c("mean", "base", "all")){
    stop("Please choose one of the following stats arguments: all, base, mean")
  }

  grps <- rlang::quos(...)
  var <- rlang::enquo(var)

  mean_name <- paste0("mean_", rlang::quo_name(var))
  sd_name <- paste0("sd_", rlang::quo_name(var))
  q10_name <- paste0("q10_", rlang::quo_name(var))
  q25_name <- paste0("q25_", rlang::quo_name(var))
  med_name <- paste0("med_", rlang::quo_name(var))
  q75_name <- paste0("q75_", rlang::quo_name(var))
  q90_name <- paste0("q90_", rlang::quo_name(var))

  # Set lower limit for number of samples
  n1 <- ifelse(protect, 3, 0)
  n2 <- ifelse(protect, 5, 0)

  if(stats %in% c("base")){
    df %>%
    dplyr::group_by(!!!grps) %>%
    dplyr::summarise(
      Anzahl=dplyr::n(),
      !!mean_name := ifelse(Anzahl <= n1, NA, mean(!!var, na.rm = T)),
      !!q25_name := ifelse(Anzahl <=n2, NA, quantile(!!var, probs = 0.25, na.rm = T)),
      !!med_name := ifelse(Anzahl<=n1,NA,median(!!var, probs=0.5, na.rm = T)),
      !!q75_name := ifelse(Anzahl<=n2,NA,quantile(!!var, probs=0.75, na.rm = T))
    )%>%
    dplyr::ungroup()

  } else if (stats %in% c("all")){
    df %>%
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
      dplyr::ungroup()

  } else if (stats %in% c("mean")){
    df %>%
      dplyr::group_by(!!!grps) %>%
      dplyr::summarise(
        Anzahl=dplyr::n(),
        !!mean_name := ifelse(Anzahl<=n1,NA,mean(!!var, na.rm = T)),
        !!sd_name := ifelse(Anzahl<=n2,NA,sd(!!var, na.rm = T))
      )%>%
      dplyr::ungroup()

  }
}
