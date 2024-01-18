confdir <- "test_env"
output_dir <- "test_out"

config_file_path <- paste0(confdir, "/statR_profile.csv")
output_file_path <- file.path(output_dir, "testfile.xlsx")

dir.create(confdir, recursive = TRUE, showWarnings = FALSE)
dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)
