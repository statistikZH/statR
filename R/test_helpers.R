cleanup_test <- function(test_env = "test_env") {
  unlink(test_env, recursive = TRUE)
  dir.create(test_env)
}
