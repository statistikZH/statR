test_that(
  "config file correctly initiated",
  {
    initUserConfigStore(store_path = confdir)

    testthat::expect_true(file.exists(config_file_path))
    config_file <- read.csv(config_file_path)

    testthat::expect_equal(names(config_file),
                           c("config_name", "config_path"))
    testthat::expect_equal(nrow(config_file), 1)
    testthat::expect_equal(config_file$config_name, "default")
    testthat::expect_equal(
      config_file$config_path,
      system.file("extdata/config/default.yaml", package = "statR"))
    testthat::expect_equal(getOption("statR_email"),
                           "datashop@statistik.zh.ch")

    cleanup_test()
  })

test_that(
  "default value is not set twice if already existing",
  {
    initUserConfigStore(store_path = confdir)
    addUserConfig(store_path = confdir)
    out <- addUserConfig(store_path = confdir)

    config_file <- read.csv(config_file_path)

    testthat::expect_equal(nrow(config_file), 1)
    testthat::expect_equal(out, "Alles bereit")

    cleanup_test()
  })


test_that(
  "second entry is added correctly",
  {
    initUserConfigStore(store_path = confdir)

    addUserConfig("test",
                  testthat::test_path("example_userconf", "test"),
                  store_path = confdir)

    config_file <- read.csv(config_file_path)

    testthat::expect_equal(nrow(config_file), 2)

    testthat::expect_error(
      addUserConfig("test",
                    testthat::test_path("example_userconf", "blabla"),
                    store_path = confdir),
    "Kein config file gefunden bei Pfad ")

  cleanup_test()
  })


test_that(
  "default can be changed and is not overwritten again",
  {
    initUserConfigStore(store_path = confdir)
    updateUserConfig("default",
                     testthat::test_path("example_userconf", "test"),
                     store_path = confdir)

    addUserConfig(store_path = confdir)
    config_file <- read.csv(config_file_path)

    testthat::expect_equal(nrow(config_file), 1)
    testthat::expect_equal(config_file$config_name, "default")

    testthat::expect_equal(
      config_file$config_path,
      testthat::test_path("example_userconf", "test"))

    cleanup_test()
  })


test_that(
  "the default entry can not be removed",
  {
    initUserConfigStore(store_path = confdir)
    testthat::expect_error(removeUserConfig("default"),
                           "Der Default-Wert kann nicht")
    cleanup_test()
  })


test_that(
  "an additional entry can be deleted",
  {
    initUserConfigStore(store_path = confdir)
    addUserConfig("test",
                  testthat::test_path("example_userconf", "test"),
                  store_path = confdir)

    config_file <- read.csv(config_file_path)
    testthat::expect_equal(nrow(config_file), 2)
    removeUserConfig("test", store_path = confdir)

    config_file <- read.csv(config_file_path)
    testthat::expect_equal(nrow(config_file), 1)

    cleanup_test()
  })


test_that(
  "an entry can not be overwritten with the addUserConfig-function",
  {
    initUserConfigStore(store_path = confdir)
    addUserConfig("test",
                  testthat::test_path("example_userconf", "test"),
                  store_path = confdir)


    testthat::expect_error(
      addUserConfig("test",
                    testthat::test_path("example_userconf", "test"),
                    store_path = confdir),
      "Diese Konfiguration existiert bereits")

    testthat::expect_error(
      addUserConfig("test",
                    testthat::test_path("example_userconf", "testone"),
                    store_path = confdir),
      "Der Konfigurationsname:")

    cleanup_test()
  })


