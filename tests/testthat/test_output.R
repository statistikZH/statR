# Checks for file-output if quickXLSX runs
testthat::test_that(
  "Check if minimal quickXLSX example yields a file",
  {
    data(mtcars)
    path <- tempfile(fileext = ".xlsx")
    quickXLSX(data = mtcars, file = path)
    testthat::expect_true(file.exists(path))
  }
)

testthat::test_that(
  "minimal aXLSX example yields a file",
  {
    testthat::skip("This test hasn't been written yet")
  }
)

testthat::test_that(
  "minimal datasetsXLSX example yields a file",
  {
    testthat::skip("This test hasn't been written yet")
  }
)

testthat::test_that(
  "Complex datasetsXLSX call produces right number of plots and sheets",
  {
    testthat::skip("This test hasn't been written yet")
  }
)

testthat::test_that(
  "Adding attributes works as expected",
  {
    testthat::skip("This test hasn't been written yet")
  }
)
