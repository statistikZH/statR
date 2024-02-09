# Checks for file-output if quickXLSX runs
testthat::test_that(
  "minimale Aufrufe erzeugen .xlsx Dateien",
  {

    testthat::expect_no_error(
      quickXLSX(data = mtcars, file = output_file_path, logo = "fixtures/logo.png"))
    testthat::expect_no_error(
      aXLSX(data = mtcars, file = output_file_path, logo = "fixtures/logo.png"))
    testthat::expect_no_error(
      splitXLSX(data = mtcars, file = output_file_path, sheetvar = "cyl",
                logo = "fixtures/logo.png"))

    cleanup_test("test_out")
  }
)

# Test, ob Grafiken eingebunden werden koennen. Namentlich, ob die Funktion
# durchlaeuft, und ob die Bilddatei auch tatsaechlich in der Output-Datei
# hinterlegt ist
testthat::test_that(
  "Grafiken koennen eingefuegt werden",
  {
    wb <- openxlsx::createWorkbook()
    plt <- ggplot2::ggplot(mtcars, ggplot2::aes(x = disp, y = hp)) +
      ggplot2::geom_point()

    testthat::expect_no_error(insert_worksheet_image(wb, "figure", plt))
    openxlsx::saveWorkbook(wb, output_file_path, overwrite = TRUE)

    unzip(output_file_path, files = "xl/media/image1.png", exdir = output_dir)
    testthat::expect_true(
      file.exists(file.path(output_dir, "xl/media/image1.png")))

    cleanup_test("test_out")
  }
)
