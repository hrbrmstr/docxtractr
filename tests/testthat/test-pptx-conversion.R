context("PPTX conversion works")
test_that("we can convert a PPTX if LibreOffice Installed", {
  lp = try({
    docxtractr:::lo_find()
  }, silent = TRUE)
  if (!inherits(lp, "try-error")) {
    path <- system.file("examples/ex.pptx", package = "docxtractr")
    pdf <- convert_pptx_to_pdf(path, pdf_file = tempfile(fileext = ".pdf"))
    expect_true(file.size(pdf) > 0)
  }
})

