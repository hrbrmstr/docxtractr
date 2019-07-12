context("PPTX conversion works")
test_that("we can convert a PPTX if LibreOffice Installed", {
  lp = try({
    docxtractr:::lo_find()
  }, silent = TRUE)
  if (!inherits(lp, "try-error")) {
    path <- system.file("examples/ex.pptx", package = "docxtractr")
    pdf <- convert_to_pdf(path, pdf_file = tempfile(fileext = ".pdf"))
    expect_true(file.size(pdf) > 0)
  }
})

test_that("we can convert a DOCX to PDF if LibreOffice Installed", {
  lp = try({
    docxtractr:::lo_find()
  }, silent = TRUE)
  if (!inherits(lp, "try-error")) {
    path <- system.file("examples/data.docx", package = "docxtractr")
    pdf <- convert_to_pdf(path, pdf_file = tempfile(fileext = ".pdf"))
    expect_true(file.size(pdf) > 0)
  }
})

