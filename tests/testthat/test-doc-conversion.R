context("DOC conversion works")
test_that("we can convert a DOC to DOCX if LibreOffice Installed", {
  lp = try({
    docxtractr:::lo_find()
  }, silent = TRUE)
  if (!inherits(lp, "try-error")) {
    path <- system.file("examples/preserve.doc", package = "docxtractr")
    doc = read_docx(path)
    expect_that(doc, is_a("docx"))
  }
})
