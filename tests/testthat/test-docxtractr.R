context("basic functionality")
test_that("we can do something", {

  doc <- read_docx(system.file("examples/data.docx", package="docxtractr"))

  expect_that(doc, is_a("docx"))
  expect_that(docx_tbl_count(doc), equals(1))
  expect_that(docx_extract_tbl(doc, 1), is_a("data.frame"))

  complx <- read_docx(system.file("examples/complex.docx", package="docxtractr"))
  expect_that(docx_tbl_count(complx), equals(5))

  tmp_3 <- docx_extract_tbl(complx, 3)
  tmp_4 <- docx_extract_tbl(complx, 4)
  tmp_5 <- docx_extract_tbl(complx, 5)

  expect_that(tmp_3, is_a("data.frame"))
  expect_that(tmp_4, is_a("data.frame"))
  expect_that(tmp_5, is_a("data.frame"))

  expect_that(nrow(tmp_3), equals(6))
  expect_that(ncol(tmp_4), equals(3))
  expect_that(nrow(tmp_5), equals(6))

})
