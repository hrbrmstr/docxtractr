context("docx extraction works")
test_that("we can do something", {

  doc <- read_docx(system.file("examples/data.docx", package="docxtractr"))

  x <- capture.output(print(doc))

  x <- capture.output(docx_describe_tbls(doc))

  expect_equal(length(docx_extract_all(doc)), 1)

  expect_equal(length(docx_extract_all_tbls(doc)), 1)

  expect_that(doc, is_a("docx"))
  expect_that(docx_tbl_count(doc), equals(1))
  expect_that(docx_extract_tbl(doc, 1), is_a("tbl"))

  complx <- read_docx(system.file("examples/complex.docx", package="docxtractr"))
  expect_that(docx_tbl_count(complx), equals(5))

  tmp_3 <- docx_extract_tbl(complx, 3)
  tmp_4 <- docx_extract_tbl(complx, 4)
  tmp_5 <- docx_extract_tbl(complx, 5)

  expect_that(tmp_3, is_a("tbl"))
  expect_that(tmp_4, is_a("tbl"))
  expect_that(tmp_5, is_a("tbl"))

  expect_that(nrow(tmp_3), equals(6))
  expect_that(ncol(tmp_4), equals(3))
  expect_that(nrow(tmp_5), equals(6))

  tmp_6 <- assign_colnames(tmp_5, 1)

  expect_equal(colnames(tmp_6), c("Aa", "Bb", "Cc"))

  cmnt <- read_docx(system.file("examples/comments.docx", package="docxtractr"))

  expect_equal(docx_cmnt_count(cmnt), 3)

  x <- capture.output(docx_describe_cmnts(cmnt))

  x <- capture.output(print(cmnt))

  expect_equal(nrow(docx_extract_all_cmnts(cmnt)), 3)

  real_world <- read_docx(system.file("examples/realworld.docx", package="docxtractr"))
  tbls <- docx_extract_all_tbls(real_world)
  expect_equal(colnames(mcga(assign_colnames(tbls[[1]], 2))),
               c("country", "birthrate", "death_rate", "population_growth_2005",
                 "population_growth_2050", "relative_place_in_transition", "social_factors_1",
                 "social_factors_2", "social_factors_3"))

})
