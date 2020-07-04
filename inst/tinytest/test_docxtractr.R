library(docxtractr)

doc <- read_docx(system.file("examples/data.docx", package="docxtractr"))

suppressMessages(x <- capture.output(print(doc)))

x <- capture.output(docx_describe_tbls(doc))

suppressMessages(expect_equal(length(docx_extract_all(doc)), 1))

expect_equal(length(docx_extract_all_tbls(doc)), 1)

expect_true(inherits(doc, "docx"))
expect_equal(docx_tbl_count(doc), 1)
expect_true(inherits(docx_extract_tbl(doc, 1), "tbl"))

complx <- read_docx(system.file("examples/complex.docx", package="docxtractr"))
expect_equal(docx_tbl_count(complx) ,5)

tmp_3 <- docx_extract_tbl(complx, 3)
tmp_4 <- docx_extract_tbl(complx, 4)
tmp_5 <- docx_extract_tbl(complx, 5)

expect_true(inherits(tmp_3, "tbl"))
expect_true(inherits(tmp_4, "tbl"))
expect_true(inherits(tmp_5, "tbl"))

expect_equal(nrow(tmp_3), 6)
expect_equal(ncol(tmp_4), 3)
expect_equal(nrow(tmp_5), 6)

tmp_6 <- assign_colnames(tmp_5, 1)

expect_equal(colnames(tmp_6), c("Aa", "Bb", "Cc"))

cmnt <- read_docx(system.file("examples/comments.docx", package="docxtractr"))

expect_equal(docx_cmnt_count(cmnt), 3)

x <- capture.output(docx_describe_cmnts(cmnt))

suppressMessages(x <- capture.output(print(cmnt)))

expect_equal(nrow(docx_extract_all_cmnts(cmnt)), 3)

real_world <- read_docx(system.file("examples/realworld.docx", package="docxtractr"))
tbls <- docx_extract_all_tbls(real_world)
expect_equal(
  colnames(mcga(assign_colnames(tbls[[1]], 2))),
  c("country", "birthrate", "death_rate", "population_growth_2005",
    "population_growth_2050", "relative_place_in_transition", "social_factors_1",
    "social_factors_2", "social_factors_3")
)

# docx-conversion ---------------------------------------------------------

if (at_home()) { # CRAN will not have LibreOffice installed

  lp = try({
    docxtractr:::lo_find()
  }, silent = TRUE)
  if (!inherits(lp, "try-error")) {
    path <- system.file("examples/preserve.doc", package = "docxtractr")
    doc = read_docx(path)
    expect_that(doc, is_a("docx"))
  }

}

# pptx conversion ---------------------------------------------------------

if (at_home()) { # CRAN will not have LibreOffice installed

  lp = try({
    docxtractr:::lo_find()
  }, silent = TRUE)

  if (!inherits(lp, "try-error")) {
    path <- system.file("examples/ex.pptx", package = "docxtractr")
    pdf <- convert_to_pdf(path, pdf_file = tempfile(fileext = ".pdf"))
    expect_true(file.size(pdf) > 0)
  }

}

if (at_home()) {
  lp = try({
    docxtractr:::lo_find()
  }, silent = TRUE)
  if (!inherits(lp, "try-error")) {
    path <- system.file("examples/data.docx", package = "docxtractr")
    pdf <- convert_to_pdf(path, pdf_file = tempfile(fileext = ".pdf"))
    expect_true(file.size(pdf) > 0)
  }

}
