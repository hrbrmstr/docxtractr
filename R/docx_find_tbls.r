#' Get number of tables in a Word document
#'
#' @param docx \code{docx} object read with \code{read_docx}
#' @return numeric
#' @export
#' @examples
#' complx <- read_docx(system.file("examples/complex.docx", package="docxtractr"))
#' docx_tbl_count(complx)
docx_tbl_count <- function(docx) {
  ensure_docx(docx)
  length(docx$tbls)
}

#' Get number of comments in a Word document
#'
#' @param docx \code{docx} object read with \code{read_docx}
#' @return numeric
#' @export
#' @examples
#' cmnts <- read_docx(system.file("examples/comments.docx", package="docxtractr"))
#' docx_cmnt_count(cmnts)
docx_cmnt_count <- function(docx) {
  ensure_docx(docx)
  length(docx$cmnts)
}
