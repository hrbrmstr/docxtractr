#' Extract all tables from a Word document
#'
#' This function makes no assumptions about an
#'
#' @param docx \code{docx} object read with \code{read_docx}
#' @param guess_header should the function make a guess as to the existense of
#'        a header in a table? (Default: \code{TRUE})
#' @param trim trim leading/trailing whitespace (if any) in cells? (default: \code{TRUE})
#' @return \code{list} of \code{data.frame}s or an empty \code{list} if no
#'         tables exist in \code{docx}
#' @seealso \code{\link{assign_colnames}}, \code{\link{docx_extract_tbl}}
#' @export
#' @examples
#' # a "real" Word doc
#'
#' real_world <- read_docx(system.file("examples/realworld.docx", package="docxtractr"))
#' docx_tbl_count(real_world)
#'
#' # get all the tables
#' tbls <- docx_extract_all(real_world)
docx_extract_all <- function(docx, guess_header=TRUE, trim=TRUE) {

  ensure_docx(docx)
  if (docx_tbl_count(docx) < 1) return(list())

  ns <- docx$ns

  lapply(1:docx_tbl_count(docx), function(i) {
    hdr <- FALSE
    if (guess_header) {
      tbl <- docx$tbls[[i]]
      rows <- xml_find_all(tbl, "./w:tr", ns=ns)
      hdr <- !is.na(has_header(tbl, rows, ns))
    }
    docx_extract_tbl(docx, i, hdr, trim)
  })

}
