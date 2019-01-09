#' Extract a table from a Word document
#'
#' Given a document read with \code{read_docx} and a table to extract (optionally
#' indicating whether there was a header or not and if cell whitepace trimming is
#' desired) extract the contents of the table to a \code{data.frame}.
#'
#' @md
#' @param docx \code{docx} object read with \code{read_docx}
#' @param tbl_number which table to extract (defaults to \code{1})
#' @param header assume first row of table is a header row? (default; \code{TRUE})
#' @param preserve preserve line breaks within a cell? Default: `FALSE`. NOTE: This overrides `trim`.
#' @param trim trim leading/trailing whitespace (if any) in cells? (default: \code{TRUE})
#' @return \code{data.frame}
#' @seealso \code{\link{docx_extract_all}}, \code{\link{docx_extract_tbl}},
#'          \code{\link{assign_colnames}}
#' @export
#' @examples
#' doc3 <- read_docx(system.file("examples/data3.docx", package="docxtractr"))
#' docx_extract_tbl(doc3, 3)
#'
#' intracell_whitespace <- read_docx(system.file("examples/preserve.docx", package="docxtractr"))
#' docx_extract_tbl(intracell_whitespace, 2, preserve=FALSE)
#' docx_extract_tbl(intracell_whitespace, 2, preserve=TRUE)
docx_extract_tbl <- function(docx, tbl_number=1, header=TRUE, preserve=FALSE, trim=TRUE) {

  ensure_docx(docx)
  if ((tbl_number < 1) | (tbl_number > docx_tbl_count(docx))) {
    stop("'tbl_number' is invalid.", call.=FALSE)
  }

  if (preserve) trim <- FALSE

  ns <- docx$ns
  tbl <- docx$tbls[[tbl_number]]

  cells <- xml2::xml_find_all(tbl, "./w:tr/w:tc", ns=ns)
  rows <- xml2::xml_find_all(tbl, "./w:tr", ns=ns)

  purrr::map_df(rows, ~{
    res <- xml2::xml_find_all(.x, "./w:tc", ns=ns)
    if (preserve) {
      purrr::map(res, ~{
        paras <- xml2::xml_text(xml2::xml_find_all(.x, "./w:p", ns=ns))
        paste0(paras, collapse="\n")
      }) -> vals
    } else {
      vals <- xml2::xml_text(res, trim=trim)
    }
    names(vals) <- sprintf("V%d", 1:length(vals))
    as.list(vals)
    # data.frame(as.list(vals), stringsAsFactors=FALSE)
  }) -> dat

  if (header) {
    hopeful_names <- make.names(dat[1,])
    colnames(dat) <- hopeful_names
    dat <- dat[-1,]
  } else {
    hdr <- has_header(tbl, rows, ns)
    if (!is.na(hdr)) {
      message("NOTE: header=FALSE but table has a marked header row in the Word document")
    }
  }

  rownames(dat) <- NULL

  class(dat) <- c("tbl_df", "tbl", "data.frame")
  dat

}
