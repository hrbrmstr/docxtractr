#' Extract all tables from a Word document
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
#' tbls <- docx_extract_all_tbls(real_world)
docx_extract_all_tbls <- function(docx, guess_header=TRUE, trim=TRUE) {

  ensure_docx(docx)
  if (docx_tbl_count(docx) < 1) return(list())

  ns <- docx$ns

  purrr::map(1:docx_tbl_count(docx), function(i) {
    hdr <- FALSE
    if (guess_header) {
      tbl <- docx$tbls[[i]]
      rows <- xml_find_all(tbl, "./w:tr", ns=ns)
      hdr <- !is.na(has_header(tbl, rows, ns))
    }
    docx_extract_tbl(docx, i, hdr, trim)
  })

}

#' Extract all tables from a Word document
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
#' tbls <- docx_extract_all_tbls(real_world)
docx_extract_all <- function(docx, guess_header=TRUE, trim=TRUE) {
  message("docx_extract_all() is deprecated; use docx_extract_all_tbls()")
  docx_extract_all_tbls(docx, guess_header, trim)
}

#' Extract all comments from a Word document
#'
#' @param docx \code{docx} object read with \code{read_docx}
#' @return \code{data_frame} of comment id, author & text
#' @export
#' @examples
#' cmnts <- read_docx(system.file("examples/comments.docx", package="docxtractr"))
#' docx_cmnt_count(cmnts)
#' docx_describe_cmnts(cmnts)
#' docx_extract_all_cmnts(cmnts)
docx_extract_all_cmnts <- function(docx) {

  ensure_docx(docx)
  if (docx_cmnt_count(docx) < 1) return(data_frame())

  ns <- docx$ns

  comments <- docx$cmnts

  map_df(xml_attrs(comments), function(x) {
    as_data_frame(t(cbind.data.frame(x, stringsAsFactors=FALSE)))
  }) -> meta

  bind_cols(meta,
            cbind.data.frame(comment_text=xml_text(comments),
                             stringsAsFactors=FALSE))

}
