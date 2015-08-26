#' Make a specific row the column names for the specified data.frame
#'
#' Many tables in Word documents are in twisted formats where there may be
#' labels or other oddities mixed in that make it difficult to work with the
#' underlying data. This function makes it easy to identify a particular row
#' in a scraped \code{data.frame} as the one containing column names and
#' have it become the column names, removing it and (optionally) all of the
#' rows before it (since that's usually what needs to be done).
#'
#' @param dat can be any \code{data.frame} but is intended for use with
#'        ones retuned by this package
#' @param row numeric value indicating the row number that is to become
#'        the column names
#' @param remove remove row specified by \code{row} after making it
#'        the column names? (Default: \code{TRUE})
#' @param remove_previous remove any rows preceeding \code{row}? (Default:
#'        \code{TRUE} but will be assigned whatever is given for
#'        \code{remove}).
#' @return \code{data.frame}
#' @seealso \code{\link{docx_extract_all}}, \code{\link{docx_extract_tbl}}
#' @export
#' @examples
#' # a "real" Word doc
#' real_world <- read_docx(system.file("examples/realworld.docx", package="docxtractr"))
#' docx_tbl_count(real_world)
#'
#' # get all the tables
#' tbls <- docx_extract_all(real_world)
#'
#' # make table 1 better
#' assign_colnames(tbls[[1]], 2)
#'
#' # make table 5 better
#' assign_colnames(tbls[[5]], 2)
assign_colnames <- function(dat, row, remove=TRUE, remove_previous=remove) {

  if ((row > nrow(dat)) | (row < 1)) return(dat)

  # just in case someone shoots us a data.table or other stranger things
  dat <- data.frame(dat, stringsAsFactors=FALSE)

  colnames(dat) <- dat[row,]
  start <- row
  end <- row
  if (remove_previous) start <- 1

  dat <- dat[-(start:end),]
  rownames(dat) <- NULL

  dat

}
