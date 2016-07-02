#' Returns a description of all the tables in the Word document
#'
#' This function will attempt to discern the structure of each of the tables
#' in \code{docx} and print this information
#'
#' @param docx \code{docx} object read with \code{read_docx}
#' @export
#' @examples
#' complx <- read_docx(system.file("examples/complex.docx", package="docxtractr"))
#' docx_tbl_count(complx)
#' docx_describe_tbls(complx)
docx_describe_tbls <- function(docx) {

  ensure_docx(docx)
  if (!docx_tbl_count(docx) > 0) {
    message("No tables in document")
    return(invisible(NULL))
  }

  ns <- docx$ns
  tbls <- docx$tbls

  cat(sprintf("Word document [%s]\n\n", docx$path))

  for (i in 1:length(tbls)) {

    tbl <- tbls[[i]]

    cells <- xml_find_all(tbl, "./w:tr/w:tc", ns=ns)
    rows <- xml_find_all(tbl, "./w:tr", ns=ns)

    cell_count_by_row <- sapply(rows, function(row) { length(xml_find_all(row, "./w:tc", ns)) })
    row_counts <- paste0(unique(cell_count_by_row), collapse=", ")
    max_cell_count <- max(cell_count_by_row)

    cat(sprintf("Table %d\n  total cells: %d\n  row count  : %d\n",
                i, length(cells), length(rows)))

    # simplistic test for whether table is uniform rows x cells == cell count
    if ((max_cell_count * length(rows)) == length(cells)) {
      cat("  uniform    : likely!\n")
    } else {
      cat(sprintf(
        "  uniform    : unlikely => found differing cell counts (%s) across some rows\n",
        row_counts))
    }

    # microsoft has a tag for some table structure info. examine it to
    # see if the creator of the header made the first row special which
    # will likely mean it's a header candidate
    hdr <- has_header(tbl, rows, ns)
    if (is.na(hdr)) {
      cat("  has header : unlikely\n")
    } else {
      cat(sprintf("  has header : likely! => possibly [%s]\n", hdr))
    }

    cat("\n")

  }

}

#' Returns information about the comments in the Word document
#'
#' @param docx \code{docx} object read with \code{read_docx}
#' @export
#' @examples
#' cmnts <- read_docx(system.file("examples/comments.docx", package="docxtractr"))
#' docx_cmnt_count(cmnts)
#' docx_describe_cmnts(cmnts)
docx_describe_cmnts <- function(docx) {

  ensure_docx(docx)
  if (!docx_cmnt_count(docx) > 0) {
    message("No comments in document")
    return(invisible(NULL))
  }

  ns <- docx$ns
  cmnts <- docx$cmnts

  cat(sprintf("Word document [%s]\n\n", docx$path))

  cat(sprintf("Found %d comments.\n", length(cmnts)))

  map_df(xml_attrs(cmnts), function(x) {
    as_data_frame(t(cbind.data.frame(x, stringsAsFactors=FALSE)))
  }) -> meta

  cmnt_df <- dplyr::bind_cols(meta,
                       cbind.data.frame(comment_text=xml_text(cmnts),
                                        stringsAsFactors=FALSE))

  aut_df <- dplyr::count(cmnt_df, author)
  aut_df <- dplyr::arrange(aut_df, -n)
  print(select(aut_df, author, `# Comments`=n))

}

#' Display information about the document
#'
#' @param x \code{docx} object
#' @param ... ignored
#' @export
print.docx <- function(x, ...) {
  docx_describe_tbls(x)
  docx_describe_cmnts(x)
}
