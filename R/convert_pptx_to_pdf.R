#' Convert a Document (usually PowerPoint) to a PDF
#'
#' @md
#' @param path path to the document, can be PowerPoint or DOCX
#' @param pdf_file output PDF file name.  By default, creates a PDF in the
#'        same directory as the `path` file.
#'        This functionality requires the use of
#'        LibreOffice and the `soffice` binary it contains.  See
#'        [set_libreoffice_path] for more information.  Note,
#' @export
#' @examples
#' \dontrun{
#' path = system.file("examples/ex.pptx", package="docxtractr")
#' pdf <- convert_to_pdf(path, pdf_file = tempfile(fileext = ".pdf"))
#' path = system.file("examples/data.docx", package="docxtractr")
#' pdf_doc <- convert_to_pdf(path, pdf_file = tempfile(fileext = ".pdf"))
#' }
convert_to_pdf <- function(path, pdf_file = sub("[.]pptx", ".pdf", path)) {
  stopifnot(is_pptx(path) | is_doc(path) | is_docx(path))

  lo_assert()
  lo_path <- getOption("path_to_libreoffice")

  # making temporary file because by default soffice
  # will make sub("[.]pptx", ".pdf", path) output
  # and don't want to do that in case pdf_file in other location
  cp_path = tempfile(fileext = ".pptx")
  cp_pdf = sub("[.](pptx|docx|doc)$", ".pdf", cp_path)
  file.copy(path, cp_path)

  if (Sys.info()["sysname"] == "Windows") {
    convert_win(lo_path, dirname(cp_path), cp_path, convert_to = "pdf")
  } else {
    convert_osx(lo_path, dirname(cp_path), cp_path, convert_to = "pdf")
  }
  if (!file.exists(cp_pdf)) {
    stop("Conversion from PPTX to PDF did not succeed")
  }
  file.copy(cp_pdf, pdf_file)
  return(pdf_file)
}
