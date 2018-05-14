# used by functions to make sure they are working with a well-formed docx object
ensure_docx <- function(docx) {
  if (!inherits(docx, "docx")) stop("Must pass in a 'docx' object", call.=FALSE)
  if (!(all(purrr::map_lgl(c("docx", "ns", "tbls", "path"), exists, where=docx))))
    stop("'docx' object missing necessary components", call.=FALSE)
}

# test if a w:tbl has a header row
has_header <- function(tbl, rows, ns) {

  # microsoft has a tag for some table structure info. examine it to
  # see if the creator of the header made the first row special which
  # will likely mean it's a header candidate
  look <- try(xml2::xml_find_first(tbl, "./w:tblPr/w:tblLook", ns), silent=TRUE)
  if (inherits(look, "try-error")) {
    return(NA)
  } else {
    look_attr <- xml2::xml_attrs(look)
    if ("firstRow" %in% names(look_attr)) {
      if (look_attr["firstRow"] == "0") {
        return(NA)
      } else {
        return(paste0(xml2::xml_text(xml_find_all(rows[[1]], "./w:tc", ns)), collapse=", "))
      }
    } else {
      return(NA)
    }
  }

}

is_url <- function(path) { grepl("^(http|ftp)s?://", path) }

is_docx <- function(path) { tolower(tools::file_ext(path)) == "docx" }

is_doc <- function(path) { tolower(tools::file_ext(path)) == "doc" }

# Copy a file to a new location, throw an error if the copy fails.
file_copy <- function(from, to) {
  fc <- file.copy(from, to)
  if (!fc) stop(sprintf("file copy failure for file %s", from), call.=FALSE)
}

# Save a .doc file as a new .docx file, using the LibreOffice command line
# tools.
convert_doc_to_docx <- function(docx_dir, doc_file) {
  lo_path <- getOption("path_to_libreoffice")
  if (is.null(lo_path)) {
    stop("Cannot determine file path to LibreOffice. ",
         "To download LibreOffice, visit: https://www.libreoffice.org/ \n",
         "If you've already downloaded the software, use function ",
         "'set_libreoffice_path()' to point R to your local 'soffice.exe' file",
         call. = FALSE)
  }
  if (Sys.info()["sysname"] == "Windows") {
    convert_win(lo_path, docx_dir, doc_file)
  } else {
    convert_osx(lo_path, docx_dir, doc_file)
  }
}

# .docx to .doc convertion for Windows
convert_win <- function(lo_path, docx_dir, doc_file) {
  cmd <- sprintf('"%s" -convert-to docx:"MS Word 2007 XML" -headless -outdir "%s" "%s"',
                 lo_path,
                 docx_dir,
                 doc_file)
  system(cmd, show.output.on.console = FALSE)
}

# .docx to .doc convertion for OSX
convert_osx <- function(lo_path, docx_dir, doc_file) {
  cmd <- sprintf('"%s" --convert-to docx:"MS Word 2007 XML" --headless --outdir "%s" "%s"',
                 lo_path,
                 docx_dir,
                 doc_file)
  res <- system(cmd, intern = TRUE)
}


#' Point to Local soffice.exe File
#'
#' Function to set an option that points to the local LibreOffice file
#' \code{soffice.exe}.
#'
#' @param path path to the LibreOffice soffice file
#'
#' @details For a list of possible file path locations for \code{soffice.exe},
#'  see \url{https://github.com/hrbrmstr/docxtractr/issues/5#issuecomment-233181976}
#'
#' @return Returns nothing, function sets the option variable
#'  \code{path_to_libreoffice}.
#' @export
#'
#' @examples \dontrun{
#' set_libreoffice_path("local/path/to/soffice.exe")
#' }
set_libreoffice_path <- function(path) {
  stopifnot(is.character(path))

  if (!file.exists(path)) stop(sprintf("Cannot find '%s'", path), call.=FALSE)
  options("path_to_libreoffice" = path)
}
