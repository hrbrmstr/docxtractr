#' Read in a Word document for table extraction
#'
#' Path must be local (i.e. not a URL)
#'
#' @param path path to the Word document
#' @importFrom xml2 read_xml
#' @importFrom tools file_ext
#' @export
#' @examples
#' doc <- read_docx(system.file("examples/data.docx", package="docxtractr"))
#' class(doc)
read_docx <- function(path) {

  path <- path.expand(path)

  if (!file_ext(path) == "docx") stop("read_docx only works with '.docx' files", call.=FALSE)
  if (!file.exists(path)) stop(sprintf("Cannot find '%s'", path), call.=FALSE)

  # make temporary things for us to work with
  tmpd <- tempdir()
  tmpf <- tempfile(tmpdir=tmpd, fileext=".zip")

  # copy docx to zip (not entirely necessary)
  file.copy(path, tmpf)
  # unzip it
  unzip(tmpf, exdir=sprintf("%s/docdata", tmpd))

  # read the actual XML document
  doc <- read_xml(sprintf("%s/docdata/word/document.xml", tmpd))

  # cleanup
  unlink(tmpf)
  unlink(sprintf("%s/docdata", tmpd), recursive=TRUE)

  # extract the namespace
  ns <- xml_ns(doc)

  # get the tables
  tbls <- xml_find_all(doc, ".//w:tbl", ns=ns)

  # make an object for other functions to work with
  docx <- list(docx=doc, ns=ns, tbls=tbls, path=path)

  # special class helps us work with these things
  class(docx) <- "docx"

  docx

}
