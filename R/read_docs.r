#' Read in a Word document for table extraction
#'
#' Local file path or URL pointing to a \code{.docx} file.
#'
#' @param path path to the Word document
#' @importFrom xml2 read_xml
#' @export
#' @examples
#' doc <- read_docx(system.file("examples/data.docx", package="docxtractr"))
#' class(doc)
#' \dontrun{
#' # from a URL
#' budget <- read_docx(
#' "http://rud.is/dl/1.DOCX")
#' }
read_docx <- function(path) {

  # make temporary things for us to work with
  tmpd <- tempdir()
  tmpf <- tempfile(tmpdir=tmpd, fileext=".zip")

  on.exit({ #cleanup
    unlink(tmpf)
    unlink(sprintf("%s/docdata", tmpd), recursive=TRUE)
  })

  if (is_url(path)) {
    download.file(path, tmpf)
  } else {
    path <- path.expand(path)
    if (!file.exists(path)) stop(sprintf("Cannot find '%s'", path), call.=FALSE)
    # copy docx to zip (not entirely necessary)
    file.copy(path, tmpf)
  }
  # unzip it
  unzip(tmpf, exdir=sprintf("%s/docdata", tmpd))

  # read the actual XML document
  doc <- read_xml(sprintf("%s/docdata/word/document.xml", tmpd))

  # extract the namespace
  ns <- xml_ns(doc)

  # get the tables
  tbls <- xml_find_all(doc, ".//w:tbl", ns=ns)

  if (file.exists(sprintf("%s/docdata/word/comments.xml", tmpd))) {
    docmnt <- read_xml(sprintf("%s/docdata/word/comments.xml", tmpd))
    # get the comments
    cmnts <- xml_find_all(docmnt, ".//w:comment", ns=ns)
  } else {
    cmnts <- xml_find_all(doc, ".//w:comment", ns=ns)
  }

  # make an object for other functions to work with
  docx <- list(docx=doc, ns=ns, tbls=tbls, cmnts=cmnts, path=path)

  # special class helps us work with these things
  class(docx) <- "docx"

  docx

}
