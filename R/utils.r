# used by cuntions to make sure they are working with a well-formed docx object
ensure_docx <- function(docx) {
  if (!inherits(docx, "docx")) stop("Must pass in a 'docx' object", call.=FALSE)
  if (!(all(sapply(c("docx", "ns", "tbls", "path"), exists, where=docx))))
    stop("'docx' object missing necessary components", call.=FALSE)
}

# test if a w:tbl has a header row
has_header <- function(tbl, rows, ns) {

  # microsoft has a tag for some table structure info. examine it to
  # see if the creator of the header made the first row special which
  # will likely mean it's a header candidate
  look <- try(xml_find_first(tbl, "./w:tblPr/w:tblLook", ns), silent=TRUE)
  if (inherits(look, "try-error")) {
    return(NA)
  } else {
    look_attr <- xml_attrs(look)
    if ("firstRow" %in% names(look_attr)) {
      if (look_attr["firstRow"] == "0") {
        return(NA)
      } else {
        return(paste0(xml_text(xml_find_all(rows[[1]], "./w:tc", ns)), collapse=", "))
      }
    } else {
      return(NA)
    }
  }

}

is_url <- function(path) { grepl("^(http|ftp)s?://", path) }

is_docx <- function(path) { tolower(file_ext(path)) == "docx" }
