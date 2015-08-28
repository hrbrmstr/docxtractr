#' docxtractr is an R pacakge for extracting tables out of Word documents (docx)
#'
#' Microsoft Word docx files provide an XML structure that is fairly
#' straightforward to navigate, especially when it applies to Word tables. The
#' docxtractr package provides tools to determine table count + table structure and
#' extract tables from Microsoft Word docx documents.
#'
#' @name docxtractr
#' @docType package
#'
#' @author Bob Rudis (@@hrbrmstr)
#' @importFrom xml2 xml_find_all xml_text xml_ns xml_find_one xml_attrs
#' @importFrom dplyr bind_rows
#' @importFrom tools file_ext
#' @importFrom utils download.file unzip
NULL
