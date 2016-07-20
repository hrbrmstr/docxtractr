#' docxtractr is an R package for extracting tables and comments out of Word documents (docx)
#'
#' Microsoft Word docx files provide an XML structure that is fairly
#' straightforward to navigate, especially when it applies to Word tables. The
#' docxtractr package provides tools to determine table count + table structure and
#' extract tables from Microsoft Word docx documents. It also provides tools to determine
#' comment count and extract comments from Word docx documents.
#'
#' @name docxtractr
#' @docType package
#'
#' @author Bob Rudis (@@hrbrmstr)
#' @importFrom xml2 xml_find_all xml_text xml_ns xml_find_first xml_attrs
#' @importFrom tibble data_frame as_data_frame
#' @importFrom dplyr bind_rows bind_cols count arrange select
#' @importFrom tools file_ext
#' @importFrom utils download.file unzip
#' @importFrom purrr map_df
NULL
