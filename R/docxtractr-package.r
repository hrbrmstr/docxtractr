#' Extract Data Tables and Comments from 'Microsoft' 'Word' Documents
#'
#' Microsoft Word `docx`` files provide an XML structure that is fairly
#' straightforward to navigate, especially when it applies to Word tables. The
#' `docxtractr`` package provides tools to determine table count + table structure and
#' extract tables from Microsoft Word docx documents. It also provides tools to determine
#' comment count and extract comments from Word `docx`` documents.
#'
#' @name docxtractr
#' @docType package
#'
#' @author Bob Rudis (bob@@rud.is)
#' @importFrom xml2 xml_find_all xml_text xml_ns xml_find_first xml_attrs read_xml
#' @importFrom dplyr bind_cols count arrange select
#' @importFrom tools file_ext
#' @importFrom utils unzip globalVariables
#' @importFrom purrr map_df map map_int map_chr map_lgl
#' @importFrom httr GET stop_for_status write_disk
NULL

