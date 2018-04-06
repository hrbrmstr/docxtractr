#' Extract all comments from a Word document
#'
#' @md
#' @param docx \code{docx} object read with \code{read_docx}
#' @param include_text if `TRUE` then the text associated with the comment will
#'        also be included
#' @return \code{data_frame} of comment id, author & text
#' @export
#' @examples
#' cmnts <- read_docx(system.file("examples/comments.docx", package="docxtractr"))
#' docx_cmnt_count(cmnts)
#' docx_describe_cmnts(cmnts)
#' docx_extract_all_cmnts(cmnts)
docx_extract_all_cmnts <- function(docx, include_text=FALSE) {

  ensure_docx(docx)
  if (docx_cmnt_count(docx) < 1) return(tibble::data_frame())

  ns <- docx$ns

  comments <- docx$cmnts

  purrr::map_df(xml2::xml_attrs(comments), function(x) {
    tibble::as_data_frame(t(cbind.data.frame(x, stringsAsFactors=FALSE)))
  }) -> meta

  dplyr::bind_cols(
    meta,
    cbind.data.frame(comment_text=xml2::xml_text(comments), stringsAsFactors=FALSE)
  ) -> out

  if (include_text) {

    doc <- docx$docx

    out$word_src <- purrr::map_chr(out$id, ~{
      xml_find_all(
        doc,
        sprintf("//w:commentRangeStart[@w:id='%s']/following-sibling::*[
             count(. | //w:commentRangeEnd[@w:id='%s']/preceding-sibling::*) =
             count(//w:commentRangeEnd[@w:id='%s']/preceding-sibling::*)]",
                .x, .x, .x)
      ) %>%
        xml_text() %>%
        paste0(collapse=" ")

    })


  }

  tibble::as_tibble(out)

}
