# 0.6.5

- Fix CRAN check errors on Fedora

# 0.6.1

- Fix for errors introduced by an update of the tidyverse

# 0.6.0

- Enable support for accepting or rejecting tracked changes when
  reading in the document. Ref #19

# 0.5.0

- .doc input supported (via Chris Muir)
- UTF-8 filename support for Windows-1252 locale

# 0.4.0

- add a `preserve` logical paramater to tbl extraction functions to support preserving intra-cell whitespace (implements #9)
- use `httr` vs `download.file()` for URL retrieval (fixes #10)

# 0.3.0 WIP

- return tibbles where possible & not stomping on input type (#7)
- change tests to test for `tbl` vs `data.frame` (related to #7)
- don't stomp on data frame-ish input type in `assign_colnames()`
- prefix `::` (non-user facing tweak)
- switch all `*apply()` to `purrr` calls since we bother to import `purrr`  (non-user facing tweak)
- Make Column Names Great Again! (`mgca()` function added. The `janitor` package has a more robust function.)

# 0.2.0 released

- update for new xml2 pkg compatibility
- added ability to extract comments

# 0.1.1 released

- had to change budget docx url since it was 404'ing

# 0.1.0

- new function to extract all tables and a function
  to cleanup column names in scraped tables
