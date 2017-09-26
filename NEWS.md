# 0.4.0 WIP

- add a `preserve` logical paramater to tbl extraction functions to support preserving
  intra-cell whitespace (ref: #9)

# 0.3.0 WIP

- return tibbles where possible & not stomping on input type (#7)
- change tests to test for `tbl` vs `data.frame` (related to #7)
- don't stomp on data frame-ish input type in `assign_colnames()`
- prefix `::` (non-user facing tweak)
- switch all `*apply()` to `purrr` calls since we bother to import `purrr`  (non-user facing tweak)
- Make Column Names Great Again! (`mgca()` function added)

# 0.2.0 released

- update for new xml2 pkg compatibility
- added ability to extract comments

# 0.1.1 released

- had to change budget docx url since it was 404'ing

# 0.1.0

- new function to extract all tables and a function
  to cleanup column names in scraped tables
