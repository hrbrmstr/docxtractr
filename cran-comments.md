## Test environments

* local OS X install, R 3.5.2
* ubuntu 14.04 (on travis-ci), R 3.5.2
* r-hub (fedora & windows)
* win-builder (devel and release)

## R CMD check results

0 errors | 0 warnings | 1 note

* checking CRAN incoming feasibility ... NOTE

---

This is an update to fix the errors introduced by the 
recent tidyverse update as noted by Kurt on 2018-01-05
(https://cran.r-project.org/web/checks/check_results_docxtractr.html)

The code has been modified to account for the new
behavior of the tidyverse. All tests and examples have been
left intact.
