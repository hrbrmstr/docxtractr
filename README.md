<!-- README.md is generated from README.Rmd. Please edit that file -->
![](docxtractr-logo.png)

docxtractr is an R pacakge for extracting tables out of Word documents (docx)

Microsoft Word docx files provide an XML structure that is fairly straightforward to navigate, especially when it applies to Word tables. The docxtractr package provides tools to determine table count, table structure and extract tables from Microsoft Word docx documents.

The following functions are implemented:

-   `read_docx`: Read in a Word document for table extraction
-   `docx_describe_tbls`: Returns a description of all the tables in the Word document
-   `docx_extract_tbl`: Extract a table from a Word document
-   `docx_tbl_count`: Get number of tables in a Word document

The following data file are included:

-   `system.file("examples/data.docx", package="docxtractr")`: Word docx with 1 table
-   `system.file("examples/data3.docx", package="docxtractr")`: Word docx with 3 tables
-   `system.file("examples/none.docx", package="docxtractr")`: Word docx with 0 tables
-   `system.file("examples/complex.docx", package="docxtractr")`: Word docx with non-uniform tables

### News

-   Version 0.0.0.9000 released

### Installation

``` r
devtools::install_github("hrbrmstr/docxtractr")
```

### Usage

``` r
library(docxtractr)

# current verison
packageVersion("docxtractr")
#> [1] '0.0.0.9000'

# one table
doc <- read_docx(system.file("examples/data.docx", package="docxtractr"))

docx_tbl_count(doc)
#> [1] 1

docx_describe_tbls(doc)
#> Word document [/Library/Frameworks/R.framework/Versions/3.2/Resources/library/docxtractr/examples/data.docx]
#> 
#> Table 1
#>   total cells: 16
#>   row count  : 4
#>   uniform    : likely!
#>   has header : likely! => possibly [This, Is, A, Column]

docx_extract_tbl(doc, 1)
#> Source: local data frame [3 x 4]
#> 
#>   This      Is     A   Column
#> 1    1     Cat   3.4      Dog
#> 2    3    Fish 100.3     Bird
#> 3    5 Pelican   -99 Kangaroo

docx_extract_tbl(doc)
#> Source: local data frame [3 x 4]
#> 
#>   This      Is     A   Column
#> 1    1     Cat   3.4      Dog
#> 2    3    Fish 100.3     Bird
#> 3    5 Pelican   -99 Kangaroo

docx_extract_tbl(doc, header=FALSE)
#> NOTE: header=FALSE but table has a marked header row in the Word document
#> Source: local data frame [4 x 4]
#> 
#>     V1      V2    V3       V4
#> 1 This      Is     A   Column
#> 2    1     Cat   3.4      Dog
#> 3    3    Fish 100.3     Bird
#> 4    5 Pelican   -99 Kangaroo

# three tables
doc3 <- read_docx(system.file("examples/data3.docx", package="docxtractr"))

docx_extract_tbl(doc3, 3)
#> Source: local data frame [6 x 2]
#> 
#>   Foo Bar
#> 1  Aa  Bb
#> 2  Dd  Ee
#> 3  Gg  Hh
#> 4   1   2
#> 5  Zz  Jj
#> 6  Tt  ii

docx_tbl_count(doc3)
#> [1] 3

docx_describe_tbls(doc3)
#> Word document [/Library/Frameworks/R.framework/Versions/3.2/Resources/library/docxtractr/examples/data3.docx]
#> 
#> Table 1
#>   total cells: 16
#>   row count  : 4
#>   uniform    : likely!
#>   has header : likely! => possibly [This, Is, A, Column]
#> 
#> Table 2
#>   total cells: 12
#>   row count  : 4
#>   uniform    : likely!
#>   has header : likely! => possibly [Foo, Bar, Baz]
#> 
#> Table 3
#>   total cells: 14
#>   row count  : 7
#>   uniform    : likely!
#>   has header : likely! => possibly [Foo, Bar]

# no tables
none <- read_docx(system.file("examples/none.docx", package="docxtractr"))

docx_tbl_count(none)
#> [1] 0

# wrapping in try since it will return an error
# use docx_tbl_count before trying to extract in scripts/production
try(docx_describe_tbls(none))
try(docx_extract_tbl(none, 2))

# 5 tables, with two in sketchy formats
complx <- read_docx(system.file("examples/complex.docx", package="docxtractr"))

docx_tbl_count(complx)
#> [1] 5

docx_describe_tbls(complx)
#> Word document [/Library/Frameworks/R.framework/Versions/3.2/Resources/library/docxtractr/examples/complex.docx]
#> 
#> Table 1
#>   total cells: 16
#>   row count  : 4
#>   uniform    : likely!
#>   has header : likely! => possibly [This, Is, A, Column]
#> 
#> Table 2
#>   total cells: 12
#>   row count  : 4
#>   uniform    : likely!
#>   has header : likely! => possibly [Foo, Bar, Baz]
#> 
#> Table 3
#>   total cells: 14
#>   row count  : 7
#>   uniform    : likely!
#>   has header : likely! => possibly [Foo, Bar]
#> 
#> Table 4
#>   total cells: 11
#>   row count  : 4
#>   uniform    : unlikely => found differing cell counts (3, 2) across some rows 
#>   has header : likely! => possibly [Foo, Bar, Baz]
#> 
#> Table 5
#>   total cells: 21
#>   row count  : 7
#>   uniform    : likely!
#>   has header : unlikely

docx_extract_tbl(complx, 3, header=TRUE)
#> Source: local data frame [6 x 2]
#> 
#>   Foo Bar
#> 1  Aa  Bb
#> 2  Dd  Ee
#> 3  Gg  Hh
#> 4   1   2
#> 5  Zz  Jj
#> 6  Tt  ii

docx_extract_tbl(complx, 4, header=TRUE)
#> Source: local data frame [3 x 3]
#> 
#>   Foo  Bar Baz
#> 1  Aa BbCc  NA
#> 2  Dd   Ee  Ff
#> 3  Gg   Hh  ii

docx_extract_tbl(complx, 5, header=TRUE)
#> Source: local data frame [6 x 3]
#> 
#>    Foo Bar Baz
#> 1   Aa  Bb  Cc
#> 2   Dd  Ee  Ff
#> 3   Gg  Hh  Ii
#> 4 Jj88  Kk  Ll
#> 5       Uu  Ii
#> 6   Hh  Ii   h
```

### Test Results

``` r
library(docxtractr)
library(testthat)

date()
#> [1] "Mon Aug 24 13:36:23 2015"

test_dir("tests/")
#> testthat results ========================================================================================================
#> OK: 0 SKIPPED: 0 FAILED: 0
#> 
#> DONE
```

### Code of Conduct

Please note that this project is released with a [Contributor Code of Conduct](CONDUCT.md). By participating in this project you agree to abide by its terms.
