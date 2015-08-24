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

-   Version 0.0.1.9001 released - pre-CRAN flight check
-   Version 0.0.1.9000 released - read from URL
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
#> [1] '0.0.1.9001'

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

# url 

budget <- read_docx("http://www.anaheim.net/docs_agend/questys_pub/MG41925/AS41964/AS41967/AI44538/DO44539/1.DOCX")

docx_tbl_count(budget)
#> [1] 2

docx_describe_tbls(budget)
#> Word document [http://www.anaheim.net/docs_agend/questys_pub/MG41925/AS41964/AS41967/AI44538/DO44539/1.DOCX]
#> 
#> Table 1
#>   total cells: 24
#>   row count  : 6
#>   uniform    : likely!
#>   has header : unlikely
#> 
#> Table 2
#>   total cells: 28
#>   row count  : 4
#>   uniform    : likely!
#>   has header : unlikely

docx_extract_tbl(budget, 1)
#> Source: local data frame [5 x 4]
#> 
#>                                      Short-term Portfolio Long-term Portfolio Total Portfolio Values
#> 1 Portfolio Balance (Market Value) *       $  123,651,911       $ 294,704,136          $ 418,356,047
#> 2                    Effective Yield               0.16 %              1.42 %                 1.05 %
#> 3             Avg. Weighted Maturity              11 Days           2.4 Years              1.7 Years
#> 4                       Net Earnings        $      18,470      $      350,554         $      369,024
#> 5                        Benchmark**               0.02 %              0.41 %                 0.27 %

docx_extract_tbl(budget, 2) 
#> Source: local data frame [3 x 7]
#> 
#>                        Amount of Funds (Market Value)  Maturity Effective Yield Interpolated Yield
#> 1 Short-Term Portfolio                  $ 123,651,911   11 days          0.16 %             0.01 %
#> 2  Long-Term Portfolio                  $ 294,704,136 2.4 years          1.42 %             0.41 %
#> 3      Total Portfolio                  $ 418,356,047 1.7 years          1.05 %             0.27 %
#> Variables not shown: Total Return Monthly (chr), Total Return Annual (chr)

# three tables
doc3 <- read_docx(system.file("examples/data3.docx", package="docxtractr"))

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

# no tables
none <- read_docx(system.file("examples/none.docx", package="docxtractr"))

docx_tbl_count(none)
#> [1] 0

# wrapping in try since it will return an error
# use docx_tbl_count before trying to extract in scripts/production
try(docx_describe_tbls(none))
#> No tables in document
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
#> [1] "Mon Aug 24 19:59:01 2015"

test_dir("tests/")
#> testthat results ========================================================================================================
#> OK: 10 SKIPPED: 0 FAILED: 0
#> 
#> DONE
```

### Code of Conduct

Please note that this project is released with a [Contributor Code of Conduct](CONDUCT.md). By participating in this project you agree to abide by its terms.
