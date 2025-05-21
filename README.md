
# psyprojmarks

<!-- badges: start -->
<!-- badges: end -->

This package provides some tools for extracting grades and comments from NTU Psychology UG project marksheets and comparing the grades and comments of pairs of markers.

## Installation

This package is probably never going to be on CRAN because it is just
something for people in the Psychology department in NTU. As such, it
should be installed from GitHub. There are multiple ways of doing this,
but one common way is as follows:

``` r
remotes::install_github("mark-andrews/psyprojmarks")
```

## Examples

Here are examples of the main commands now:

``` r
library(psyprojmarks)

# Get grades and comments from the marksheet of one marker
extract_grades("UGPROJ25_666_1.docx")

# Get grades and comments from the marksheets of both markers
# and put the grades and comments of each marker for each criterion side by side
join_grades("UGPROJ25_666_1.docx", "UGPROJ25_666_2.docx") 

# Write the data frame produced by the previous command to an Excel workbook.
# The workbook is formatted for readability:
# - The match or mismatch of the two markers is indicated by green tick or red cross, respectively
# - The columns with the comments are wide and text wrapped 
# View this workbook in Excel on desktop (or Libreoffice etc) rather than online Word.
join_grades("UGPROJ25_666_1.docx", "UGPROJ25_666_2.docx") |> 
  write_joined_grades(file = 'UGPROJ25_666_agreement.xlsx')
```

