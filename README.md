
<!-- README.md is generated from README.Rmd. Please edit that file -->

# image2xlsx

<!-- badges: start -->

![](https://img.shields.io/badge/cool-useless-green.svg)
![](https://img.shields.io/badge/status-rough-red.svg)
<!-- badges: end -->

`image2xlsx` provides a single function to turn an image into an excel
(xlsx) spreadsheet.

It stands on the shoulders of the following giants:

  - [openxlsx](https://cran.r-project.org/package=openxlsx)
  - [magick + imagemagick](https://cran.r-project.org/package=magick)
  - [dplyr](https://cran.r-project.org/package=dplyr)
  - [tidyr](https://cran.r-project.org/package=tidyr)
  - [raster](https://cran.r-project.org/package=raster)

## Future/ToDo

Patches welcomed\!

  - Deal with transparency
      - Get `magick` to flatten transparency to a user-specified
        background colour?
  - Place the image at a desired location
      - Currently image is placed with top left corner at `A1` on the
        sheet.
      - Be nice to place it offscreen so that it’s hidden when you first
        open the sheet
  - Animation using `openxlsx::conditionalFormatting()`
      - `gifanim2xlsx()`

## Installation

You can install from
[GitHub](https://github.com/coolbutuseless/image2xlsx) with:

``` r
# install.packages("devtools")
devtools::install_github("coolbutuseless/image2xlsx")
```

## Rlogo

``` r
library(image2xlsx)
image2xlsx("working/RStudio.png", "man/figures/rlogo.xlsx")
```

A screenshot of `rlogo.xlsx` opened in LibreOffice

<img src="man/figures/rlogo.png" width="100%" />

## Bender Bending Rodríguez

``` r
library(image2xlsx)
image2xlsx("working/bender.jpg", "man/figures/bender.xlsx")
```

A screenshot of `bender.xlsx` opened in LibreOffice

<img src="man/figures/bender.png" width="100%" />

## Grappling Hook\!

``` r
library(image2xlsx)
image2xlsx("working/mabel.png", "man/figures/mabel.xlsx")
```

A screenshot of `mabel.xlsx` opened in LibreOffice

<img src="man/figures/mabel.png" width="100%" />
