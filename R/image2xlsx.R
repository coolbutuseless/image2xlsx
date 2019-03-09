
globalVariables(c('colour'))

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#' Convert an image to an excel spreadsheet
#'
#' @param image_filename filename of an image supported by \code{magick}
#' @param xlsx_filename filename of the output excel spreadsheet.
#' @param height Height of output i.e. number of vertical cells the image should occupy. Default: 40
#' @param width Width of output. Default: NULL - scale to height and keep aspect ratio
#' @param xoffset,yoffset Starting cooridinates for top-left of image. Default (0, 0)
#' @param overwrite Overwrite the excel spreadsheet if it already exists? Default: FALSE
#'
#' @return invisibly returns the excel workbook created by \code{openxlsx}
#'
#' @import dplyr
#' @import magick
#' @import openxlsx
#' @importFrom raster as.raster
#' @importFrom tidyr gather
#'
#' @export
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
image2xlsx <- function(image_filename, xlsx_filename = NULL, height = 40, width = NULL, xoffset = 0, yoffset = 0, overwrite = FALSE) {

  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  # load the image
  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  image <- magick::image_read(image_filename) %>%
    magick::image_scale(magick::geometry_size_pixels(height = height, width = width)) %>%
    magick::image_data()

  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  # Turn it into a raster object - this will get us hex colour codes at
  # every location
  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  nrows  <- nrow(image[1,,])
  ncols  <- ncol(image[1,,])
  image <- raster::as.raster(as.integer(image)/255)

  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  # Turn the cell data into a data.frame (ready for openxslx)
  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  image <- as.character(image)
  dim(image) <- c(nrows, ncols)
  image <- as.data.frame(image)

  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  # Figure out the hex colour code for every cell
  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  suppressWarnings({
    colours <- image %>%
      tidyr::gather(row, colour) %>%
      group_by(row) %>%
      mutate(col = seq(n())) %>%
      ungroup() %>%
      mutate(
        row = as.integer(substr(row, 2, 20)),
        colour = substr(colour, 1, 7) # just drop transparency
      )
  })


  colours <- colours %>% mutate(
    col = col + xoffset,
    row = row + yoffset
  )


  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  # Create a blank data.frame the same size as the colours data.frame
  # i.e the cells of this spreadsheet will only contain blanks, and then I'll
  # set the "fill" colour to create the pixels
  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  min_cols <- min(colours$col)
  max_cols <- max(colours$col)
  max_rows <- max(colours$row)
  blank <- as.data.frame(rep(
    list(rep("", max_rows)),
    max_cols
  ))

  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  # 1. Create a workbook
  # 2. Add a worksheet
  # 3. Add blank data of the required size to the worksheet
  # 4. set the column widths so that the aspect ratio of final image is OK
  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  sheet_name <- basename(image_filename)
  wb <- createWorkbook("image2xlsx")
  addWorksheet(wb, sheet_name, gridLines = TRUE)
  writeData(wb, sheet = 1, blank, colNames = FALSE)
  setColWidths(wb, sheet=1, cols=seq(min_cols, max_cols), widths = 2)

  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  # Set the style of the cells (in order to define the colour)
  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  for (col_df in split(colours, colours$colour)) {
    cell_style <- createStyle(fgFill = col_df$colour[[1]])
    addStyle(wb, sheet = 1, cell_style, rows = col_df$row, cols = col_df$col, gridExpand = FALSE)
  }

  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  # Save the workbook
  #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  if (!is.null(xlsx_filename)) {
    openxlsx::saveWorkbook(wb, xlsx_filename, overwrite = TRUE)
  }


  invisible(wb)
}

