# example of how to used the package

file <- system.file("inst/template.docx", package = "R2wdaddons")
wdGet2(file)

# image file to put in document
img <- system.file("inst/sky.jpeg", package = "R2wdaddons")

# replace and rescale first image
shp <- wdReplaceTextByImage("[placeholder 1]", img)
wdScaleImage(ishp = shp, width=200, units="px")
wdAddImageCaption(ishp = shp, title = "Image is 200 pixels wide")

# replace and rescale second image
shp <- wdReplaceTextByImage("[placeholder 2]", img)
wdScaleImage(ishp = shp, width=80, units="widthpercent")
wdAddImageCaption(ishp = shp, title = "Image width is 80% of text width")

# add a table at placeholder
d <- mtcars[1:10, 1:6]
tbl <- wdReplaceTextByTable("[placeholder 3]", d)

# replace placeholder text
wdReplaceTextByText("[placeholder 4]", "SOME REPLACEMENT TEXT")

# save as temp file
file <- tempfile(fileext=".docx")
wdSave2(file)


# # count number of shapes
# ishps <- .R2wd[["Selection"]][["Range"]][["InlineShapes"]]
# ishps[["Count"]]
