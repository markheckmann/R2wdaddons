file <- system.file("inst/template.docx", package = "R2wdaddons")
wdGet2(file)

# image file to put in document
img <- system.file("inst/sky.jpeg", package = "R2wdaddons")

# replace and rescale first image
shp <- wdReplaceTextByImage("[placeholder 1]", img)
wdScaleImage(ishp = shp, width=200, units="px")
wdAddImageCaption(ishp = shp, title = "Image is 200 pixels wide")
