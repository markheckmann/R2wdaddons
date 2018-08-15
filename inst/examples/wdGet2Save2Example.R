file <- system.file("inst/template.docx", package = "R2wdaddons")
wdGet2(file)

# save as temp file
file <- tempfile(fileext=".docx")
wdSave2(file)

# Fails: Does not yet notice that the path is absolute
# wdGet2(file)

