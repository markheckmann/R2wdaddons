wdGet2("inst/test.docx")

# save as temp file
file <- tempfile(fileext=".docx")
wdSave2(file)

# Fails: Does not yet notice that the path is absolute
# wdGet2(file)

