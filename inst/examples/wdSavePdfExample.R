
file <- system.file("inst/template.docx", package = "R2wdaddons")
wdGet2(file)

file <- tempfile(fileext=".pdf")  # file in temp directory
wdSaveAsPdf(file)

