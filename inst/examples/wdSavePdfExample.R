wdGet2("inst/test.docx")          # open word doc in package

file <- tempfile(fileext=".pdf")  # file in temp directory
wdSaveAsPdf(file)
