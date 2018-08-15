file <- system.file("inst/template.docx", package = "R2wdaddons")
wdGet2(file)
wdAddImageCaption(i = 1, title = "My image")
