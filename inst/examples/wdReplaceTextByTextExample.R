file <- system.file("inst/template.docx", package = "R2wdaddons")
wdGet2(file)

# replace placeholder text
wdReplaceTextByText("[placeholder 4]", "SOME REPLACEMENT TEXT")
