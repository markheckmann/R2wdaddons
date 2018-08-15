file <- system.file("inst/template.docx", package = "R2wdaddons")
wdGet2(file)

# does the search string exist?
wdSearchString("[placeholder 1]")
