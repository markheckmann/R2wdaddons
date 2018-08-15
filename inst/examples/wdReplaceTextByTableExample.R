file <- system.file("inst/template.docx", package = "R2wdaddons")
wdGet2(file)

# add a table at placeholder
d <- mtcars[1:10, 1:6]
tbl <- wdReplaceTextByTable("[placeholder 3]", d)

