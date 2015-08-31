wdGet("inst/test.docx")
d <- mtcars[1:10, 1:6]
tbl <- wdReplaceTextByTable("Some text to search for!", d)
