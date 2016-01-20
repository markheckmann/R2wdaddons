

# # 0=single, 1=one and a half, 2=double spacing
# replaceCaptionText <- function(i, ishp=NULL, text, wdapp = .R2wd, linespacing=1)
# {
#   wddoc <- wdapp[["ActiveDocument"]]
#   ishps <- wddoc[["InlineShapes"]]        # get inlineshapes
#   if (is.null(ishp))
#     ishp <- ishps$Item(i)                 # get ith inline shape 
#   
#   ishp$Select()
#   wdsel <- wdapp[["Selection"]]
#   wdsel$Previous(Unit=5)$Select()         # select caption
#   r <- wdapp[["Selection"]][["Range"]]    # get range of selection
#   wdsel <- wdapp[["Selection"]]
#   pformat <- wdsel[["ParagraphFormat"]]
#  # style <- r[["Style"]]
#   r[["Text"]] <- text
#   pformat[["LineSpacingRule"]] <- 2       # 0=single, 1=one and a half, 2=double spacing
#   #r[["Style"]] <- style
#   invisible(r)                             # Return Range object
# }