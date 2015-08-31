# wdSearchTextGetRange <- function(find = "", wdapp = .R2wd) 
# {
#   success <- wdSearchString(find, wdapp=wdapp)    # search string in document and select
#   
#   # if text was not found
#   if (!success) {
#     if (warn)
#       warning("The find text was not found in document", call. = FALSE)
#     return(FALSE)
#   } 
#   
#   # get range of selection
#   r <- wdapp[["Selection"]][["Range"]]            # get range of selection
#   invisible(r)
# }





