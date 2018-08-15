# utility functions

#' Is file a MS Word .docx file
#' 
file_is_docx <- function(f)
{
  ext <- tools::file_ext(f)
  tolower(ext) == "docx"  
}
