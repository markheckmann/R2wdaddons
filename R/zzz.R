.onLoad <- function(libname, pkgname) {
  
  if (!require(RDCOMClient)) 
    warning("\nWARNING: \nThe package RDCOMClient is unavailable. For this package to run, the RDCOMClient has to be installed. This early version does not support rcom.")

      
}