# replacement of buggy function in R2wd


#' Convert path slashes to path with double backslashes.
#' 
#' RDCOMClient / the COM interface seems to only support this type of path.
#' 
#' @param file A file path.
#' @return Path with double backslashes instead of slahes. 
#' @export
#' 
to_win_path <- function(file) 
{
  stringr::str_replace_all(file, "/", "\\\\")
}



#' wdGet2
#' 
#' There is a bug in \code{R2wd::wdGet} as path slashes are not replaced by double backslashes.
#' This may cause an error when using RCDOMClient on some plattforms. This version simply 
#' does the replacement.
#' 
#' @param x A filename.
#' @param path A Path to the file, if not set otherwise the working directory is used.
#' @param method Use \code{"rcom"} (default) or \code{"RDCOMClient"}.
#' @param visible If the Applciation should be visible.
#' 
#' @return Store invisible file handle named in global environment.
#' @example /inst/examples/wdGet2Save2Example.R
#' @export
#' 
wdGet2 <- function (filename = NULL, 
                    path = getwd(), 
                    method = "RDCOMClient", 
                    visible = TRUE) 
{
  # check if filename is absolute path 
  # if not combine with path (defaults to working directory)
  # to create full path to file
  path_is_absolute <- R.utils::isAbsolutePath(filename)
  if (!path_is_absolute)
    filename <- R.utils::filePath(path, filename)
  
  # does the file exist? with throw an error if not.
  file_exists(filename)
  
  # is it a .docx file?
  if (! file_is_docx(filename))
    stop("The file ", filename, " is no .docx file.", call. = FALSE)
  
  #   # replace by eympty string if is NA or NULL
  # if (is.null(path))
  #   path <- ""
  # if (is.na(path))
  #   path <- ""
  
  # code copied from r2wd
  if (method == "rcom") {
    if (!require(rcom)) {
      warning("The package rcom is unavailable.")
      if (require(RDCOMClient)) {
        warning("Using RDOMClient package instead of rcom.")
        client <- "RDCOMClient"
      }
      else {
        stop("Neither rcom or RDCOMClient packages are installed.")
        client <- "none"
      }
    }
    else {
      client <- "rcom"
      if ("package:RDCOMClient" %in% search()) {
        warning("\nUsing rcom package. Detaching RDCOMClient package to avoid conflicts.")
        try(detach("package:RDCOMClient"))
      }
    }
  }
  if (method == "RDCOMClient") {
    if (!require(RDCOMClient)) {
      stop("The package RDCOMClient is unavailable. \n \n\t\tTo install RDCOMClient use:\n \n\t\tinstall.packages('RDCOMClient' repos = 'http://www.omegahat.org/R')")
    }
    client <- "RDCOMClient"
    if ("package:rcom" %in% search()) {
      warning("Using RDCOMClient package. Detaching rcom package to avoid conflicts.")
      try(detach("package:rcom"))
    }
  }
  
  # create handle to word application 
  switch(client, rcom = {
    wdapp <- comGetObject("Word.Application")
    if (is.null(wdapp)) wdapp <- comCreateObject("Word.Application")
  }, RDCOMClient = {
    wdapp <- COMCreate("Word.Application")
  }, none = stop("no client"))
  
  # set application to visible if prompted
  if (visible) 
    wdapp[["visible"]] <- TRUE
  
  # create word document if no filename has been supplied
  # and no document is opened
  if (is.null(filename)) {
    if (wdapp[["Documents"]][["Count"]] == 0) 
      wdapp[["Documents"]]$Add()
  }
  else {
    wddocs <- wdapp[["Documents"]]
    found <- FALSE
    if (wddocs[["Count"]] > 0) {
      for (i in 1:wddocs[["Count"]]) {
        wddoc <- wddocs$Item(i)
        if (wddoc[["Name"]] == filename) {
          wddoc$Activate()
          found <- TRUE
          break
        }
      }
    }
    # if document cannot be found in open word docs create it
    if (!found) {
      # if (path == "") {   # when file path is dropped
      #   file <- filename  # I assume that the filename contains the full path
      # } else {
      #   file <- paste(path, filename, sep="/")   # append filename to given path 
      # } 
      filename <- to_win_path(filename)          # convert to windows file path
      wddoc <- try(wdapp[["Documents"]]$Open(filename))
      if (class(wddoc) == "try-error" | is.null(wddoc)) {
        if (wddocs[["Count"]] == 0) 
          wdapp$Quit()
        print(paste("File", file, "not found"))
      }
    }
  }
  .R2wd <<- wdapp
  invisible(wdapp)
}


#' wdSave2
#' 
#' There is a bug in \code{R2wd::wdSave} as path slashes are not replaced by double backslashes.
#' This may cause an error when using RCDOMClient on some plattforms. This version simply 
#' does the replacement.
#' 
#' @param file File name (if missing, Word will ask).
#' @param path Path to file. Current working directory is used as default.
#' @param wdapp The handle to the Word Application (usually not needed).
#' @example /inst/examples/wdGet2Save2Example.R
#' @export
#' 
wdSave2 <- function (file = NULL, path = getwd(), wdapp = .R2wd) 
{
  wddoc <- wdapp[["ActiveDocument"]]
  
  if (is.null(file)) {
    wddoc$Save()
  } else {
    if (!R.utils::isAbsolutePath(file)) {   # make absolute if file is relative path
      file <- file.path(path, file)
    }
    file <- to_win_path(file)   # convert slashes to backslashes 
    wddoc$SaveAs(file)
  }
  invisible(wdapp)
}


