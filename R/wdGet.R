wdGet <-
    function (filename = NULL, path = getwd(), doc = NULL, visible = TRUE)
{
    require(rcom)
    comRegisterServer()
    tmp <- comGetObject("Word.Application")
    if (is.null(tmp))
        tmp <- comCreateObject("Word.Application")
    if (is.null(filename)) { if (is.null(tmp[["ActiveDocument"]]))
                                 tmp[["Documents"]]$Add()}
    else {
        if (tmp[["Documents"]][["Count"]]==0 || !(tmp[["ActiveDocument"]][["Name"]] == filename)) {tt<-tmp[["Documents"]]$Open(paste(path, filename, sep = "/"))
                if (!(tmp[["ActiveDocument"]][["Name"]] == filename))
            if (is.null(tt)) warning(paste("File",paste(path, filename, sep = "/"),"not found"))}
      }
    if (visible)
        tmp[["visible"]] <- TRUE
    .R2wd <<- tmp
    invisible(tmp)
}
