wdApplyTemplate <-
function(filename,wdapp=.R2wd){
    wdapp[['ActiveDocument']]$ApplyTemplate(filename)
    return()
}

