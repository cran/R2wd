wdTable <-
    function (data, caption = "", bookmark = NULL, pointsize = 9,
              padding = 5, autoformat = 1, align = c("l", rep("r", ncol(data))),
              wdapp = .R2wd)
{
    wdtab <- wdapp[["ActiveDocument"]][["Tables"]]
    wdsel <- wdapp[["Selection"]]
    wddoc <- wdapp[["ActiveDocument"]]
    wdsel$TypeParagraph()
    wdInsertBookmark("R2wdEndmark")
    bookmarkcounter <- wddoc[["Bookmarks"]][["Count"]]
    wdsel$MoveUp()
    nr <- nrow(data)
    nc <- ncol(data)
    out <- matrix("", nrow = nr + 1, ncol = nc + 1)
    out[1 + (1:nr), 1 + (1:nc)] <- as.matrix(data)
    out[1, 1 + (1:nc)] <- colnames(data)
    out[1 + (1:nr), 1] <- row.names(data)
    tt<-paste(apply(out,1,paste,collapse="\t"),collapse="\n")
    wdsel[['Text']]<-tt
    tab<-wdsel[['Range']]$ConvertToTable(1,nr+1,nc+1)
    tab$AutoFormat(autoformat)
    if (as.numeric(.R2wd[['Version']])>10)
    {
        tab[["Rows"]][["Height"]] <- pointsize + padding
        tab[["Rows"]][["HeightRule"]] <- 2
        tab[["Range"]][["Cells"]][["VerticalAlignment"]] <- 1
    }
    tab$AutoFitBehavior(1)
    tab[["Range"]]$Select()
    wdsel[["Font"]][["Size"]] <- pointsize
    for (i in 1:length(align)) {
        tab[["Columns"]]$Item(i)$Select()
        wdsel[["ParagraphFormat"]][["Alignment"]] <- c(l = 0,
                                                       c = 1, r = 2)[align[i]]
    }
    tab[["Range"]]$Select()
    caption <- paste(" ", caption, sep = "")
    wdsel$InsertCaption("Table", caption, "", 1, 0)
    if (is.null(bookmark))
        bookmark <- paste("Table", bookmarkcounter + 1, sep = "")
    tab[["Range"]]$Select()
    wdInsertBookmark(bookmark)
    wddoc[["Bookmarks"]]$Item(bookmarkcounter)$Select()
    wdGoToBookmark("R2wdEndmark")
    return()
}
