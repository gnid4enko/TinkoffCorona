### 0) Get raw data from URL
require("rjson")
x <- readLines("https://index.tinkoff.ru/corona-index/papi/charts?region=all")
y <- fromJSON(x)

### 1) The data for consumer activity (daily)
## 1.1) Get start date and format it as date
strt <- y[1]$total$start
strt <- as.Date(paste0(substr(strt,7,10),"-",substr(strt,4,5),"-",substr(strt,1,2)))
## 1.2) Get aggregated data for consumer activity
z1 <- as.data.frame(t(as.data.frame(y[1]$total$points))); row.names(z1) <- c()
z1$add <- seq.Date(strt, strt+nrow(z1)-1, by = "day") ## add the dates
names(z1) <- c("Индекс","Доля онлайн-платежей","Потреб. активность", "Дата")
z1 <- subset(z1, select=c("Дата","Индекс","Доля онлайн-платежей","Потреб. активность"))
## 1.3) Get detailed data (consumer activity by categories)
i=1
while (i < length(y[2]$categories)) {
  zz <- data.frame(y[2]$categories[i][[1]]$points) # <get data frames for each category>
  names(zz) <- names(y[2]$categories[i]) # <add column names for each category>
  z1 <- cbind(z1, zz) # <add the data frame for the new category to the aggregated data frame>
  i = i+1
  rm(zz)
}

### 2) The data for business activity (weekly)
## 2.1) Get start date and format it as date
strt <- y[3]$businessTotal$start
strt <- as.Date(paste0(substr(strt,7,10),"-",substr(strt,4,5),"-",substr(strt,1,2)))
## 2.2) Get aggregated data for business activity
z2 <- as.data.frame(as.data.frame(y[3]$businessTotal$points)); row.names(z2) <- c()
nrrr <- z1[nrow(z1),1]-strt # <number of days since strt (for z2) till the end (for z1)>
z2$add1 <- seq.Date(strt, strt+nrrr-6, by = "week")
z2$add2 <- seq.Date(strt+6, strt+nrrr, by = "week")
names(z2) <- c("Обороты бизнеса","Начало недели","Конец недели")
z2 <- subset(z2, select=c("Начало недели","Конец недели","Обороты бизнеса"))
## 2.3) Get detailed data (business activity by categories)
i=1
while (i < length(y[4]$businessCategories)) {
  zz <- data.frame(y[4]$businessCategories[i][[1]]$points) # <get data frames for each category>
  names(zz) <- names(y[4]$businessCategories[i]) # <add column names for each category>
  z2 <- cbind(z2, zz) # <add the data frame for the new category to the aggregated data frame>
  i = i+1
  rm(zz)
}

### 3) Write the data to Excel (two sheets)
require("openxlsx")
path <- paste0("C:/Users/", Sys.info()[["user"]], "/Downloads/tinkoff_data.xlsx") # <set path for the download>
dfs <- list("potreb" = z1, "business" = z2) # <"potreb" & "business" are sheet names>
write.xlsx(dfs, file = path, row.names=F)

### 4) Do formatting (manupulate with Excel from R with the help of "openxlsx" library)
wb <- openxlsx::loadWorkbook(path) # <open "tinkoff_data.xlsx">
modifyBaseFont(wb, fontSize = 11, fontName = "Arial Narrow")
bold <- createStyle(textDecoration = "Bold", wrapText = TRUE, halign = "center", valign = "center") # <style for the 1st row>
d <- createStyle(wrapText = F, halign = "center", valign = "center", numFmt="DATE") # <style for dates>
simple <- createStyle(wrapText = F, halign = "center", valign = "center") # <style for the rest>
ColNo <- c(ncol(z1),ncol(z2)); RowNo <- c(nrow(z1),nrow(z2)) # <save the key parameters of the two data frames>
i=1
while (i < 3) { # <loop through the two sheets>
  freezePane(wb, i, firstActiveRow = 2, firstActiveCol = 1+i, firstRow = F, firstCol = F) # <freeze panes>
  setColWidths(wb, sheet = i, cols = 1:ColNo[i], widths = 15) # <set col width to 15 for all columns>
  addStyle(wb, i, style = d, rows=1:(RowNo[1]+1), cols=1:(i), gridExpand=T) # <add style for the 1st row>
  addStyle(wb, i, style = bold, rows=1, cols=1:ColNo[i], gridExpand=T) # <add style for dates>
  addStyle(wb, i, style = simple, rows=2:(RowNo[1]+1), cols=(i+1):ColNo[i], gridExpand=T) # <add style for the rest>
  i = i+1
}
saveWorkbook(wb, file = path, overwrite = TRUE) # <save the workbook back>
#
