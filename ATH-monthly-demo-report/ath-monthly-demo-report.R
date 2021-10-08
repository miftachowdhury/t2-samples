#install.packages('readr')
#install.packages('pivottabler')
#install.packages('basictabler')
#install.packages('flextable')
#install.packages('openxlsx')
#install.packages('officer')
#install.packages('stringi')

setwd('~/ATH_Rpt')

library(readr)
library(pivottabler)
library(basictabler)
library(flextable)
library(openxlsx)
library(officer)
library(stringr)
source("pvt_functions.R")
#library(flextable)

wordFileName = "Sep_Report.docx"

boroDict <- read_csv('boroughsByZIP.csv')
boroDict$ZIP <- as.character(boroDict$ZIP)

df <- dbGetQuery(db, statement = read_file('AtHomeTesting_DemoData_6.sql'))
noBoro <- df[df["BOROUGH"]=='', ]
hasZIP <- noBoro[noBoro["ZIP"]!='', ]

cat('Total N:', nrow(df),
    '\n N missing BOROUGH:', nrow(noBoro), 
    '\n  N missing BOROUGH but has ZIP:', nrow(hasZIP)
    )

df2 <- df
str(df2$`Created Date`)

#Attempt to recover missing boroughs using ZIP values
i <- 1
for(i in 1:nrow(hasZIP)) {
  id <- hasZIP[i, "ID"]
  zip <- hasZIP[i, "ZIP"]
  boro <- as.character(boroDict[boroDict["ZIP"]==zip,"Borough"])
  if (boro!="character(0)") df2[df2$ID==id, "BOROUGH"] <- boro
}

print(nrow(df2[df2["BOROUGH"]=='', ]))

df2['Age Group'] <- cut(df2$AGE, c(-1,14,24,34,44,54,64,74,120), 
                        labels = c('0-14', '15-24', '25-34', '35-44', '45-54', '55-64', '65-74', '75+'))

df2 <- df2[df2["BOROUGH"]!='',]
df2[df2["AT HOME TESTING"]=='', "AT HOME TESTING"] <- "(blank)"

vars <- c("Overall", "GENDER", "PREFERRED LANGUAGE", "Race", "Age Group")
ptList <- list()

i <- 1
for (i in 1:length(vars)) {
  
  getColPct <- function(pivotCalculator, netFilters, calcFuncArgs, format, fmtFuncArgs, baseValues, cell) {
    df2 <- pivotCalculator$getDataFrame("df2")
    netFilters$setFilterValues(variableName="AT HOME TESTING", values=NULL, action="replace")
    if(i>=2) netFilters$setFilterValues(variableName=calcFuncArgs$var, values=NULL, action="replace")
    df2Filtered <- pivotCalculator$getFilteredDataFrame(df2, netFilters)
    totalBoro <- nrow(df2Filtered)
    pct <- baseValues$N / totalBoro * 100
    value <- list()
    value$rawValue <- pct
    value$formattedValue <- pivotCalculator$formatValue(pct, format=format)
    return(value)
  }
  
  pt <- PivotTable$new()
  pt$addData(df2)
  pt$addColumnDataGroups("BOROUGH")
  ifelse (i>=2, {o <- list(isEmpty=FALSE, mergeSpace="dataGroupsOnly"); oTotal<-TRUE}, {o<-NULL; oTotal<-FALSE})
  pt$addRowDataGroups("AT HOME TESTING", header = '', outlineBefore = o, outlineTotal = oTotal)
  if (i>=2) pt$addRowDataGroups(vars[i], header = vars[i], addTotal=FALSE)
  pt$defineCalculation(calculationName="N", summariseExpression = "n()", visible=FALSE)
  pt$defineCalculation(calculationName = "Percentage", format="%.1f%%", basedOn="N", type="function", noDataCaption="-", calculationFunction=getColPct, calcFuncArgs=list(var=vars[i]))
  pt$evaluatePivot()
  pt$renderPivot(showRowGroupHeaders = TRUE)
  
  assign(paste0("pt", i), pt)
  ptList[[i]] <- pt
}

#Create function to export pivot tables to a Word doc (as flextables)
exportToWord <- function(wordFileName) {
  docx <- read_docx()
  i<-1
  for (i in 1:length(ptList)){
    pt <- ptList[[i]]

    heading <- ""
    if(i==1){heading<-"Overall demographic aggregation:";
    headers <- c(paste0(stringr::str_to_title(vars[i]), "/Borough"), "Bronx", "Brooklyn", "Manhattan", "Queens", "Staten Island", "Total")}
    if(i==2){heading<-"Demographic aggregation breakdown in Gender, Preferred Language, Race, and Age Group:"}
    if(i>=2){headers <- c(paste0(stringr::str_to_title(vars[i]), "/Borough"), "", "Bronx", "Brooklyn", "Manhattan", "Queens", "Staten Island", "Total")}
    
    tbl <- convertPvtTblToBasicTbl(pt, showRowGroupHeaders=FALSE)
    print(tbl$cells$getValue(r=4,c=5))
    
    tbl$cells$deleteColumn(tbl$columnCount)
    tbl$cells$deleteRow(1)
    tbl$cells$insertRow(1)
    tbl$cells$setRow(rowNumber = 1, cellTypes = "columnHeader", rawValues = headers)
    
    rows <- list()
    if(i>=2){

      j <- 1
      for (r in 1:tbl$rowCount){
        print(tbl$cells$getValue(r=r, c=1))
        if (!is.null(tbl$cells$getValue(r=r, c=1))) {rows[j]<-r; j <- j+1}
      }
      
      rows <- as.numeric(unlist(rows))
      newRows <- list()
      
      k <- 1
      for (m in 1:length(rows)){
        if(rows[m]<3) {
          newRows[[k]] <- rows[m]
          k <- k+1
        } else if (rows[m] != rows[m-1]+1) {
          newRows[[k]] <- rows[m]
          k <- k+1
          }
      }
      print(newRows)
      
      for(n in 1:length(newRows)){
        print(newRows[n])
        row <- newRows[n]
        tbl$mergeCells(rFrom=as.numeric(row), rTo=as.numeric(row), cFrom=1, cTo=2)
      }
      
    }

    ft <- tbl$asFlexTable()
    
    docx <- body_add_par(docx, heading)
    docx <- body_add_flextable(docx, value = ft)
    ps <- prop_section(page_size = page_size(orient = "landscape"))
    docx <- body_end_block_section(
      x = docx,
      value = block_section(property = ps))
  }
  print(docx, target=wordFileName)
  print('Word doc export completed')
}

exportToWord(wordFileName)
