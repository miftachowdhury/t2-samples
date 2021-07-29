#NYS Report Generator
#Last Modified: 2020-12-25
#Author: MAC

#tinytex::install_tinytex()
install.packages("readxl")
install.packages("openxlsx")
install.packages("xlsx")
install.packages("hash")
install.packages("qdap")
install.packages("qdapTools")
install.packages("lubridate")
install.packages("stringr")
install.packages("XLConnect")

Sys.setenv(JAVA_HOME='C:/Program Files/Java/jre1.8.0_271')

library(readxl)
library(openxlsx)
library(xlsx)
library(hash)
library(qdap)
library(qdapTools)
library(lubridate)
library(stringr)
library(XLConnect)

setwd("C:/NYS_Reports")

##Read in daily survey testing report file, create workbook, keep only "State template"
svyDate <- format(Sys.Date()-1, "%Y%m%d")
svyFile <- sprintf("Svy_Reports/Survey Testing Report_%svShared_(OVERVIEW + STATE + TAT-Pending).xlsx", svyDate)
svyWB <- openxlsx::loadWorkbook(svyFile)
rm_sheets <- openxlsx::getSheetNames(svyFile)
rm_sheets <- rm_sheets[rm_sheets!="State template"]
for(x in rm_sheets){
  removeWorksheet(svyWB, x)
}

rptDate <- gsub("/", "_", format(Sys.Date(), "%x"))
rptFile <- sprintf("C:/NYS_Reports/Output/NYC_SchoolCases_and_testing_%s.xlsx", rptDate)
svyWB <- removeFilter(svyWB, "State template")
openxlsx::setColWidths(svyWB, "State template", cols=10:51, widths=36.57)
openxlsx::saveWorkbook(svyWB, rptFile, overwrite=TRUE)

##Read in existing data
templ_data <- read_excel(svyFile, sheet="State template", col_names=T, skip=1)
templ_data <- templ_data[,10:52]
str(templ_data)

##Write existing data into NYS report file and save
svyWB <- XLConnect::loadWorkbook(rptFile, create=FALSE)
setStyleAction(wb, XLC$STYLE_ACTION.NONE)
XLConnect::clearRange(svyWB, 
                      sheet="State template",
                      coords=aref2idx("J3:AZ1600"))
XLConnect::writeWorksheet(svyWB, templ_data, "State template", startRow=3, startCol=10, header=F)
XLConnect::saveWorkbook(svyWB)

#################################################################################################
#################################################################################################

##Read in NYS template file
rpt <- read_excel(rptFile, col_names=TRUE, skip=1)
colnames(rpt)

##Read in case data
caseData <- sprintf("Data/Cases_%s.xlsx", rptDate)
df_cases <- read_excel(caseData, 1, col_names=TRUE, cell_cols("D:O"))
names(df_cases)[3]<-"DateTime"
names(df_cases)[8]<-"DBN"

##Read in BEDS Code values and map DBN to BEDS in Cases data frame
beds_dict <- read.csv("BEDS_Codes.csv", header=TRUE, stringsAsFactors = FALSE)
df_cases$BEDS <- lookup(df_cases$DBN, beds_dict$DBN, beds_dict$BEDS, missing=NA)

##Filter for only cases in the time window of interest
if (weekdays(Sys.Date())!="Monday"){
  dt_start <- ymd(Sys.Date()-1)+hms("15:45:00")
} else {
  dt_start <- ymd(Sys.Date()-3)+hms("15:45:00")
}
dt_end <- ymd(Sys.Date())+hms("15:45:00")
df_cases2<-df_cases[which(df_cases$DateTime>dt_start & df_cases$DateTime<=dt_end), ]

##Drop cases where BEDS Code is NA
df_omitNA <- df_cases2[!is.na(df_cases2$BEDS),]

##Count student cases by BEDS code
df_cases2S <- df_omitNA[which(df_omitNA$`Person Type`=="Student"),]
count_stud <- function(x){
  sum(df_cases2S$BEDS==x, na.rm=TRUE)
}
v_students <- sapply(rpt$`SED BEDS Code`, count_stud)

##Count employee cases by BEDS code
df_cases2E <- df_omitNA[which(df_omitNA$`Person Type`=="Employee"),]
count_emp <- function(x){
  sum(df_cases2E$BEDS==x, na.rm=TRUE)
}
v_employees <- sapply(rpt$`SED BEDS Code`, count_emp)

##Create a vector of today's date
v_date <- rep(mode="any", format(Sys.Date(), "%x"), times=nrow(rpt))

##Write counts to worksheet and save file
wb <- XLConnect::loadWorkbook(rptFile, create=FALSE)
setStyleAction(wb, XLC$STYLE_ACTION.NONE)
writeWorksheet(wb, v_date, 1, startRow=3, startCol=8, header=FALSE)
writeWorksheet(wb, v_students, 1, startRow=3, startCol=16, header=FALSE)
writeWorksheet(wb, v_employees, 1, startRow=3, startCol=18, header=FALSE)
XLConnect::saveWorkbook(wb)

#######################################################################
#           SAVE FILTERED CASE DATA (OPTIONAL STEP)                   #
# COMMENT OUT CODE (CTRL+SHIFT+C) IF FILTERED CASE DATA IS NOT NEEDED #
#######################################################################

cf_file <- sprintf("Cases_Filtered/Cases_filtered_%s.xlsx", rptDate)
cfWB <- openxlsx::createWorkbook()

sheetList <- list("with_NA", "no_NA", "Students_no_NA", "Employees_no_NA")
dfList <-  list(df_cases2, df_omitNA, df_cases2S, df_cases2E)

makeTable <- function(wb_sheet, df){
  openxlsx::addWorksheet(cfWB, sheetName = wb_sheet)
  openxlsx::writeDataTable(cfWB, sheet=wb_sheet, x=df, tableStyle="TableStyleMedium2")
}

mapply(makeTable, sheetList, dfList)

openxlsx::saveWorkbook(cfWB, file=cf_file, overwrite=TRUE)     

#####################################################################
#                     END FILTERING CASE DATA                       #
#####################################################################
