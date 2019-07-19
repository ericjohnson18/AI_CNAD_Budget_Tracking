data <- read.csv("~/Downloads/Forecast Details - Campaign Program.csv")
budgetdata <- read.csv("~/Documents/Budget Tracking/Budget Totals.csv")
data$Date_Pulled <- Sys.Date()
names(data) <- c("Function", "Budget_Name", "Campaign", "Marketing_Activity", 
                 "Marketing_Activity_Type", "Line_Item", "Activity_Description", 
                 "Quarter", "Spent", "Forecast", "Date_Pulled")
data$Campaign <- substr(data$Campaign, 11, 30)
data$Campaign[data$Campaign=='Cloud-Native App Dev'] <- 'Cloud Native App Dev'
data$Negative_Forecast_Tag <- ifelse(data$Forecast<0,"Negative Forecast","Positive Forecast")
data$Negative_Spent_Tag <- ifelse(data$Spent<0,"Negative Spent","Positive Spent")
data$OpenShift_Tag <- ifelse(grepl("openshift|ocp|containers|container", data$Line_Item, 
                                    ignore.case = TRUE)=="TRUE", "OpenShift", 
                              (ifelse(grepl("openshift|ocp|containers|container", data$Activity_Description,
                                    ignore.case = TRUE)=="TRUE","OpenShift","Middleware")))
data <- data[,c(1,2,3,4,5,6,7,8,9,10,11,12,13,14)]
duplicate <- data[duplicated(data[,1:7]),]
unique <- data[!duplicated(data[,1:7]),]
duplicatenotzero <- duplicate[!(duplicate$Spent==0 & duplicate$Forecast==0),]
uniquenotzero <- unique[!(unique$Spent==0 & unique$Forecast==0),]
allnotzero <- rbind(duplicatenotzero, uniquenotzero)
duplicatezero <- duplicate[(duplicate$Spent==0 & duplicate$Forecast==0),]
uniquezero <- unique[(unique$Spent==0 & unique$Forecast==0),]
allzero <- rbind(duplicatezero, uniquezero)
allzeroduplicate <- allzero[duplicated(allzero[,1:7]),]
allzerounique <- allzero[!duplicated(allzero[,1:7]),]
finaldata <- rbind(allzerounique,allnotzero)
xlsxfilename <- paste0("~/Documents/Budget Tracking/Budget Tracking Data R.xlsx")
library(openxlsx)
newdata <- loadWorkbook("~/Documents/Budget Tracking/Budget Tracking Data R.xlsx")
addWorksheet(newdata, Sys.Date())
writeData(newdata, sheet=as.character(Sys.Date()), x=finaldata)
freezePane(newdata, sheet = as.character(Sys.Date()), firstRow = TRUE)
saveWorkbook(newdata, xlsxfilename, overwrite = TRUE)