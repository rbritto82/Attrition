library(xlsx)
library(tidyverse)

#reading file and changing columns' names
egiPersonel <- tbl_df(read.xlsx("./EGI developers detail.xlsx", sheetIndex=1, stringsAsFactors = FALSE))
sdpNetworks <- tbl_df(read.xlsx("./Kerstin_2014-2016_all_extractwithSDPNetworks.xlsx", sheetIndex=1, stringsAsFactors = FALSE))

names(egiPersonel) <- c("ID", "Name", "Date", "Hours", "HourType", "Activity")

sdpNetworks <- sdpNetworks[!is.na(sdpNetworks$node),]

#selecting only training hours and removing lines with zero hours
egiPersonel2 <- egiPersonel[grep("Training", egiPersonel$HourType), ]

#selecting only productive hours regarding PCs and trnasfer cost, and removing zero hours
egiPersonel3 <- egiPersonel[grep("Productive.", egiPersonel$HourType), ]

#Getting only the productive hours from SDP networks
egiPersonel4 <- egiPersonel3[(egiPersonel3$Activity %in% sdpNetworks$description), ]



#combining training and productive hours
egiPersonel3 <- rbind(egiPersonel2, egiPersonel3)

#removing lines with zero variables
egiPersonel3 <- egiPersonel3[egiPersonel3$Hours > 0, ]

#Create one factor for all pcs
egiPersonel3$Activity <- gsub("CHTT-036|CIN-[0-9][0-9][0-9]|ETIE-121|FTML-066|IDEA-173|
                              IND-176|TELSTRA-[0-9][0-9][0-9]|VOICE-[0-9][0-9][0-9] CEB|
                              VOICE-[0-9][0-9][0-9] CEB RR|WIND-206|ZAINK-[0-9][0-9][0-9]|TURK-[0-9][0-9][0-9]", "PC", egiPersonel3$Activity)

#Fetching month and year from the date column and inserting both informations
egiPersonel3 <- mutate(egiPersonel3, Day = format(Date, "%d"), Month = format(Date, "%m"), 
                       Year = format(Date, "%Y"), MonthYear = paste(as.character(Month),as.character(Year), sep = ""))
#Summary of number of developers, total hours and man-month per month  
monthlyData <- egiPersonel3 %>% 
  group_by(MonthYear, HourType, Activity) %>%
  summarise(total = sum(Hours), numDev = NROW(unique(Name)), manHours = sum(Hours)/NROW(unique(Name)))
#Writing summary in a spreadsheet
write.xlsx(monthlyData, file="Hours developers.xlsx", sheetName="Monthly")

#Summary of number of developers, total hours and man-month per year  
yearlyData <- monthlyData %>% mutate(Month = substr(MonthYear, 1,2), Year = substr(MonthYear, 3,6)) %>% 
  group_by(Year, HourType, Activity) %>%
  summarise(meanTotalHours = mean(total), meanNumDev = mean(numDev), meanManMonth = mean(manHours))
#Writing summary in a spreadsheet
write.xlsx(yearlyData, file="Hours developers.xlsx", sheetName="Yearly", append=TRUE)
