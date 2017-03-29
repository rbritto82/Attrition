library(xlsx)
library(tidyverse)
library(lubridate)
library(stringr)

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
egiPersonel3 <- egiPersonel3[(egiPersonel3$Activity %in% sdpNetworks$description), ]

#combining training and productive hours
egiPersonel3 <- rbind(egiPersonel2, egiPersonel3)

#removing lines with zero variables
egiPersonel3 <- egiPersonel3[egiPersonel3$Hours > 0, ]

#Getting the start and dates for each developer and ordering by start date, end date and name
startEndDev <- egiPersonel3 %>% 
  group_by(Name) %>%
  summarise(startDate = min(Date), endDate = max(Date)) 
  
startEndDev <- startEndDev[order(startEndDev$startDate, startEndDev$endDate, startEndDev$Name), ]

#Getting months in the period and setting the date for the last day of each month
months <- seq(min(startEndDev$startDate), max(startEndDev$endDate), by="month") - 1
months <- months %m+% months(1)

#Calculating attrition, new developers, 6-month developers, 12 month developers, total
totalDev <- numeric(length = length(months))
attritionDev <- numeric(length = length(months))
sixMonthDev <- numeric(length = length(months))
twelveMonthDev <- numeric(length = length(months))
newDev <- numeric(length = length(months))
devPerMonth <- list()
devPerMonth[[1]] <- startEndDev$Name[(startEndDev$startDate <= months[1]) & (startEndDev$endDate >= months[1])]
totalDev[1] <- sum(table(devPerMonth[[1]]))
newDev[1] <- totalDev[1] - sixMonthDev[1] - twelveMonthDev[1]
sixthMonth <- devPerMonth[[1]]
twelvethMonth <- devPerMonth[[1]]
index <- vector(length = length(startEndDev$Name))

for(i in 2:length(months)){
  devPerMonth[[i]] <- startEndDev$Name[ (startEndDev$startDate <= months[i]) & (startEndDev$endDate >= months[i]) ]
  totalDev[i] <- sum(table(devPerMonth[[i]]))
  indexAux <- !(devPerMonth[[i-1]] %in% devPerMonth[[i]])
  index[1:length(indexAux)] <- indexAux
  attritionDev[i] <- sum(table(startEndDev$Name[index]))
  index[1:length(indexAux)] <- FALSE
  
  if(i >= 7){
    indexAux <- (sixthMonth %in% devPerMonth[[i]])
    index[1:length(indexAux)] <- indexAux
    sixMonthDev[i] <- sum(table(startEndDev$Name[index]))
    sixthMonth <- devPerMonth[[i-5]]
    index[1:length(indexAux)] <- FALSE
  }

  if(i >= 13){
    indexAux <- (twelvethMonth %in% devPerMonth[[i]])
    index[1:length(indexAux)] <- indexAux
    twelveMonthDev[i] <- sum(table(startEndDev$Name[index]))
    twelvethMonth <- devPerMonth[[i-11]]
    index[1:length(indexAux)] <- FALSE
  }

  indexAux <- (devPerMonth[[i-1]] %in% devPerMonth[[i]])
  index[1:length(indexAux)] <- indexAux
  newDev[i] <- totalDev[i] - sum(table(startEndDev$Name[index]))
  index[1:length(indexAux)] <- FALSE
}

trackDev <- data.frame(yearMonth = format(as.Date(months), "%Y-%m"), totalDevelopers = totalDev, newDevelopers = newDev,
                       attrition = attritionDev, sixMonthDevelopers = sixMonthDev, twelveMonthDevelopers = twelveMonthDev,
                       percentageSixMonthDevelopers = 100 * sixMonthDev/totalDev, percentageTwelveMonthDevelopers = 100 * twelveMonthDev/totalDev)


write.xlsx(startEndDev, file="./Tracking of developers.xlsx", sheetName="Developers")
write.xlsx(trackDev, file="./Tracking of developers.xlsx", sheetName="Measures", append=TRUE)

