library(openxlsx)
library(tidyverse)
library(lubridate)
library(stringr)
#library(reshape2)

#reading file
egiPersonel <- as_tibble(read.xlsx("./Individual SDP Time reports.xlsx", sheet=2, cols = c(14, 19, 22, 46:58), detectDates = TRUE))
sdpNetworks <- as_tibble(read.xlsx("./Kerstin_2014-2016_all_extractwithSDPNetworks.xlsx", sheet=1))

#Tidying and removing unrelated data
egiPersonel[is.na(egiPersonel)] <- 0 #replace NAs with zeros
egiPersonel <- egiPersonel[egiPersonel$Trading.Partner == "EGI",] #selecting just desired data (India)
egiPersonel <- egiPersonel %>% gather(`Jan.Actual.Hours`:`Dec.Actuals.Hours`, key = "month", value = "hours") #reshaping the data to have months as one variable
egiPersonel$Network <- str_to_lower(egiPersonel$Network) #lower case for the netwrok type
egiPersonel$Network <- str_replace_all(egiPersonel$Network, "\\s", "") #remove white spaces
sdpNetworks <- sdpNetworks[!is.na(sdpNetworks$node),] #select the networks of interest
sdpNetworks$description <- str_to_lower(sdpNetworks$description) #lower case for the netwrok type
sdpNetworks$description <- str_replace_all(sdpNetworks$description, "\\s", "") #remove white spaces
egiPersonel <- egiPersonel[(egiPersonel$Network %in% sdpNetworks$description), ] #Select data related to the networks of interest
egiPersonel <- egiPersonel[, c(2, 4, 5)] #Select relevant columns
#Converting months in date format



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

#Cahnge here the persistence part
wb <- createWorkbook("Tracking of developers.xlsx") 
write.xlsx(startEndDev, file="./Tracking of developers.xlsx", sheetName="Developers", overwrite = TRUE)
write.xlsx(trackDev, file="./Tracking of developers.xlsx", sheetName="Measures", append=TRUE)