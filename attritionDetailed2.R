library(openxlsx)
library(tidyverse)
library(lubridate)
library(stringr)
#library(reshape2)

#reading file
egiPersonelRaw <- as_tibble(read.xlsx("./Individual SDP Time reports.xlsx", sheet=2, cols = c(14, 19, 25, 46:58), detectDates = TRUE))
egiConsultants <- as_tibble(read.xlsx("./Consultants.xlsx", detectDates = TRUE))
#sdpNetworks <- as_tibble(read.xlsx("./Kerstin_2014-2016_all_extractwithSDPNetworks.xlsx", sheet=1))

#Tidying and removing unrelated data
egiPersonelRaw[is.na(egiPersonelRaw)] <- 0 #replace NAs with zeros
egiPersonelRaw <- egiPersonelRaw[egiPersonelRaw$Unit == "EGI",] #selecting just desired data (India)
egiPersonelRaw <- egiPersonelRaw %>% gather(`Jan.Actual.Hours`:`Dec.Actuals.Hours`, key = "month", value = "hours") #reshaping the data to have months as one variable
egiPersonelRaw <- egiPersonelRaw[egiPersonelRaw$hours != 0, ] #Remove months with zero hours

#egiPersonelRaw$Network <- str_to_lower(egiPersonelRaw$Network) #lower case for the netwrok type
#egiPersonelRaw$Network <- str_replace_all(egiPersonelRaw$Network, "\\s", "") #remove white spaces

#sdpNetworks <- sdpNetworks[!is.na(sdpNetworks$node),] #select the networks of interest
#sdpNetworks$description <- str_to_lower(sdpNetworks$description) #lower case for the netwrok type
#sdpNetworks$description <- str_replace_all(sdpNetworks$description, "\\s", "") #remove white spaces

#egiPersonelRaw <- egiPersonelRaw[(egiPersonelRaw$Network %in% sdpNetworks$description), ] #Select data related to the networks of interest
egiPersonelRaw <- egiPersonelRaw[, c(2, 1, 4, 5)] #Select relevant columns
names(egiPersonelRaw) <- c("name", "netwrok", "year", "month")
names(egiConsultants) <- c("name", "startDate", "endDate", "type")

#Converting months in date format
rawMonth <- unique(egiPersonelRaw$month)
for(i in 1:length(rawMonth)){
  if(i < 10){
    egiPersonelRaw$month[egiPersonelRaw$month == rawMonth[i]] <- paste("0", i, sep = "")
  }
  else{
    egiPersonelRaw$month[egiPersonelRaw$month == rawMonth[i]] <- as.character(i)
  }
}

egiPersonelRaw <- mutate(egiPersonelRaw, date = ymd(paste(year, month, "01", sep = "-")))

#Getting the start and dates for each developer and ordering by start date, end date and name
egiPersonel <- egiPersonelRaw %>% 
  group_by(name) %>%
  summarise(startDate = min(date), endDate = max(date) %m+% months(1) - 1) 

egiPersonel <- egiPersonel[order(egiPersonel$startDate, egiPersonel$endDate, egiPersonel$name), ]


#Adding the type of employee (Ericsson) and consultants
egiPersonel <- mutate(egiPersonel, type = "Ericsson")
egiPersonelFull <- rbind(egiPersonel, egiConsultants)

#Getting months in the period and setting the date for the last day of each month
months <- seq(min(egiPersonelFull$startDate), max(egiPersonelFull$endDate), by="month")  %m+% months(1) - 1

#Calculating attrition, new developers, 6-month developers, 12 month developers, total
totalDev <- numeric(length = length(months))
attritionDev <- numeric(length = length(months))
sixMonthDev <- numeric(length = length(months))
twelveMonthDev <- numeric(length = length(months))
newDev <- numeric(length = length(months))
devPerMonth <- list()
devPerMonth[[1]] <- egiPersonelFull$name[(egiPersonelFull$startDate <= months[1]) & (egiPersonelFull$endDate >= months[1])]
totalDev[1] <- sum(table(devPerMonth[[1]]))
newDev[1] <- totalDev[1] - sixMonthDev[1] - twelveMonthDev[1]
sixthMonth <- devPerMonth[[1]]
twelvethMonth <- devPerMonth[[1]]
index <- vector(length = length(egiPersonelFull$name))

for(i in 2:length(months)){
  devPerMonth[[i]] <- egiPersonelFull$name[(egiPersonelFull$startDate <= months[i]) & (egiPersonelFull$endDate >= months[i])]
  totalDev[i] <- sum(table(devPerMonth[[i]]))
  indexAux <- !(devPerMonth[[i-1]] %in% devPerMonth[[i]])
  index[1:length(indexAux)] <- indexAux
  attritionDev[i] <- sum(table(egiPersonelFull$name[index]))
  index[1:length(indexAux)] <- FALSE
  
  if(i >= 7){ #calculate developers with at least 6-month experience
    indexAux <- (sixthMonth %in% devPerMonth[[i]])
    index[1:length(indexAux)] <- indexAux
    sixMonthDev[i] <- sum(table(egiPersonelFull$name[index]))
    sixthMonth <- devPerMonth[[i-5]]
    index[1:length(indexAux)] <- FALSE
  }
  
  if(i >= 13){ #calculate developers with at least 12-month experience
    
    indexAux <- (twelvethMonth %in% devPerMonth[[i]])
    index[1:length(indexAux)] <- indexAux
    twelveMonthDev[i] <- sum(table(egiPersonelFull$name[index]))
    twelvethMonth <- devPerMonth[[i-11]]
    index[1:length(indexAux)] <- FALSE
  }
  
  indexAux <- (devPerMonth[[i-1]] %in% devPerMonth[[i]])
  index[1:length(indexAux)] <- indexAux
  newDev[i] <- totalDev[i] - sum(table(egiPersonelFull$name[index]))
  index[1:length(indexAux)] <- FALSE
}

measures <- data.frame(yearMonth = format(as.Date(months), "%Y-%m"), totalDevelopers = totalDev, newDevelopers = newDev,
                       attrition = attritionDev, sixMonthDevelopers = sixMonthDev, twelveMonthDevelopers = twelveMonthDev,
                       percentageSixMonthDevelopers = 100 * sixMonthDev/totalDev, percentageTwelveMonthDevelopers = 100 * twelveMonthDev/totalDev)

#Creating workbook
wb <- createWorkbook("Tracking of developers.xlsx")
addWorksheet(wb, "EGI Ericsson Raw")
addWorksheet(wb, "EGI Ericsson")
addWorksheet(wb, "EGI Consultants")
addWorksheet(wb,  "EGI Full")
addWorksheet(wb, "Measures")

#Writing and save data
writeData(wb, "EGI Ericsson Raw", egiPersonelRaw, withFilter = TRUE)
writeData(wb, "EGI Ericsson", egiPersonel, withFilter = TRUE)
writeData(wb, "EGI Consultants", egiConsultants, withFilter = TRUE)
writeData(wb, "EGI Full", egiPersonelFull, withFilter = TRUE)
writeData(wb, "Measures", measures, withFilter = TRUE)
saveWorkbook(wb, file = "./Tracking of developers.xlsx", overwrite = TRUE)
