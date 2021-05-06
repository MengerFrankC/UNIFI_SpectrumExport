#Script to transform (MSMS) data exported from UNIFI into a data structure that can be used to run e.g. MetFrag
#Frank Menger, 2021-02-10

#The general idea is to export MSMS information via a report, and link back this information to the corresponding feature in the component list (which is also exported)
#This script can deal with empty MSMS data files (when no MSMS information is found for a feature, UNIFI simply skips this compound, and does not introduce an empty sheet)
#Note: This script cannot handle errors caused when MSMS data of components was exported that were not included in the exported component list
##### Setting up work environment #####

# Load packages 
library(readxl)
library(xlsx)
library(dplyr)

# Define functions 
# Reads all sheets of an Excel workbook into one list and uses the sheet names ad list object names.
read_excel_allsheets <- function(filename, tibble = FALSE) {
  sheets <- readxl::excel_sheets(filename)
  x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X))
  if(!tibble) x <- lapply(x, as.data.frame)
  names(x) <- sheets
  x
}

# File & directory paths
wd <- getwd()  #note: this assumes you have navigated to the folder containing the output files from UNIFI
wd.MSMS <- paste0(wd, "MSMS\\")
setwd(wd)

#Create directory for separate MSMS files (if wanted/needed)
ifelse (dir.exists(wd.MSMS), yes = "Directory already exists!" , no = dir.create(wd.MSMS)) 

#####Read high CE information --------------------------------------------------------------------------------------------------------------
#The exported report is read using the previously defined function. 
#Each excel sheet becomes one object in the list 'MSMS.list'
MSMS.list <- read_excel_allsheets("Example_data_highCE.xls")

#The sheet names of the Excel document become the names of the list objects.
#These names are created automatically by UNIFI and consist of 2 parts
# 1. Channel information 
# 2. Time of the scan, i.e. time +/- time window in minutes (each with 4 decimals) 
# Example: The name "HighenergyTime4.71080.0387minut" reads as follows: 
#          1. The exported scan was a high energy scan
#          2. The scan was acquired at 4.7108 +/- 0.0387 min


#####Extract scan times from sheet names (without window) -----------------------------------------------------------------------------------
#Create vector of list-object names and remove all text before the first number (removes "HighenergyTime")
times_MSMS <- gsub("^[^0-9]+","", names(MSMS.list))
#Extract the times of the scans. 
#As the format is slighly different for times of 10 minutes and more (because of the additional digit), a loop is used.
for (i in 1:length(times_MSMS)) {
  temp <- times_MSMS[i]   #Temporarily store the name
  if(substr(temp, nchar(temp), nchar(temp)) == "t"){  #If the last letter of the name is "t" (and therefore the time <10 min)
    times_MSMS[i] <- substr(temp, 1, 6)               #Extract the first 6 characters of the string (i.e. the one digit minutes + the dot + the 4 decimals)
  }else{
    times_MSMS[i] <- substr(temp, 1, 7)               #Else, extract the first 7 characters (i.e. the two digit minutes + the dot + the 4 decimals)
  }
}
#Transform the times to numeric and round to 3 digits (necessary because in the component list the times have more decimals)
times_MSMS <- round(as.numeric(times_MSMS), digits = 3)

#####Extract m/z and intensity [%] columns of MSMS data -----------------------------------------------------------------------------------
MSMS.list.clean <- MSMS.list  #Create copy that will be 'cleaned', while keeping the original data
for (i in 1:length(MSMS.list.clean)) {
  #Extract columns 3 and 5 (m/z and intensity [%], respectively).
  MSMS.list.clean[i] <- lapply (MSMS.list.clean[i], function(x) x[c(3,5)])    #Note: Unfortunately, the column names given by UNIFI make actual numbers here necessary. 
}

#Read component list from UNIFI
cmpd.info <- read_excel_allsheets("Example_ComponentList_from_UNIFI.xls")

#Insert identifiers consisting of sample name + count in a $data_file column (which are used to link with separately saved MSMS files)
for (i in 1:length(cmpd.info)) {
  for (j in 1:length(cmpd.info[[i]][,1])) {
    temp <- paste(names(cmpd.info[i]), j, sep = "_") #create temporary (temp) vector of the sample name + number (counting the length of the component list of a sample i)
    cmpd.info[[i]]$data_file[j] <- temp              #create column $data_file from vector temp
  }
}

#Paste MSMS info into a column $MSMS (after a check that MSMS information is  available based on the time)  
for (i in 1:length(cmpd.info)) {       #For loop cycling through the different samples
  if(i == 1) {
    temp_times_MSMS <- times_MSMS
    error_count <- 0
    MSMS_count <- 1
  }
  cmpd.info[[i]]$MSMScheck <- "."      #Column $MSMScheck is introduced, which is below used to track errors
  for (j in 1:length(cmpd.info[[i]][,1])) {  #For loop cycling through the  components of a sample
    #Check that the RT of the component matches the time of the scan
    if  (between(round(cmpd.info[[i]]$`Observed RT (min)`[j], digits = 3), temp_times_MSMS[MSMS_count]-0.02, temp_times_MSMS[MSMS_count]+0.02)){
      x <- MSMS_count-error_count
      cmpd.info[[i]]$MSMS[j] <- paste(MSMS.list.clean[[x]][,1],MSMS.list.clean[[x]][,2], sep = "_", collapse = ";" )
    }else{
      cmpd.info[[i]]$MSMScheck[j] <- "error"   #In case times do not match an error is marked, and
      error_count <- error_count+1             #the error count is increased by 1, and
                                               #data_file names get moved down by one row to keep the corresponding names correct, and
      cmpd.info[[i]]$data_file[(j+1):length(cmpd.info[[i]][,1])] <- cmpd.info[[i]]$data_file[j:(length(cmpd.info[[i]][,1])-1)]
                                               #a placeholder "0" is inserted into temp_times_MSMS (effectively prolonging the vector by one, and continuing to use the current MSMS data until the fitting component is found)
      temp_times_MSMS <- c(temp_times_MSMS[1:(MSMS_count-1)], 0, temp_times_MSMS[MSMS_count:length(temp_times_MSMS)])
    }
    MSMS_count <- MSMS_count+1
  }
}
setwd(wd)

#Combine the list objects of cmpd.info to one dataframe df.cmpd.info
for (i in 1:length(cmpd.info)) {
  if(i == 1){
    df.cmpd.info <- cmpd.info[[i]]
  }else{
    df.cmpd.info <- rbind(df.cmpd.info, cmpd.info[[i]])
  }
}

#Save as a .csv file, which can be used to run e.g. MetFrag.
write.csv(df.cmpd.info, paste0("", ".csv"), row.names = F)

#Save MSMS in separate files (for better oversight / control; can also be used to run MetFrag)
setwd(wd.MSMS)
for (i in 1:length(MSMS.list.clean)) {                           #For loop cycling through the list MSMS.list.clean
  if (i == 1) number <- 0                                        #Set number to 0 at the start of the loop
  item.name <- gsub(" ","", MSMS.list[[i]][1, c("Item name")])   #extract the item.name (i.e. the sample name) from the i'th list object
  number <- number + 1                                           
  temp <- paste(item.name, number, sep = "_")                    #Build the file name from the sample name + a continuous count for that sample, starting from 1
  #Create file with the MSMS information using temp as the file name.
  write.table(MSMS.list.clean[[i]], file = paste(temp, ".txt", sep = ""), quote = FALSE, row.names = FALSE, col.names = FALSE)
  #Re-start the counter 'number' when a new sample starts
  if (MSMS.list[[i]][1, c("Item name")] != MSMS.list[[i+1]][1, c("Item name")]){  #Note: there will be an error "Error in MSMS.list[[i + 1]] : subscript out of bounds", as this condition cannot be tested for the last listobject
    number <- 0
  }
}
setwd(wd)


