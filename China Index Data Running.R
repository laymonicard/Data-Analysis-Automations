###########################################################
#CHINA INDEX DATA AGGREGATION AND CALCULATION
#PREPARED FOR ONE OF CELIOS' RESEARCH PROJECTS
#COMBINED, MODIFIED, AND CUSTOMIZED FROM VARIOUS SOURCES
#SPECIAL CREDITS FOR STACK OVERFLOW COMMUNITY AND JONATHAN NG
#REFERENCES ARE PLACED TO THE BOTTOM OF THIS CODE
###########################################################
#LAY MONICA RATNA DEWI
#3 JULY 2024
###########################################################

#Case: This code was written to combine multiple files from multiple sheets and perform calculations
#The files are transformed into data frames before being processed for calculation

# Setting up libraries
library(purrr)
library(readxl)
library(openxlsx)
library(tidyverse)
library(dplyr)
library(rio)
library(readr)

# Setting up directory
setwd("D:/celios/chinaindex")
datasets <- list.files("D:/celios/chinaindex",pattern = "*.xlsx",full.names = TRUE)


####################PREPARING DATAFRAME####################
# Review sheet name
excel_sheets(datasets[1])

# Remove unused sheets ("Intro" and "Selection Options") for all files
#Set up a function to remove unused sheets
clean_sheet <- function(datasets){
  #Loading the workbook
  wb <- loadWorkbook(datasets) 
  
  #Create list of unused sheets
  sheet_remove <- c("Intro","Selection Options")
  
  #Remove unused sheets
  for (sheet_name in sheet_remove) {
    if (sheet_name %in% excel_sheets(datasets[1])) {
      removeWorksheet(wb, sheet = sheet_name)
    } else {
      print("No remaining unused sheets")
    }
  }
  
  #Save the workbook
  saveWorkbook(wb,file= datasets,overwrite = TRUE)
  
}
#Apply function to all worksheets
lapply(datasets, clean_sheet)



####################MERGE SHEETS############################
#Automate the below commands
#df_1 <- excel_sheets(datasets[1]) %>% map_df(~read.xlsx(datasets[1],.))
#df_2 <- excel_sheets(datasets[1]) %>% map_df(~read.xlsx(datasets[2],.))
#df_n <- excel_sheets(datasets[n]) %>% map_df(~read.xlsx(datasets[n],.))

#Create function that automates data merging
merge_datasets <- function(datasets) {
  # Create a list to store the results
  dataframes <- list()
  
  # Loop over each dataset
  for (i in seq_along(datasets)) {
    # Get the sheet names for the current dataset
    sheets <- excel_sheets(datasets[[i]])
    
    # Read all sheets for the current datasets and bind them row-wise
    df <- sheets %>% 
      map_df(~read.xlsx(datasets[[i]], sheet = .x)) %>% mutate_all(as.character)
    
    # The result must be stored
    dataframes[[i]] <- df
  }
  
  # Return the result
  return(dataframes)
}

# Generate list of data sets to which the function will be applied to
datasets <- list(".../data/Bali 2_7 done.xlsx", #insert path and file name
                 ".../" #insert other used paths here
                 ".../data/Sumatera Utara 3_7 done.xlsx")
                 
#Apply function to merge the data sets to the worksheets
map_df(datasets,merge_datasets)
processed_data <- merge_datasets(datasets)


# Access processed_data[[1]], processed_data[[2]], etc. to review
# Convert all processed data as data frames
new_df <- lapply(processed_data,as.data.frame)



####################ASSIGNING NAME OF PROVINCES TO EACH DATA FRAME####################
#Create a function to assign province label to each data frame
labelling <- function(labelling_province){

#Generate list of provinces to label data frames  
province_list <- c('Bali',
                   'Bangka Belitung',
                   'Banten',
                   'Bengkulu',
                   'Daerah Istimewa Yogyakarta',
                   'DKI Jakarta',
                   'Gorontalo',
                   'Jambi',
                   'Jawa Barat',
                   'Jawa Tengah',
                   'Jawa Timur',
                   'Kalimantan Barat',
                   'Kalimantan Selatan',
                   'Kalimantan Tengah',
                   'Kalimantan Timur',
                   'Kalimantan Utara',
                   'Kepulauan Riau',
                   'Lampung',
                   'Maluku',
                   'Maluku Utara',
                   'Nanggroe Aceh Darussalam',
                   'Nusa Tenggara Barat',
                   'Nusa Tenggara Timur',
                   'Papua',
                   'Papua Barat Daya',
                   'Papua Pegunungan',
                   'Papua Selatan',
                   'Papua Tengah',
                   'Riau',
                   'Sulawesi Barat',
                   'Sulawesi Selatan',
                   'Sulawesi Tengah',
                   'Sulawesi Tenggara',
                   'Sulawesi Utara',
                   'Sumatera Barat',
                   'Sumatera Selatan',
                   'Sumatera Utara')

#Automate the following code
#province_1 <- paste(province_list[1])
#indv_df_1 <- as.data.frame(cbind(province_1,new_df[[1]]))

#Create an empty list as a placeholder
final_indv_df <- list()

#Write loop function to automate the label assignment
for (i in seq_along(new_df)) {
  #Load each new_df
  new_df_load <- new_df[[i]]
  
  # Add province label as a new column to df
  province_label <- province_list[i]
  df <- cbind(province_label = province_label, new_df_load)
  
  # Store the result in final_indv_df
  final_indv_df[[i]] <- df
}


#Return the result
return(final_indv_df)
}
#Apply the data labelling function to the datasets
new_data <- labelling(datasets[[i]])

#Combine the data frames
combined_df <- bind_rows(new_data)

#Overview the result 
View(combined_df)

#Filter data frames by desired columns
combined_df_filtered <- combined_df %>% 
  select(province_label,Domain,z,Indicator,Response) %>% 
  filter(!is.na(Response))

#Rename wrong column name and combine it with the correct one
combined_df_final <- combined_df_filtered %>% 
  mutate(Category = coalesce(z,Domain)) %>% 
  select(-Domain,-z)

View(combined_df_final)

#Remove duplicated values
unique(combined_df_final$Response)

#Assign score to each response based on the predetermined criteria
final_df <- combined_df_final %>% 
  mutate(Score = case_when(
    Response == "0 - No" ~ 0,
    Response == "1 - Yes (Few, insignificant)" ~ 1,
    Response == "2 - Yes (More than a few, insignificant)" ~ 2,
    Response == "3 - Yes (Few, significant)" ~ 3,
    Response == "4 - Yes" ~ 4,
    Response == "4 - Yes (More than a few, significant)" ~ 4,
    TRUE ~ NA_real_
  ))

View(final_df)
colnames(final_df)

final_df <- distinct(final_df)
View(final_df)

#Create scores summaries based on provinces, domain, or both
score_by_province_and_domain <- (xtabs(Score ~ province_label+Category, data = final_df)) %>% as.data.frame() %>% rename(provinces = province_label)
View(score_by_province_and_domain)

score_by_province <- (xtabs(Score ~ province_label, data = final_df)) %>% as.data.frame()
View(score_by_province)

score_by_domain <- (xtabs(Score ~ Category, data = final_df)) %>% as.data.frame()
View(score_by_domain)

#Store to Excel workbook as backups
score_1 <- write.xlsx(x = score_by_province_and_domain,file = "D:/celios/rawinfluencescore1.xlsx", asTable = TRUE)
score_2 <- write.xlsx(x = score_by_province,file = "D:/celios/rawinfluencescore2.xlsx", asTable = TRUE)
score_3 <- write.xlsx(x = score_by_domain,file = "D:/celios/rawinfluencescore3.xlsx", asTable = TRUE)

#Calculate influence score based on the predetermined mechanism
#Assign score to each response based on the predetermined criteria
score_by_province_and_domain <- score_by_province_and_domain %>% 
  mutate(max_score = case_when(
    Category == "Academia" ~ 4*11,
    Category == "Domestic Politics" ~ 4*10,
    Category == "Economy" ~ 4*6,
    Category == "Foreign Policy" ~ 4*5,
    Category == "Law Enforcement" ~ 4*7,
    Category == "Media" ~ 4*11,
    Category == "Society" ~ 4*9,
    Category == "Technology" ~ 4*7,
    TRUE ~ NA_real_
  )) %>% mutate(influence_score_percent = Freq/max_score*100)
View(score_by_province_and_domain)

final_score <- write.xlsx(x = score_by_province_and_domain,file = "D:/celios/influencescore_all_provinces.xlsx", asTable = TRUE)

#REFERENCES
#https://stackoverflow.com/questions/73901996/how-to-combine-multiple-excel-files-into-one-using-r
#https://stackoverflow.com/questions/60527703/r-loop-to-read-multiple-excel-files-in-list
#https://stackoverflow.com/questions/17018138/using-lapply-to-apply-a-function-over-list-of-data-frames-and-saving-output-to-f
#https://stackoverflow.com/questions/42198715/how-to-remove-the-first-sheets-or-read-the-second-sheets-of-the-excel-files
#https://stackoverflow.com/questions/68187461/cant-combine-x-character-and-x-double
#https://stackoverflow.com/questions/71576456/how-to-add-a-new-column-in-r-dataframe-with-a-constant-number
#https://stackoverflow.com/questions/40550415/how-can-i-merge-multiple-sheets-of-an-excel-workbook-in-r
#https://stackoverflow.com/questions/68860922/r-combine-multiple-excel-sheets-after-applying-row-to-names-function-over-each-s
#https://stackoverflow.com/questions/58199857/create-a-new-column-in-data-frame-by-matching-two-values-between-data-frames
#https://stackoverflow.com/questions/67947832/create-a-new-column-in-r-by-looking-up-another-data-frame
#https://stackoverflow.com/questions/7531868/how-to-rename-a-single-column-in-a-data-frame
#https://stackoverflow.com/questions/68287677/merging-multiple-dataframes-in-r
#https://stackoverflow.com/questions/46305724/merging-of-multiple-excel-files-in-r
#https://stackoverflow.com/questions/12945687/read-all-worksheets-in-an-excel-workbook-into-an-r-list-with-data-frames
#Jonathan Ng. (2018). Combine Multiple Excel Sheets into a Single Table using R Tidyverse. Accessed from https://www.youtube.com/watch?v=3TUyp4ZMu88.
#OpenAI. (2023). ChatGPT (Sep 25 version) [Large language model]. https://chat.openai.com/chat.
                     
