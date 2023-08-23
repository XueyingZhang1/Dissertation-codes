#load packages
library(readxl)
library(Benchmarking)
library(dplyr)
library(openxlsx)
library(deaR)
library(writexl)
#DEA method
# Define the function to run the dea result according to different sheets.
# Each sheet represents one year, from 2008 to 2014.
run_dea <- function(file_name, sheet_num) {
   # Load the data
   data <- read_xlsx(file_name, sheet = sheet_num)
   # Normalize the data
   log_transform <- function(x) {
      if (is.numeric(x)) {
         return(log(x))
      } else {
         return(x)
      }
   }
   
   # Apply the function to each numeric column
   transformed_data <- data %>%
      mutate(across(where(is.numeric), log_transform))
   # Define input and output data
   input_data <- data[,c("land_area", "population", "energy_use","renewable")]
   output_data <- data[,c("GDP","CO2_emission")]
   
   # Ensure data frames
   input_data <- as.data.frame(input_data)
   output_data <- as.data.frame(output_data)
   
   # Run the DEA analysis
   dea_results <- dea(input_data,output_data,ORIENTATION = "in")
   
   # Return the results
   return(dea_results)
}

all_dea_results <- list()

# Run all years to get dea results.
# Iterate over the sheet numbers
for (sheet_number in 2:8) {
   # Get the result for the current sheet
   dea_result <- run_dea("energy eff.xlsx", sheet_number)
   
   # Store the current result in the list
   all_dea_results[[sheet_number]] <- dea_result$eff
   
}

# Save results to excel file
# Iterate over the sheet numbers
for (sheet_number in 2:8) {
   # Get the result for the current sheet
   dea_result_print <- data.frame(data[1],all_dea_results[[sheet_number]])
   
   # Generate a unique file name for the current sheet
   file_name <- paste0("dea_results_sheet", sheet_number, ".xlsx")
   
   # Export the result to a excel file
   write.xlsx(dea_result_print, file = file_name)
}





# Malmquist index method
# Read the data
data <- read_xlsx("energy eff.xlsx", sheet=1)

# Get numeric columns only
numeric_cols <- sapply(data, is.numeric)

# Apply a log transformation only to numeric columns of the dataset
data_log_transformed <- data
data_log_transformed[, numeric_cols] <- log(data[, numeric_cols])

# Using the make_malmquist function to perform a Malmquist index calculation 
# on a dataset that has undergone logarithmic transformation. 
# The function considers 7 periods of analysis and takes into account 4 inputs 
# and 2 outputs, with an additional parameter related to undesirable output.
data_example <- make_malmquist(data_log_transformed, 
                               nper=7,
                               arrangement="horizontal", 
                               ni=4,
                               no=2,
                               ud_outputs=2)
# Calculate Malmquist Index
result <- malmquist_index(data_example, orientation = "io",rts = "vrs")
mi <- as.data.frame(mi)
efficiency_change <- as.data.frame(efficiency_change)
Pure_Technical_Efficiency_Change <- as.data.frame(Pure_Technical_Efficiency_Change)
Scale_Efficiency_Change <- as.data.frame(Scale_Efficiency_Change)
Technological_Change <- as.data.frame(Technological_Change)


# Put the dataframes in a named list
my_data <- list("Malmquist Index" = mi, 
                "efficiency_change" = efficiency_change, 
                "Technological_Change" = Technological_Change,
                "Pure_Technical_Efficiency_Change" = Pure_Technical_Efficiency_Change,
                "Scale_Efficiency_Change" = Scale_Efficiency_Change)

# Write the data to Excel
write_xlsx(my_data, "my_results_latest.xlsx")


