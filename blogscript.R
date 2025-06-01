# step 1 - load packages and data

library(dplyr)
library(readxl)
library(lubridate)
library(officer)
library(shiny) # Retained as it was in original
source("R/offer_letter_utils.R") # Load centralized functions

# Check for positiondata.xlsx before attempting to load it
if (!file.exists("positiondata.xlsx")) {
  stop("Error: positiondata.xlsx not found. Please ensure it is in the working directory.")
}
positions_data <- read_excel("positiondata.xlsx", sheet = "positions")
union_data <- read_excel("positiondata.xlsx", sheet = "unions")
appointment_data <- read_excel("positiondata.xlsx", sheet = "appointment_type")
department_data <- read_excel("positiondata.xlsx", sheet = "departments")

# step 2 - create a function for inputs
# (The offerletter_inputs function is now sourced from R/offer_letter_utils.R)

# step 3 - create function for entering information into Word
# (The document generation logic is now in generate_offer_document from R/offer_letter_utils.R)

# step 4 - provide inputs
inputs <- offerletter_inputs(
  your_name = "Matt Churchill",
  your_email = "matthanc@gmail.com",
  candidate_first_name = "John",
  candidate_last_name = "Doe",
  department_id = 456,
  job_code = 300,
  salary_step = 1,
  appointment_type = "PERM",
  supervisor = "Jane Smith",
  work_address = "1234 Mission Street"
)

# step 5 - generate document by running offer letter function
offerletter <- generate_offer_document(inputs) # Use the centralized function

# step 6 - save file with desired naming convention
print(offerletter, target = glue::glue("Offer Letter - {inputs$candidate_first_name} {inputs$candidate_last_name} ({inputs$appointment_type} {inputs$job_code}).docx"))
