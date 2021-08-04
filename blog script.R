# step 1 - Load packages and data

library(dplyr)
library(readxl)
library(lubridate)
library(officer)
library(shiny)

positions_data <- read_excel("positiondata.xlsx", sheet = "positions")
union_data <- read_excel("positiondata.xlsx", sheet = "unions")
appointment_data <- read_excel("positiondata.xlsx", sheet = "appointment_type")
department_data <- read_excel("positiondata.xlsx", sheet = "departments")

# Step 2 - create a function for inputs

offerletter_inputs <- function(your_name, your_email, candidate_first_name, candidate_last_name, department_id, job_code, salary_step, appointment_type, supervisor, work_address) {
  
  your_name <- your_name
  your_email <- your_email
  candidate_first_name <- candidate_first_name
  candidate_last_name <- candidate_last_name
  department_id <- department_id
  job_code <- job_code
  salary_step <- salary_step
  appointment_type <- appointment_type
  supervisor <- supervisor
  work_address <- work_address
  
  # create single row table with input information
  df <- tibble(your_name = your_name,
               your_email = your_email,
               candidate_first_name = candidate_first_name,
               candidate_last_name = candidate_last_name,
               department_id = department_id,
               job_code = job_code,
               salary_step = salary_step,
               appointment_type = appointment_type,
               supervisor = supervisor,
               work_address = work_address)
  
  
  # join position information
  df <- left_join(df, positions_data, by = c("job_code" = "Job Code"))
  
  # determine hourly rate
  df <- df %>% mutate(hourly_salary = df[[1,14+df$salary_step]])
  
  # join union information
  df <- left_join(df, union_data, by = "Union Code")
  
  #join appointment information
  df <- left_join(df, appointment_data, by = c("appointment_type" = "abbreviation"))
  
  #join department information
  df <- left_join(df, department_data, by = c("department_id" = "Department ID"))
  
  return(df)
}

# Step 3 - Create function for entering information into Microsoft Word

offerletterfunciton <- function(offerinputs) {
  
  # read the offerletter and bring to environment
  df <- read_docx("offerletter.docx")
  
  # funcitnos to find and replace merge text (16 total)
  
  # replace "letterdate" with today's date
  df <- body_replace_all_text(df,
                              old_value = "letterdate",
                              new_value = paste0(month(today(), label = TRUE, abbr = FALSE ), " ", day(today()),",", " ", year(today())))
  
  df <- body_replace_all_text(df,
                              old_value = "firstname",
                              new_value = inputs$candidate_first_name)
  
  df <- body_replace_all_text(df,
                              old_value = "lastname",
                              new_value = inputs$candidate_last_name)
  
  df <- body_replace_all_text(df,
                              old_value = "departmentname",
                              new_value = inputs$`Department Name`)
  
  df <- body_replace_all_text(df,
                              old_value = "appointment_abbr",
                              new_value = inputs$appointment_type)
  
  df <- body_replace_all_text(df,
                              old_value = "appointment_type",
                              new_value = inputs$appointment)
  
  df <- body_replace_all_text(df,
                              old_value = "jobcode",
                              new_value = as.character(inputs$job_code)) #change numeric variables to character
  
  df <- body_replace_all_text(df,
                              old_value = "jobtitle",
                              new_value = inputs$`Job Title`)
  
  df <- body_replace_all_text(df,
                              old_value = "unionname",
                              new_value = inputs$Union)
  
  df <- body_replace_all_text(df,
                              old_value = "stepnumber",
                              new_value = as.character(inputs$salary_step))
  
  df <- body_replace_all_text(df,
                              old_value = "hourlysalary",
                              new_value = as.character(inputs$hourly_salary))
  
  df <- body_replace_all_text(df,
                              old_value = "supervisorname",
                              new_value = inputs$supervisor)
  
  df <- body_replace_all_text(df,
                              old_value = "worklocation",
                              new_value = as.character(inputs$work_address))
  
  df <- body_replace_all_text(df,
                              old_value = "signdate",
                              new_value = paste0(month(today()+5, label = TRUE, abbr = FALSE ), " ", day(today()+5),",", " ", year(today()+5)))
  
  df <- body_replace_all_text(df,
                              old_value = "yourname",
                              new_value = inputs$your_name)
  
  df <- body_replace_all_text(df,
                              old_value = "youremail",
                              new_value = inputs$your_email)
  

  return(df)
  
  }

# Step 4 Provide inputs

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
  work_address = "1234 Market Street"
)


#Step 5 Generate document by running offer letter funciton

offerletter <- offerletterfunciton()

#Step 6 Save file - since we're using R we can specify naming convention

#save conditional offer letter
print(offerletter, target = glue::glue("Offer Letter - {inputs$candidate_first_name} {inputs$candidate_last_name} ({inputs$appointment_type} {inputs$job_code}).docx"))



#shiny server
#your_name, your_email, candidate_first_name, candidate_last_name, department_id, job_code, salary_step, appointment_type, supervisor, work_address

ui <- fluidPage(
  textInput("your_name", "Your Name"),
  textInput("your_email","Your Email"),
  textInput("candidate_first_name","Candidate's First Name"),
  textInput("candidate_last_name","Candidate's Last Name"),
  selectInput("department_id","Department ID", unique(department_data$`Department ID`)),
  selectInput("job_code","Job Code", unique(positions_data$`Job Code`)),
  numericInput("salary_step", "Salary Step", value = 1, min = 1, max = 5),
  selectInput("appointment_type","Appointment Type", unique(appointment_data$abbreviation)),
  textInput("supervisor","Supervisor Name"),
  textInput("work_address","Work Address (Street number and name only)"),
  downloadButton("downloadoffer", "Download Offer Letter")
)

#attempt 2
server <- function(input, output, session) {
  
  output$downloadoffer <- downloadHandler(
    
    
    filename = function() glue::glue("Offer Letter - {input$candidate_first_name} {input$candidate_last_name} ({input$appointment_type} {input$job_code}).docx"),
    
    content = function(file) {
      
      inputs <- offerletter_inputs(your_name = input$your_name,
                         your_email = input$your_email,
                         candidate_first_name = input$candidate_first_name,
                         candidate_last_name = input$candidate_last_name,
                         department_id = as.double(input$department_id),
                         job_code = as.double(input$job_code),
                         salary_step = input$salary_step,
                         appointment_type = input$appointment_type,
                         supervisor = input$supervisor,
                         work_address = input$work_address
                         )
      
      # read the offerletter and bring to environment
      df <- read_docx("offerletter.docx")
      
      # funcitnos to find and replace merge text (15 total)
      
      # replace "letterdate" with today's date
      df <- body_replace_all_text(df,
                                  old_value = "letterdate",
                                  new_value = paste0(month(today(), label = TRUE, abbr = FALSE ), " ", day(today()),",", " ", year(today())))
      
      df <- body_replace_all_text(df,
                                  old_value = "firstname",
                                  new_value = inputs$candidate_first_name)
      
      df <- body_replace_all_text(df,
                                  old_value = "lastname",
                                  new_value = inputs$candidate_last_name)
      
      df <- body_replace_all_text(df,
                                  old_value = "departmentname",
                                  new_value = inputs$`Department Name`)
      
      df <- body_replace_all_text(df,
                                  old_value = "appointment_abbr",
                                  new_value = inputs$appointment_type)
      
      df <- body_replace_all_text(df,
                                  old_value = "appointment_type",
                                  new_value = inputs$appointment)
      
      df <- body_replace_all_text(df,
                                  old_value = "jobcode",
                                  new_value = as.character(inputs$job_code)) #change numeric variables to character
      
      df <- body_replace_all_text(df,
                                  old_value = "jobtitle",
                                  new_value = inputs$`Job Title`)
      
      df <- body_replace_all_text(df,
                                  old_value = "unionname",
                                  new_value = inputs$Union)
      
      df <- body_replace_all_text(df,
                                  old_value = "stepnumber",
                                  new_value = as.character(inputs$salary_step))
      
      df <- body_replace_all_text(df,
                                  old_value = "hourlysalary",
                                  new_value = as.character(inputs$hourly_salary))
      
      df <- body_replace_all_text(df,
                                  old_value = "supervisorname",
                                  new_value = inputs$supervisor)
      
      df <- body_replace_all_text(df,
                                  old_value = "worklocation",
                                  new_value = as.character(inputs$work_address))
      
      df <- body_replace_all_text(df,
                                  old_value = "signdate",
                                  new_value = paste0(month(today()+5, label = TRUE, abbr = FALSE ), " ", day(today()+5),",", " ", year(today()+5)))
      
      df <- body_replace_all_text(df,
                                  old_value = "yourname",
                                  new_value = inputs$your_name)
      
      df <- body_replace_all_text(df,
                                  old_value = "youremail",
                                  new_value = inputs$your_email)
      
      print(df, target = file)
      
    }
  )
}

shinyApp(ui, server)
