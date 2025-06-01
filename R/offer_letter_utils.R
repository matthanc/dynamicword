# Utility functions for offer letter generation

# It's good practice to note required packages, though they are loaded in main scripts.
# Ensure dplyr, readxl, lubridate, officer are available in the environment when these are called.

offerletter_inputs <- function(your_name, your_email, candidate_first_name, candidate_last_name, department_id, job_code, salary_step, appointment_type, supervisor, work_address) {

  # Parameter assignments (somewhat redundant but matches original script style)
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

  # Create a single-row tibble with input information
  # Assumes 'tibble' function is available (from dplyr or tibble package)
  df <- tibble::tibble(
    your_name = your_name,
    your_email = your_email,
    candidate_first_name = candidate_first_name,
    candidate_last_name = candidate_last_name,
    department_id = department_id,
    job_code = job_code,
    salary_step = salary_step,
    appointment_type = appointment_type,
    supervisor = supervisor,
    work_address = work_address
  )

  # Join with external data tables. These expect specific global variables to be loaded:
  # positions_data, union_data, appointment_data, department_data.

  # Join position information
  # Assumes 'dplyr::left_join' is available
  df <- dplyr::left_join(df, positions_data, by = c("job_code" = "Job Code"))

  # Determine hourly rate using the logic from the shiny application
  # df[[1, N+df$salary_step]] where N is the column index before salary steps start.
  # The shiny app used 12. This relies on 'positions_data' having a specific structure.
  df <- dplyr::mutate(df, hourly_salary = df[[1, 12 + df$salary_step]])

  # Join union information
  df <- dplyr::left_join(df, union_data, by = "Union Code")

  # Join appointment information
  df <- dplyr::left_join(df, appointment_data, by = c("appointment_type" = "abbreviation"))

  # Join department information
  df <- dplyr::left_join(df, department_data, by = c("department_id" = "Department ID"))

  return(df)
}

generate_offer_document <- function(inputs, template_path = "offerletter.docx") {
  # 'inputs' is expected to be the data frame returned by offerletter_inputs()
  # 'template_path' is the path to the .docx template file

  # Assumes 'officer::read_docx' and 'officer::body_replace_all_text' are available.
  # Assumes 'lubridate::today', 'lubridate::month', 'lubridate::day', 'lubridate::year' are available.

  if (!file.exists(template_path)) {
    stop(paste("Template file not found:", template_path))
  }

  doc <- officer::read_docx(template_path)

  # Date formatting for letterdate and signdate
  current_date_formatted <- paste0(
    lubridate::month(lubridate::today(), label = TRUE, abbr = FALSE), " ",
    lubridate::day(lubridate::today()), ", ",
    lubridate::year(lubridate::today())
  )
  sign_date_formatted <- paste0(
    lubridate::month(lubridate::today() + 5, label = TRUE, abbr = FALSE), " ",
    lubridate::day(lubridate::today() + 5), ", ",
    lubridate::year(lubridate::today() + 5)
  )

  # Replacement pairs. Using a list for clarity, then iterating.
  replacements <- list(
    "letterdate"       = current_date_formatted,
    "firstname"        = inputs$candidate_first_name,
    "lastname"         = inputs$candidate_last_name,
    "departmentname"   = inputs$`Department Name`, # Uses backticks for space
    "appointment_abbr" = inputs$appointment_type,
    "appointment_type" = inputs$appointment, # Full appointment name
    "jobcode"          = as.character(inputs$job_code),
    "jobtitle"         = inputs$`Job Title`, # Uses backticks for space
    "unionname"        = inputs$Union,
    "stepnumber"       = as.character(inputs$salary_step),
    "hourlysalary"     = as.character(inputs$hourly_salary),
    "supervisorname"   = inputs$supervisor,
    "worklocation"     = as.character(inputs$work_address),
    "signdate"         = sign_date_formatted,
    "yourname"         = inputs$your_name,
    "youremail"        = inputs$your_email
  )

  for (old_val in names(replacements)) {
    doc <- officer::body_replace_all_text(doc, old_value = old_val, new_value = replacements[[old_val]], warn = FALSE)
  }

  return(doc)
}
