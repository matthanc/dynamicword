This repository contains scripts and files used in this tutorial on how to generate form letters with R and the officer package: https://matthanc.ghost.io/how-to-generate-dynamic-word-documents-with-r/

A working version of the shiny application can be found here: https://matthanc.shinyapps.io/dynamicworddemo/
## Prerequisites

*   **R**: The R environment must be installed. You can download it from [CRAN (The Comprehensive R Archive Network)](https://cran.r-project.org/).
*   **R Packages**: The following R packages are required:
    *   `dplyr`
    *   `readxl`
    *   `lubridate`
    *   `officer`
    *   `shiny`
    *   `glue`

## Installation

1.  **Install R**: If you haven't already, download and install R from CRAN.
2.  **Install R Packages**: Open an R console and run the following command to install the necessary packages:
    ```R
    install.packages(c("dplyr", "readxl", "lubridate", "officer", "shiny", "glue"))
    ```

## Usage

This repository contains two main R scripts:

1.  **`blogscript.R`**:
    *   This script demonstrates the core logic of generating an offer letter using predefined example data.
    *   To run it:
        1.  Ensure all prerequisite packages are installed.
        2.  Open R or an R-compatible IDE (like RStudio).
        3.  Set your working directory to the root of this repository.
        4.  Source the script: `source("blogscript.R")`
        5.  This will generate an offer letter named `Offer Letter - John Doe (PERM 300).docx` in the root directory, using the example data hardcoded in the script.

2.  **`shinyapplicaiton.R`**:
    *   This script launches a Shiny web application that provides a user interface for inputting offer letter details and downloading the generated document.
    *   To run it:
        1.  Ensure all prerequisite packages are installed.
        2.  Open R or an R-compatible IDE (like RStudio).
        3.  Set your working directory to the root of this repository.
        4.  Run the Shiny application: `shiny::runApp("shinyapplicaiton.R")`
        5.  This will launch the application in your default web browser. You can then fill in the form and click "Download Offer Letter".

**Note**: Both scripts rely on `positiondata.xlsx` for position, union, appointment, and department details, and `offerletter.docx` as the template for the offer letter. These files must be present in the root directory of the project. The utility functions for generating the letter are located in `R/offer_letter_utils.R`.