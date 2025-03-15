# Forms Automation

This project automates the process of validating and submitting data from an Excel file to a Google Form. It consists of several Python scripts that work together to create a worktray, validate the data and submit it to a Google Form. The project is designed to be robust, with detailed logging for easy debugging.

---

## Table of Contents

1. [Project Overview](#project-overview)
2. [Features](#features)
3. [How It Works](#how-it-works)
4. [Installation](#installation)
5. [Usage](#usage)
6. [Logs](#logs)
7. [File Structure](#file-structure)
8. [Dependencies](#dependencies)
9. [Contributing](#contributing)
10. [License](#license)

---

## Project Overview

The **Forms Automation** project is designed to automate the following tasks:
1. **Worktray creation**: Generate an Excel file (`worktray.xlsx`) from an input file (`input_file.xlsx`) using a template (`worktray_template.xlsx`).
2. **Validate Data**: Validate the data in the worktray, ensuring that all required fields are filled, dates are in the correct format, and text/number fields are valid.
3. **Submit to Google Forms**: Submit the validated data to a Google Form, updating the worktray with the submission status.

The project is modular, with each task handled by a separate Python script. Detailed logs are generated to facilitate debugging and monitoring.

---

## Features

- **Worktray Creation**: Automatically generates a worktray from an input file using a predefined template.
- **Data Validation**:
  - Checks for empty fields.
  - Validates date formats (both Excel numeric dates and text dates).
  - Ensures text fields contain only allowed characters.
  - Validates numeric fields.
- **Google Forms Submission**: Submits validated data to a Google Form and updates the worktray with submission status.
- **Detailed Logging**: Logs every step of the process for easy debugging and monitoring.
- **Error Handling**: Robust error handling to ensure the process stops gracefully if an error occurs.

---

## How It Works

1. **Worktray Creation**:
   - The `worktray_creation.py` script reads the input file (`input_file.xlsx`) and creates the worktray (`worktray.xlsx`) using a template (`worktray_template.xlsx`).
   - It ensures that all required columns are present and initializes the worktray with default values.

2. **Data Validation**:
   - The `input_file_processment.py` script validates the data in the worktray.
   - It checks for empty fields, validates date formats, ensures text fields contain only allowed characters, and validates numeric fields.
   - Invalid rows are flagged, and observations are added to the worktray.

3. **Google Forms Submission**:
   - The `google_forms_submission.py` script submits the validated data to a Google Form.
   - It updates the worktray with the submission status and logs any errors that occur during submission.

4. **Main Script**:
   - The `main.py` script orchestrates the entire process, running the above scripts in sequence and logging the results.

---

## Installation

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/your-repo/forms-automation.git
   cd forms-automation

2. **Install Dependencies**:
   Ensure you have Python 3.x installed. Then, install the required packages.
   ```bash
   pip install -r requirements.txt

3. **Set Up Input Files**:
   - Place your input file (input_file.xlsx) in the input directory.
   - Ensure the worktray_template.xlsx file is also in the input directory.

4. **Configure Google Forms URL**:
    - Update the GOOGLE_FORM_URL variable in google_forms_submission.py with your Google Form's submission URL.

---

## Usage

1. **Run the Main Script**:
python main.py

2. **Check the Logs**:
    - Logs are saved in the _logs directory
    - The main log file is process_logs.log

3. **Review the Worktray**:
    - The updated worktray (worktray.xlsx) is saved in the process_data directory.
    - It includes validation results and submission status.

## Logs

The project generates detailes logs for each step of the process. Logs are saved in the _logs directory with the following structure:
   - process_logs.log: Main log file for the entire process.
   - worktray_creation.log: Logs for the worktray creation step.
   - input_file_processment.log: Logs for the data validation step.
   - google_forms_submission.log: Logs for the Google Forms submission step.

---   

## File Structure

1. **forms-automation/**
   - _logs/                       # Log files
   - input/
      - input_file.xlsx           # Input data file
      - worktray_template.xlsx    # Worktray template
   - process_data/
      - worktray.xlsx             # Generated worktray
   - main.py                      # Main script to run the entire process
   - worktray_creation.py         # Script to create the worktray
   - input_file_processment.py    # Script to validate the worktray
   - google_forms_submission.py   # Script to submit data to Google Forms
   - README.md                    # Project documentation
   - requirements.txt             # List of dependencies

---

## Dependencies

- Python 3.x
- openpyxl (for Excel file handling)
- pandas (for data manipulation)
- requests (for Google Forms submission)
- shutil (for High-level file operations)
- logging (For the process logs)
- re (Regular expressions for special character validation)