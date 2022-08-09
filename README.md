# OCDC: OpenClinica Data Cleaner
## About
OCDC is a data cleaner for excel spreadsheet files exported from OpenClinica.

It contains 4 major functionalities:
1. Data Cleaning
2. Category and Questionnaire Filtering
3. Diagnosis, DOB and Gender Appending 
4. Undiagnosed Participant Filtering

## Installation
1. Download or clone OCDC
2. Run 'cleandata.m' (found in datacleaner folder) in MatLab 2018b or above
3. If you would like to append diagnoses to 

## How To use
### 1. Import Data
1. Open OCDC in MatLab 2018b or above
2. Run (green arrow button or F5 (Windows)
3. Select the Excel spreadsheet file you want to clean (must be an unedited excel spreadsheet downloaded from OpenClinica)
* Important note: OpenClinica currently exports corrupted Excel files. To repair a file, first open it in Microsoft Excel and re-save it as an .xlsx file with a .xlsx extension
4. The excel spreadsheet will then be imported (may take some time)

### 2. Optional: Select Specific Questionnaires/Categories
1. OCDC can identify questionnaires and categories within the datast
2. If you would like to only keep specific data then click 'yes' and select the categories you would like to keep
3. The dropdown list supports multiselection of items
4. Press 'done'
5. OCDC will now filter for the categories you chose to keep and remove all others

### 3. Optional: Append Diagnoses, DOBs and Gender
1. OCDC can append Clinical Conductor diagnoses, DOBs and genders to the OpenClinica dataset
2. This requires a participant list to be exported from Clinical Conductor (see instructions below)
3. If you would like to append this information, click 'yes' and select the Clinical Conductor patient list
4. Enter the name of the diagnosis for that list of participants
5. Press 'OK'

### 4. Optional: Remove Undiagnosed Participants
1. Click 'yes' if you would like to remove any participants in the dataset who did not have a diagnosis appended to them

### 5. Export Data
1. OCDC will now export the cleaned dataset as an excel spreadsheet to the same location the as uncleaned spreadsheet file you first chose with "CLEANED_" appended to the beginning of the filename.
* If running on Mac, each Study Event will be saved as a separate Excel file with the Study Event appended to the end of the filename.
