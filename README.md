# OCDC: OpenClinica Data Cleaner
## About
OCDC is a data cleaner for excel spreadsheet files exported from OpenClinica.

It contains 3 major functionalities:
1. Data Cleaning
2. Category and Questionnaire Filtering
3. Diagnosis, DOB and Gender Appending 

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

### 4. Export Data
1. OCDC will now export the cleaned dataset as an excel spreadsheet to the same location the as uncleaned spreadsheet file you first chose with "CLEANED_" appended to the filename


## Downloading diagnoses, DOBs, and genders from Clinical Conductor
1. In Clinical Conductor, go to Patient List.
2. Make sure all participants are displayed by clicking the radio button Maximum to Display 'All'.
3. Go to Display Options and make sure 'Participant Code' is ticked.
4. Hit refresh.
5. Go to EXTENDED FILTER
6. Filter by condition (e.g. Alzheimer's disease) by selecting 'Use a General Data Filter' and your filter, (e.g. AD).
7. From the menu above the patient list, select EDIT EXPORT.
8. Make sure ONLY Participant Code, DOB, and gender are selected. Click save.
9. Of the four export option, hit the rightmost button (says 'General export to an Excel' on hover). Rename the Excel file to something meaningful to you.
