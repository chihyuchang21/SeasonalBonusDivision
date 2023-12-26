# Automated Quarterly Bonus File Splitting Program

## Structure & Function
This program, according to practical needs, is split into five files based on functionality:
- **SB01_Directory**
  - Main Function: Creating Department Folders on Desktop  
  - User Input the Bonus Year & Season (e.g., 2020Q4)
  - Opening Workbook and Looping through Data
  - Creating Folders
    - Logic: 
      1. Depending on the values in column A and B, it creates specific folder structures on the Desktop.
      2. func2_folder is created based on the values in column A.
      3. func1_folder is created if values in column B are different from those in column A.
      4. plant_folder is created within func1_folder based on additional conditions.
      5. Handles cases where func1 is equal to func2 and a plant_folder exists.  

- **SB02_Dept**  
  - Main Function: Splitting Files on Desktop
  - User Input the Bonus Year & Season (e.g., 2020Q4)
  - Opening Workbook and Extracting File Name
  - Creating Files
    - Logic:
      1. Creates specific folder structures and performs data filtering based on departmental criteria.
      2. Utilizes a nested loop to iterate through sheets (b) and filter data accordingly.
      3. Adjusts row height, column width, and zoom for better visualization.  
  - Saving Files  
    - Saves the files in the appropriate folders based on the conditions:
      1. If foldernameFunc2 is not equal to foldernameFunc1 and foldernamePlant is 0, saves in Func1 folder.
      2. If foldernameFunc2 is not equal to foldernameFunc1 and foldernamePlant is not 0, saves in Func1 subfolder.
      3. If foldernameFunc2 is equal to foldernameFunc1 and foldernamePlant is not 0, saves in Func2 subfolder.
      4. If none of the above conditions apply, saves in Func2 folder.
  - Closing Workbooks

- **SB03_SecSheets**  
  - Main Function: Checking if the workbook is open / Adjusting the row height, column width, and zoom level  

- **SB04_Formula&Appearance**  
  - Main Function: Reconstructing Formulas which were previously formulated for operational needs.
  
- **SB05_PrintingSetting**  
  - Main Function: Setting up the printing layout.


