# Automated Quarterly Bonus File Splitting Program

## Structure & Function
This program, according to practical needs, is split into five files based on functionality:
- **SB01_Directory**  
  - Creating Folder on Desktop  
  - User Input the Bonus Year & Season (e.g., 2020Q4)
  - Opening Workbook and Looping through Data
  - Creating Folders  
    - Logic: 
        - Depending on the values in column A and B, it creates specific folder structures on the Desktop.
        - func2_folder is created based on the values in column A.
        - func1_folder is created if values in column B are different from those in column A.
        - plant_folder is created within func1_folder based on additional conditions.
        - Handles cases where func1 is equal to func2 and a plant_folder exists.  


- **SB02_Dept**
- **SB03_SecSheets**
- **SB04_Formula&Appearance**
- **SB05_PrintingSetting**


1. Getting the current username of the computer user.
2. Creating a folder named "季獎金切檔" (Seasonal Bonus Cut Files) on the desktop (if it doesn't already exist).
3. Prompting the user to input information about the year and season, for example, "2020Q4".
4. Opening an Excel file named "Year季獎金調整清冊" (Year Seasonal Bonus Adjustment List).
5. Iterating through the data in the Excel file and creating a folder structure based on certain conditions:
6. If the value in column "A" of a row is different from the previous row, a folder with the name "Year季獎金-value" is created.
7. If the value in column "B" of a row is different from the value in column "A", a folder with the name "Year季獎金-value-value" is created.
8. If the value in column "C" of a row is not zero, a folder with the name "Year季獎金調整清冊-value" is created. This code automates the process of creating folders dynamically based on the data in the Excel sheet, allowing for the organized storage of related files.