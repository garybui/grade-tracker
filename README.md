# grade-tracker
Google Apps Script project that will create a dashboard to view the assignment status for each student and create email drafts notifying students and parents of work that is past due.
- Retrieves data from Google Classroom so information is always up to date
- Saves compiled data in Google Sheets for easy viewing and access
- Creates email drafts listing missing assignments and new due date 

Preview the dashboard here: https://docs.google.com/spreadsheets/d/10Sb6kbfUxcyGOUjHBLYftv7nn5GYsYkQKZUOizr5qTs/edit?usp=sharing (The control panel is not accessible in view only mode)

## Installation

### Setting Up the Files
Follow these steps to import the files into Google Apps Script
  1. Download the included files
  2. With the Google account that creates the Google Classroom(s), create a new Google Sheet
  3. Open the Script Editor: **Tools > Script editor**
  4. Delete the current contents of “Code.gs”
  5. Copy the code from the downloaded *“main.gs”* file and paste into *“Code.gs”*
  6. We need to enable permissions for the spreadsheet. Open the manifest file: **View > Show manifest file**
    - If the project has not been named then a dialog box will appear
  7. A new file should appear in the left column called *“appsscript.json”* 
  8. Delete the current contents of the file
  9. Copy the code from the downloaded *“permissions.txt”* file and paste into *"appsscript.json" 
    - See development notes below for a full list of permissions required
  10. From the toolbar go to **File > New > HTML file** and create a new HTML file named *“controlpanel.html”* 
  11. Clear the current contents of the newly created *“controlpanel.html”* 
  12. Copy the code from the downloaded *“controlpanel.html”* and paste into the new *“controlpanel.html”*
  13. Repeat steps 9-11 to create a new “popup.html” using the downloaded “popup.html”
  14. Save the project: **File > Save All**
  
### Including the Required Information in *main.gs*
Follow these steps to fill out the required identification numbers for the code to function properly
  1. Run the function “getOwnerId”: **Run > Run Function > getOwnerId**  
    - A popup will appear to request access to your Google Account. Click “Review Permissions” to authorize access.
  2. Access the logs to view your Google Classroom owner ID: **View > Logs**
    - There may be several different Owner IDs. Identify one class that you created and copy the corresponding ID number. 
  3. Paste the 21 digit ID number into line 2 of *main.gs*
    - OWNER_ID = *‘PASTE OWNER ID HERE”*
  4. Create a new Google Document with the contents of the downloaded file *“EmailTemplate.txt”* 
  5. Identify the document ID through the URL of the newly created Google Doc. Copy the document ID in-between *“docs.google.com/document/d/”* and *“/edit”*. Do not include the backslashes (/).
    - Users should create their own email template Google Document with the provided email template. See developer notes for more information on how to customize the template. 
  6. Paste the document ID into line 3 of *main.gs*
    - EMAIL_TEMPLATE_ID = *“PASTE DOCUMENT ID HERE”*
  7. Optional: To include parents on emails about missing assignments, repeat steps 4 and 5 for the Google Spreadsheet with guardian email addresses. See developer notes below for more information.
  8. Optional: To create course name aliases, specify key-value pairs for COURSE_DICTIONARY on line 8 of Code.gs. See developer notes below for more information. 
  9. Save the project: **File > Save All

### Final Steps
  1. Create a trigger to open the control panel automatically: **Edit > Current Project Triggers 
    - Press the “Add Trigger” button on the bottom right of the screen. 
    - Trigger settings:
      - Choose which function to run: whenOpen
      - Which runs at demployment: Head
      - Select event source: From spreadsheet
      - Select event type: On open
      - Failure notification settings: Notify me daily
  2. Return to the script editor. Run the function “initalizeSpreadsheet”: **Run > Run Function > initializeSpreadsheet  
    - When the function is complete, the following tabs should be created in the Google Spreadsheet:
      - Email Hub
      - Master Roster
  3. Setup is complete! Access the Control Panel from the Spreadsheet: **Add-ons > [Project Name] > Control Panel

  
