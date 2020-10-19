# grade-tracker
Google Apps Script project that will create a dashboard to view the assignment status for each student and create email drafts notifying students and parents of work that is past due.
- Retrieves data from Google Classroom so information is always up to date
- Saves compiled data in Google Sheets for easy viewing and access
- Creates email drafts listing missing assignments and new due date 

Create a copy of the file and preview the dashboard here: https://docs.google.com/spreadsheets/d/10Sb6kbfUxcyGOUjHBLYftv7nn5GYsYkQKZUOizr5qTs/edit?usp=sharing 

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


## Developer Notes
  1. **Count the number of _missing assignments, optional tasks,_ and _extra credit_.** The categories can be defined using the corresponding global variables. 
    - Missing assignments (type 1 or 2) trigger the email draft to be created, while other missing items (type 3 or default) do not. 
    - If a student is missing assignments, any missing extra credit will also be added onto the email draft as an additional support for their grade.
    - At the time of creation, the “Assignment Type” property in Google Classroom was not accessible via API calls so the current approach is to identify keywords within assignment names. Any assignment name that does not include a keyword will be processed automatically as an optional task. All keywords should be lowercase. 
    - Specify the keywords in the corresponding arrays:
      - ASSIGNMENT_TYPE 
      - TYPE#_KEYWORDS
  2. **Display X number of topics.** The number of topics can be restricted in order to reduce the time it takes to generate the Missing Assignment Tracker. The number of topics can be increased to track trends for various students. The default is set to (3) topics being displayed as the intent is that older topics would not be looked at as the semester progresses. 
  3. **Guardian Emails are imported from a separate spreadsheet or can be manually input.** At the time of development, guardian emails were not reliably linked to student profiles. It is important that the spreadsheet with the guardian emails has headers that match the headers in the code. If no spreadsheet is specified then that section of code will be skipped and it is assumed that the information will be manually input or excluded. This information is not required for the email draft to be created. 
  4. **The order of sheets must remain consistent.** The assumption is that classes do not change throughout the semester. If a class is removed for some reason, then the tab must be manually deleted. It is recommended to create an entirely new Google Sheets document for each new semester instead of recycling old ones. 
  5. **Due date format must follow “MM/DD (day of the week)”.** The due date is saved in the Email Hub tab of the spreadsheet. When date validation is run, a string is passed instead of a DateTime object. The string is parsed first by the ‘(‘ and secondly by the ‘/’. 
  6. **Class names have aliases.** Class names have the option to be aliased so students can easily recognize what class is being referenced. This option was developed for grade school educators where it is more recognizable for students to see “history class” instead of a lengthy course name. 
Add as many key-value pairs in COURSE_DICTIONARY as necessary to incorporate all courses. The key should identify the course name and the value is what will be displayed in the email. Leave this as an empty array if the full course name is acceptable.  
  7. **Email draft parsing.** Sections of the email encased in curly brackets {} will be automatically replaced with the corresponding items from the Email Hub tab. The email draft is automatically parsed depending on the user’s selection and missing assignments. 
The email draft is broken into (3) different sections as designated by each “[BREAK]”. The first section is the general email that is always included. The other two sections are optional and depend on the user's choice and student’s assignments’ statuses. If the student is missing extra credit work then the second section will be included and if the user chooses to copy parents onto the email then the last section will be included. 
  8. **Conditional formatting is automatically generated.** Conditional formatting is part of the code in order to adjust based on changing information locations. See the “Google Sheets design” section for more information on each formatting rule. 

#### Design Notes:
  1. **Email Hub Tab** - Information for email draft is pulled from here. This tab must remain as the first in the spreadsheet. If the number of students changes then the roster will be automatically updated with the next generation of the spreadsheet. This sheet does not get cleared once it is generated. If a student’s information, such as email address, needs to be updated then the user must manually make the change.  
  - Students are grouped by class and listed in alphabetical order by last name
  2. **Master Roster Tab** - Compilation of all students w/ quick summary of missing assignments. The number of missing assignments as well as a list of missing assignments are listed here for the user to quickly get a sense of how many students are missing assignments or how many assignments a student is missing. The information in this tab is cleared and refreshed every time the grade report is generated. 
  - Students are grouped by class and listed in alphabetical order by last name 
  - The ratio is the number of missing assignments to total number of assignments of that category
  - Assignments are separated into three categories: Homework, Classwork, Extra Credit
  - The “Unit: MW” column displays the list of both missing homework, classwork, and extra credit
  - Conditional Formatting:
    - Green - If student is not missing any assignments
    - Red - Student is missing more than half of the total assignments
    - Orange - Student is missing at least one assignment, but not more than half of total assignments
  3. **Class Tab(s)** - Name of tab will have the class name and section. The sheet will contain the full breakdown of assignment statuses for each student. These tabs are intended for the user to be able to see the submission status for each assignment. This can be helpful for determining which assignments still need to be graded. 
  - Students are listed in alphabetical order by last name
  - Unit name in the first row spans the number of columns equal to the number of assignments + 1. The empty grey column designates the end of a unit and beginning of a different unit
  - Due date displays the date after the assignment is due (Google classroom is zero indexed) 
  - Conditional formatting & legend
    - Light green - “SUBMITTED” The assignment has been submitted on time and has not been returned with a grade.
    - Dark green - “GRADED” The assignment has been returned with a grade.
    - Light orange - “NOT YET DUE” The assignment has not been submitted and deadline has not passed yet or no deadline has been specified
    - Dark orange - “MM/DD - LATE” The assignment has been submitted late (past the 1 hour grace period of the deadline). 
    - Light red - “NOT YET SUB” The classwork has not been submitted and the deadline has passed.
    - Dark red - “MISSING” The homework has not been submitted and the deadline has passed.
    - Blue - “EXTRA CREDIT” or “OPTIONAL” Optional assignment, does not get counted as missing.
  4. **createCourseRosters()** requests data from Google Classroom and organizes it into Google Sheets.
  - Prints to class tabs and “Master Roster” tab via appendRow() thus the sheets are cleared every time the function is called
  - Prints to “Email Hub” tab via setValue() since Email Hub should never be cleared
  - Records class sizes prior to sheets being cleared in order to detect any changes in number of students
  - Iterates through classes in order that they were created in Google Classroom
  - The variable rawAssignments is an array containing json objects of every assignment in the course ordered by last update time. The array is filtered by topicId to properly categorize each assignment. The array is then sorted by creation time so that the order matches what is displayed in Google Classroom.
  - Topics with unpublished assignments will appear in rawTopics, but the assignments will not appear in rawAssignments. So if the topic does not have any assignments then it is not counted as a valid topic in numValidTopics.
  - *allAssignmentIds* is used to get the assignment statuses. It is created in this section so that the array is aligned with the order of assignments.
Spaces are added to each array since the spreadsheet format has empty columns in between each topic.
  - Assignment statuses were not accessible per student during development. Query items were limited to assignmentID and courseID and the result would be an array of studentSubmissionStatus json objects. The approach is to create an array containing all of the studentSubmissionStatus objects and then filter by student’s userId.
  - Changes in the number of students are checked after each class tab has been created. If the number of students increased then, the new student is identified by iterating through the classTab and corresponding section of the Email Hub tab. When a difference is found then a new row will be inserted at the top of the section in the Email Hub tab and that section will be sorted by last name. The same approach is taken when the number of students decreases except the row in the Email Hub tab will be deleted instead. 
    - The first four columns of Master Roster will be copied to Email Hub the first time the function is run. The assumption is that nothing has been input to Email Hub so a direct copy of the roster is accurate.
  - The date the function is run is saved into “Email Hub” to be displayed in the HTML control panel and the sheets are formatted. 
  5. **getIndexOf** finds the index of searchParam within a two dimensional array, twoDimsArray. Using getValues() on a row in the spreadsheet returns a two dimensional array so this function is used to find headers in the spreadsheet. 
  6. **getGuardianEmails** imports guardian email addresses from the specified Google Sheet ID “IMPORT_GUARDIAN_EMAILS_FROM”. If no ID number is specified then the function returns a blank array. Otherwise the function will search for the student’s email address and return the corresponding guardian emails. The specified spreadsheet must use the following headers:
  - Email address 
  - Guardian Email 1 
  - Guardian Email 2
  7. **generateEmailDrafts** compiles information and creates email drafts
  - Information that replaces the blank values in the email template are compiled in replaceWithThese. The information is pushed onto the array in the order that they appear in the email template. 
  - *listOfMissingAssignments* is split into the three categories so that the user can easily change whether or not they want to include classwork. The extra newlines are removed for email formatting.
  - To include other emails as a copy, the input needs to be a string with emails separated by commas. Guardian emails are saved in two separate columns so they need to be checked and formatted
  - The email body changes depending on user input. The key string [BREAK] is used so that if additional sections are needed they can be easily integrated. 
  - If there was no email draft previously generated then a new draft will be created
  - If the previous email draft was sent out then the draftId is no longer valid and a new draft will be created
  - Otherwise the draft will be retrieved and updated

