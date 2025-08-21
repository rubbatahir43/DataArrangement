1.	Data Arrangement
This project is a desktop application used for arranging the data in a particular format. In this project 4 inputs have been provided by the user
a.	The reference excel file
b.	The data source folder
c.	The data destination folder (also arranging the data in the particular requirement)
d.	The template file path
The tool was designed to automate the process of making excel file as a reference, gets the data from the excel and arrange the data in a particular format at the destination folder
2.	Features
This tool automatically does the following tasks
•	Reads the data from the excel
•	Extract Names and Service Dates from the excel
•	Handles multiple date formats
•	Converts the date to a specific folder path year/month year/ mm.dd.yyyy
•	Navigates a source folder hierarchy 
•	Creating necessary changes in the word file
•	Filtering Excel rows as per requirement
•	Finds date folders with exact patient name matching
•	Copies them to the destination folder
•	Creates an excel mentioning the missing data against the names 
3.	Tech Stack
•	Python 3.x**
•	Pandas for excel data handling
•	Openpyxl for excel reading and writing
•	Os and Shutil for data copy pasting
•	Tkinter for creating GUI
•	Re for using Regular Expressions
•	From docx to do changes in the word template file
4.	Example Workflow
1.	It will make sure that all the inputs have been provided by the user, if anyone is missing it will not continue
2.	It will make sure that are all valid paths, if there is no directory like that it will create
3.	It has sub functions which are mentioned against checkboxes only required functions can also be performed as per requirement. The checkboxes include
a.	Make Folders
b.	Filter Excel Files
c.	Search for ERAS
d.	Rename Word files
e.	Place Data Folders
f.	Rename Data
4.	On clicking the Process Files Button, it will start processing. It will also show the Console output as it was the requirement at that time.
5.	So first it will open the excel file and get the values of service date, names, claim no, facility, type and dispute id
6.	It will perform all the functions as per the checkboxes selected and also create the log file of missing data
7.	Further it will do the editing in the ms word files and and arrange the data at the destination folder in a particular requirements and also rename the files.
5.	Future Enhancements:
•	The GUI is in basic form, the GUI could be made better
•	The data can be uploaded to a particular location (e.g. Dropbox) as per requirement
•	Enable database integration for auto tracking
6.	Author
•	Rubba Tahir
•	rubbatahir43@gmail.com