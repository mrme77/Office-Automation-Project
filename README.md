# Office-Automation-Project
This is a collection of PowerShell and VBA scripts created/modified to automate day to day office taskers and save work hours. All PowerShell scripts can be run directly in PowerShell, or execute within a VBS wrapper.
### Prerequisites
- PowerShell
- Visual Basic Scripts
### Deployment
- **Retrieving Computer Info:** 
It collects the current date, computer name, user domain, serial number, and appendes them to a csv file, <i>WorkStation_Serial_Number.csv</i>). It was created to conduct a computer inventory audit during Covid to avoid direct exposure between auditor and end users.
- **Data Sanitization:**
It deletes PII data stored in restricted folders with a LastWriteTime attribute older than 365 days. It also creates a log of the files deleted.
- **Daily Task Notification Email:**
It checks and retrieves a Transaction Report(s). The report(s) is attached to an email which is already populated and send out daily if needed.On Mondays,it will also check for transaction reports arrived over the weekend.
- **Replicating Directory Structure Without Files:**
It creates an identical directory structure (to include subdirectories) from one location to another rather than creating it manually. The script only replicates the directory structure and excludes any file.
- **Randomly select personnel records:**
It identifies 10 non duplicate personnel records for testing purposes. Records information is stored in a .csv file. Everytime records are identified, they are removed from the main roster and will not be participate in the next random audit. 
- **Find and Remove Identical Files:**
It finds identical files based on the SHA1 algorithm, and moves them to <i> path_duplicates.</i>.SHA1 is in a hash or message digest algorithm where it generates 160-bit unique value from the input data. The input data size doesn’t matter as SHA1 always generates the same size message digest or hash which is 160 bit.
- **Remove Duplicate Rows Excel Document:**
The VBA subroutine removes duplicate rows based on the content of multiple columns.
- **Merging Multiple Excel Files:**
This VBA subroutine below allows user to merge multiple Excel Workbooks in one document.
### Acknowledgments
- Chapeau to StackOverflow, GitHub, and Google platforms for providing some guidance code.
