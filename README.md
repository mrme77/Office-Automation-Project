{
 "cells": [
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "# Office Automation Project"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## Introduction\n",
    "This is a collection of scripts in Visual Basic and Power Shell to automate some daily office tasks and save workhours, every minute counts!\n",
    "All scripts below can be run directly in PowerShell, or execute within a VBS wrapper."
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Prerequisites\n",
    "- PowerShell\n",
    "- Visual Basic Scripts"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Deployment\n",
    "- **Retrieving Computer Info:** \n",
    "It collects the current date, computer name, user domain, serial number, and appendes them to a csv file, <i>WorkStation_Serial_Number.csv</i>). It was created to conduct a computer inventory audit during Covid to avoid direct exposure between auditor and end users.\n",
    "\n",
    "- **Data Sanitization:** \n",
    "It deletes PII data stored in restricted folders with a LastWriteTime attribute older than 365 days. It also creates a log of the files deleted.\n",
    "\n",
    "- **Daily Task Notification Email:**\n",
    "It checks and retrieves a Transaction Report(s). The report(s) is attached to an email which is already populated and send out daily if needed.On Mondays,it will also check for transaction reports arrived over the weekend.\n",
    "\n",
    "- **Replicating Directory Structure Without Files:**\n",
    "It creates an identical directory structure (to include subdirectories) from one location to another rather than creating it manually. The script only replicates the directory structure and excludes any file.\n",
    "\n",
    "- **Daily Reports Allocation:**\n",
    "It distributes daily reports to the appropriate folders. Rather than opening each file received, to determine its destination, end-user can rely on the script to properly identify each report based on their content, and place them in the final destination folder. The script runs on an infinite loop that starts all over every two hours.\n",
    "\n",
    "- **Randomly Select Personnel Records:**\n",
    "It identifies 10 non duplicate personnel records for testing purposes. Records information is stored in a .csv file. Everytime records are identified, they are removed from the main roster and will not be participate in the next random audit.\n",
    "\n",
    "- **Find and Remove Identical Files:** \n",
    "It finds identical files based on the SHA1 algorithm, and moves them to path_duplicates..SHA1 is in a hash or message digest algorithm where it generates 160-bit unique value from the input data. The input data size doesn’t matter as SHA1 always generates the same size message digest or hash which is 160 bit.\n",
    "\n",
    "- **Remove Duplicate Rows Excel Document:**\n",
    "The VBA subroutine removes duplicate rows based on the content of multiple columns.\n",
    "\n",
    "- **Merging Multiple Excel Files:** \n",
    "The VBA subroutine allows user to select to merge multiple Excel Workbook togetether in one document."
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Acknowledgments\n",
    "-  Chapeau to StackOverflow, GitHub, and Google platforms for providing some guidance code."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": []
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.7.6"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 4
}
