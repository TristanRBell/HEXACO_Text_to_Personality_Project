# HEXACO_Analysis_Final
The Final Repo for a HEXACO personality measure Asynchronous Video Interview Transcript Analysis project.

# Process For Using the Script
Download the .py file and the excel files and place them in an accessible folder.
Import the .txt files used for analysis into the folder you just created, where there are an assumed 5 files per PID, named as such "P--- Q-.txt". Ensure that all .txt files are named in the same format.
If the .txt files are named something different, alter line 162-166 in the .py script to accomodate that change (for instance, take out the space at the begining of the quotation marks).

# Package Installations
You will need to install openpyxl and pandas. the package re is inbuilt in Python and does not require install. Refer to the first 3 lines of the script for install instructions, if these fail, refer to your IDE specific package install instructions and types the necessary lines into the terminal.

# Running the Script
**Ensure that the Output file is closed when running the script (named hexaco_output_final.xlsx).
Once pre-prep has been done, all necessary files are in the folder, named correctly, and the script has been adjusted to your naming parameters, hit run and enter the PID number as directed, making sure to match the format of your own PID naming in your .txt files.
You should now have a score across all 6 dimensions of the hexaco, with the PID recorded, in the output file (hexaco_output_final.xlsx). Check this before continuing with data collection.
For all further data collection, ensure the output file is closed, and run as many times as you need to complete all PIDs.

