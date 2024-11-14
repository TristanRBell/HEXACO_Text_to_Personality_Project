# HEXACO_Analysis_Final
The Final Repo for a HEXACO personality measure Asynchronous Video Interview Transcript Analysis project.

# Process For Using the Script
Download the **.py file (Now Running V2 as main file), P300, P301, 302 & 304 text-files and the excel files** and place them in an accessible folder. Run a Test prior to further analysis (**Refer to 'Package Installation', 'Running the Script' & 'Pre-Data Collection Test' prior to the next step**). 
Import the .txt files used for analysis into the folder you created, where there are an assumed 5 files per PID, named as such "P--- Q-.txt". Ensure that all .txt files are named in the same format. If the .txt files are named something different, alter line 191195 in the V2.py script to accomodate that change (for instance, take out the space at the begining of the quotation marks).

# Package Installations
You will need to install openpyxl and pandas. the package re is inbuilt in Python and does not require install. Refer to the first 3 lines of the script for install instructions, if these fail, refer to your IDE specific package install instructions and types the necessary lines into the terminal.

# Running the Script
**Ensure that the Output file is closed when running the script (named 'hexaco_output_final_norm_raw_auto.xlsx').

Once pre-prep has been done, all necessary files are in the folder, named correctly, and the script has been adjusted to your naming parameters, hit run and enter the PID numbers as directed, making sure that line 186-190 match the format of your own PID naming in your .txt files. 
You should now have a score across all 6 dimensions of the hexaco (Both normed and raw in a wide format), with the PID recorded, in the output file (hexaco_output_final_norm_raw_auto.xlsx). Check this before continuing with data collection. 
For all further data collection, ensure the output file is closed, and run as many times as you need to complete all PIDs.

# Pre-Data Collection Test
Test Running via usage of P300 & P3001 (test files contained in the Repo) - output should be as follows
PID|H_Norm     |E_Norm     |X_Norm     |A_Norm     |C_Norm     |O_Norm     |H_Raw      |E_Raw       |X_Raw      |A_Raw      |C_Raw      |O_Raw
300|0,530266815|0,485254007|0,517839054|0,510152395|0,554038408|0,475561844|0,057072402|-0,005919966|0,039680594|0,028923637|0,090339132|-0,019483491
301|0,530266815|0,485254007|0,517839054|0,510152395|0,554038408|0,475561844|0,057072402|-0,005919966|0,039680594|0,028923637|0,090339132|-0,019483491


If the following output is attained, the script and all dictionary & output files are running as intended. For testing using pid range 300-304 (excluding 303), all outputs should be the same.

# Development Pipeline
# 29th Sep 2024 - New Additions
Automation of the code and concurrent running and recording of raw and normed dictionary scores are implemented. The user must now enter starting and ending PID numbers, and the script will itterate through all PID numbers within the given range and output normed and raw scores in a wide format per PID.
Additional bug finding and fixing was completed, such that "Speaker 1" was not being removed due to an upper-lower case issue, and 'speaker' is contained in the dictionary. This was leading to a score inflation in some traits, and a reduction in others. This was resolved by lowering all text first, and then replacing "speaker 1" with a blank space. This has corrected the issue, and new, accurate scores checked, and confirmed. No other score confounds are suspected, such as punctuation confounds etc. 

# Future Additions
**Additions 1 & 2 are to be completed by the 13th of October 2024**
**Additions 1 & 2 Completed - New Version is V2**
Accounting for Non-intiger values input by the user both through additional code comments, and input re-requests if the user fails to provide an intiger input.
Try statements for document read-in to account for non-continuous intiger values for PID's within data source, e.g., PID 5 & 12 are excluded from the data source and are therefore passed over without script failure as a failure to read in a value just iterates to the next value in the PID range input. 
**Addition 3 & 4 is to be completed if Requested by the Client**
Further considerations for outputs include dictionary key excerpts for appendices and sample key:value outputs for the same purpose.
The Usage of Mean Score values per HEXACO dimension may need revision, if scores across the dataset lack variance. This may necessitate more complex calculation of final scores to manage the tendency of scores to cluster around a central point.
