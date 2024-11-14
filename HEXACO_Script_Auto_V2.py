# Import libraries for Excel-file writing and manipulation and creating data-frames of HEXACO scores
# Install using Terminal by entering the below lines respectively
# (py -m pip install openpyxl), (py -m pip install pandas)

import openpyxl
import pandas as pd
import re as re

# Functions are defined here - Collapse these functions with the downwards arrow on the left, assuming they are open.

# This function reads in the HEXACO dictionaries (raw and normed score variants)
def read_in_dict(dict_file):
    df = pd.read_excel(dict_file, header=None)
    D_H = dict(zip(df[0], df[1]))
    D_E = dict(zip(df[0], df[2]))
    D_X = dict(zip(df[0], df[3]))
    D_A = dict(zip(df[0], df[4]))
    D_C = dict(zip(df[0], df[5]))
    D_O = dict(zip(df[0], df[6]))
    
    df2 = pd.read_excel(dict_file_raw, header=None)
    D_H_R = dict(zip(df2[0], df2[1]))
    D_E_R = dict(zip(df2[0], df2[2]))
    D_X_R = dict(zip(df2[0], df2[3]))
    D_A_R = dict(zip(df2[0], df2[4]))
    D_C_R = dict(zip(df2[0], df2[5]))
    D_O_R = dict(zip(df2[0], df2[6]))

    return D_H, D_E, D_X, D_A, D_C, D_O, D_H_R, D_E_R, D_X_R, D_A_R, D_C_R, D_O_R


# This function reads in the transcript text files
def read_in_txt(t1, t2, t3, t4, t5):
    with open(t1 , 'r') as Input, open('t1_output.txt', 'w+') as Output:
        lines = Input.readlines()
        cleaned = [line.strip("\n").lower().replace("speaker 1", " ").replace(".", ' ').replace(",", " ").replace("'", "")
                   .replace("!", " ").replace("@", ""). replace(":", "") for line in lines]
        joined = ' '.join(cleaned)
        Output.write(joined)
    with open('t1_output.txt', '+r') as Text1:
        for i in Text1:
            Cleaned_Text1 = re.sub('[!@#$.,]', '', i)
    Analysis_Text1 = Cleaned_Text1.split(" ")

    with open(t2 , 'r') as Input, open('t2_output.txt', 'w+') as Output:
        lines = Input.readlines()
        cleaned = [line.strip("\n").lower().replace("speaker 1", " ").replace(".", ' ').replace(",", " ").replace("'", "")
                   .replace("!", " ").replace("@", ""). replace(":", "") for line in lines]
        joined = ' '.join(cleaned)
        Output.write(joined)
    with open('t2_output.txt', '+r') as Text2:
        for i in Text2:
            Cleaned_Text2 = re.sub('[!@#$.,]', '', i)
    Analysis_Text2 = Cleaned_Text2.split(" ")

    with open(t3 , 'r') as Input, open('t3_output.txt', 'w+') as Output:
        lines = Input.readlines()
        cleaned = [line.strip("\n").lower().replace("speaker 1", " ").replace(".", ' ').replace(",", " ").replace("'", "")
                   .replace("!", " ").replace("@", ""). replace(":", "") for line in lines]
        joined = ' '.join(cleaned)
        Output.write(joined)
    with open('t3_output.txt', '+r') as Text3:
        for i in Text3:
            Cleaned_Text3 = re.sub('[!@#$.,]', '', i)
    Analysis_Text3 = Cleaned_Text3.split(" ")

    with open(t4 , 'r') as Input, open('t4_output.txt', 'w+') as Output:
        lines = Input.readlines()
        cleaned = [line.strip("\n").lower().replace("speaker 1", " ").replace(".", ' ').replace(",", " ").replace("'", "")
                   .replace("!", " ").replace("@", ""). replace(":", "") for line in lines]
        joined = ' '.join(cleaned)
        Output.write(joined)
    with open('t4_output.txt', '+r') as Text4:
        for i in Text4:
            Cleaned_Text4 = re.sub('[!@#$.,]', '', i)
    Analysis_Text4 = Cleaned_Text4.split(" ")

    with open(t5 , 'r') as Input, open('t5_output.txt', 'w+') as Output:
        lines = Input.readlines()
        cleaned = [line.strip("\n").lower().replace("speaker 1", " ").replace(".", ' ').replace(",", " ").replace("'", "")
                   .replace("!", " ").replace("@", ""). replace(":", "") for line in lines]
        joined = ' '.join(cleaned)
        Output.write(joined)
    with open('t5_output.txt', '+r') as Text5:
        for i in Text5:
            Cleaned_Text5 = re.sub('[!@#$.,]', '', i)
    Analysis_Text5 = Cleaned_Text5.split(" ")
    complete_word_list = Analysis_Text1 + Analysis_Text2 + Analysis_Text3 + Analysis_Text4 + Analysis_Text5
    return complete_word_list

# This function scores the words occuring in the transcript
def word_score(cleaned_list):
    # Pre-defines Lists per PID to append to these lists later on
    H_list = []
    E_list = []
    X_list = []
    A_list = []
    C_list = []
    O_list = []
    H_R_list = []
    E_R_list = []
    X_R_list = []
    A_R_list = []
    C_R_list = []
    O_R_list = []
    for i in cleaned_list:  
        if i in D_H:
            H_list.append(D_H.get(i))
            E_list.append(D_E.get(i))
            X_list.append(D_X.get(i))
            A_list.append(D_A.get(i))
            C_list.append(D_C.get(i))
            O_list.append(D_O.get(i))
            H_R_list.append(D_H_R.get(i))
            E_R_list.append(D_E_R.get(i))
            X_R_list.append(D_X_R.get(i))
            A_R_list.append(D_A_R.get(i))
            C_R_list.append(D_C_R.get(i))
            O_R_list.append(D_O_R.get(i))
        else:
            continue
    return [H_list, E_list, X_list, A_list, C_list, O_list, H_R_list, 
            E_R_list, X_R_list, A_R_list, C_R_list, O_R_list]

# This function averages the occuring word-scores per HEXACO dimension and outputs a final score
def score_manipulate(H_list, E_list, X_list, A_list, C_list, O_list, 
    H_R_list, E_R_list, X_R_list, A_R_list, C_R_list, O_R_list):
    H_Var = sum(H_list)/len(H_list)
    E_Var = sum(E_list)/len(E_list)
    X_Var = sum(X_list)/len(X_list)
    A_Var = sum(A_list)/len(A_list)
    C_Var = sum(C_list)/len(C_list)
    O_Var = sum(O_list)/len(O_list)
    H_R_Var = sum(H_R_list)/len(H_R_list)
    E_R_Var = sum(E_R_list)/len(E_R_list)
    X_R_Var = sum(X_R_list)/len(X_R_list)
    A_R_Var = sum(A_R_list)/len(A_R_list)
    C_R_Var = sum(C_R_list)/len(C_R_list)
    O_R_Var = sum(O_R_list)/len(O_R_list)
    return H_Var, E_Var, X_Var, A_Var, C_Var, O_Var, H_R_Var, E_R_Var, X_R_Var, A_R_Var, C_R_Var, O_R_Var

# This function outputs final HEXACO scores per participant to an Excel file 
# Make sure the excel file is closed during any running of the script.
def output_score(path, pid, H_Var, E_Var, X_Var, A_Var, C_Var, O_Var, H_R_Var, E_R_Var, X_R_Var, A_R_Var, C_R_Var, O_R_Var):
    wb_obj = openpyxl.load_workbook(path)
    sheet_obj = wb_obj.active

    # Add data to the Excel sheet
    data = [
        [pid, H_Var, E_Var, X_Var, A_Var, C_Var, O_Var, H_R_Var, E_R_Var, X_R_Var, A_R_Var, C_R_Var, O_R_Var],
    ]
    for row in data:
        sheet_obj.append(row)
    # Save the workbook to a file
    wb_obj.save("hexaco_output_final_norm_raw_auto.xlsx")
    # Print a success message
    print("PID", pid, "has been processed successfully")



## Main

# Input the excel file name into 'path' as "excel_file_name.xlsx", that you want to export to.
# This excel file must already be created and in the working folder. 
# Ensure this file is closed before running the script.
path = "hexaco_output_final_norm_raw_auto.xlsx"

# Sources the Dictionary file-path. If using the same dictionary across trials(assumed),
# replace next line with dict_file = "dictionary_file.xlsx", where dictionary file is whatever your dictionary is called.
dict_file = "hexaco_dict_global_0-1-norm.xlsx"
dict_file_raw = "hexaco_dict_raw.xlsx"

# Predefines the HEXACO dictionaries of words and their respective key:Value contents.
D_H, D_E, D_X, D_A, D_C, D_O, D_H_R, D_E_R, D_X_R, D_A_R, D_C_R, D_O_R = read_in_dict(dict_file)

# Sources the PID numbers from the script user.
# Enter this value as just an integer (format "300", rather than "P300").
while True:
    start_pid = input("What PID integer value are you starting with (e.g., 1)? ")
    end_pid = input("What PID integer value are you ending with (e.g., 100)? ")
    try:
        start_pid = int(start_pid)
        end_pid = int(end_pid)
    except ValueError:
        print("Invalid entry - Please enter an integer (format e.g., '1')")

    for pid in range(start_pid, end_pid+1):
        try:
            # Acquires a referent textfile path from which to read contents as a list for manipulation and cleaning.
            # Inspect file names and ensure appropriate naming below - e.g., " Q1.txt" or "Q1.txt" (space/no-space before Q).
            t1 = "P"+str(pid)+" Q1.txt"
            t2 = "P"+str(pid)+" Q2.txt"
            t3 = "P"+str(pid)+" Q3.txt"
            t4 = "P"+str(pid)+" Q4.txt"
            t5 = "P"+str(pid)+" Q5.txt"

            # Compile word list using all 5 text-files and clean it of non-content words to make processing faster.
            complete_word_list = read_in_txt(t1, t2, t3, t4, t5)
            #print(complete_word_list)
            cleaned_list = [x for x in complete_word_list if x in D_H]
            print(cleaned_list)

            # Iterate through the cleaned list for each dictionary key and check congruence between list item i
            # and each dictionary, grab value in respective dictionaries and save to a list per variable.
            H_list, E_list, X_list, A_list, C_list, O_list, H_R_list, E_R_list, X_R_list, A_R_list, C_R_list, O_R_list = word_score(cleaned_list)


            # Sum and average the scores in each variable list, to produce an averaged expression of a HEXACO domain to,
            # calculate the PID HEXACO scores by dimension and occurence rate & print all HEXACO scores for each dimension
            H_Var, E_Var, X_Var, A_Var, C_Var, O_Var, H_R_Var, E_R_Var, X_R_Var, A_R_Var, C_R_Var, O_R_Var = score_manipulate(H_list, E_list, X_list, A_list, C_list, O_list, H_R_list, E_R_list, X_R_list, A_R_list, C_R_list, O_R_list)

            # Open a pre-specified excel file and output the participant's HEXACO scores to the excel document.
            # Output in format 'PID|H|E|X|A|C|O|H|E|X|A|C|O', 
            # where the first HEXACO output is normed scores, and the second is raw scores
            output_score(path, pid, H_Var, E_Var, X_Var, A_Var, C_Var, O_Var, H_R_Var, E_R_Var, X_R_Var, A_R_Var, C_R_Var, O_R_Var)
        except: continue