import csv

from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
import os
import pandas as pd

import haematology

# Import LST excel file into pandas dataframe
def read_lst(path):
    df = pd.read_excel(path, header=0)
    return df

# Import Summary Services excel file into pandas dataframe
def read_summary_services(path):
    # Get sheet names
    try:
        xl = pd.ExcelFile(path)
        sheet_names = xl.sheet_names
        # Read sheet with name 'Interpretation'
        if 'Interpretation' in sheet_names:
            df = pd.read_excel(path, 'Interpretation', header=0)
        else:
            return False
        return df
    except:
        return False

# Read folders to get a list of Excel xls files recursively into a dictionary
# Input: path to folder
# Output: dictionary of folder names and a list containing the path and the name of the file
def read_folders(path):
    folder_dict = {}
    for root, dirs, files in os.walk(path):
        # Open subfolders
        for dir in dirs:
            # Read files in subfolders
            for root2, dirs2, files2 in os.walk(path + '\\' + dir):
                # Loop through files with index
                for index, file in enumerate(files2):
                    # Check if file is an excel file
                    if file.endswith('.xls'):
                        # Check if key exists
                        if dir in folder_dict:
                            # Check the modification time of the file
                            if os.path.getmtime(root2 + '\\' + file) > os.path.getmtime(folder_dict[dir][0] + '\\' + folder_dict[dir][1]):
                                # Update the value
                                folder_dict[dir] = [root2, file]
                        else:
                            # Add new key
                            folder_dict[dir] = [root2, file]
    return folder_dict


# Determine whether Haematology is contained within a text string
# Input: text string
# Case insensitive
# Return True if it is, False if it is not
def is_haematology(text):
    if 'haematology' in text.lower():
        return True
    else:
        return False


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # define local service tracker information
    LSTPath = '\\\\ctc-network.intranet\\dfs\\BIOT\\01 Study Management\\02 Trackers\\Local Services Tracker.xlsm'
    LST = read_lst(LSTPath)
    # define summary services information
    BPATProjectFolder = '\\\\ctc-network.intranet\\dfs\\BPATCR\\Contract-Archive\\Project'
    # Define unique test names for haematology
    HaematologyTests = []
    # Read the first level of BPATProjectFolder
    for SponsorFolder in os.listdir(BPATProjectFolder):
        ServiceFolder = BPATProjectFolder + '\\' + SponsorFolder
        ServiceFolderDict = read_folders(ServiceFolder)
        # Loop through dictionary
        for key, value in ServiceFolderDict.items():
            # Extract first digits before alphabet from key and then stop
            digits = ''
            for i in range(len(key)):
                if key[i].isdigit():
                    digits += key[i]
                else:
                    break
            # Read summary services excel file
            ServicePath = value[0] + '\\' + value[1]
            # Split value by underscore
            split_string = value[1].split('_')
            # Get the item from split_string that contains key
            CtcNo = key
            for item in split_string:
                if key in item:
                    CtcNo = item
            if CtcNo == key:
                CtcNo = key + 'HKU1'
            print("Loading " + CtcNo + "...")
            # Check file exists
            if os.path.exists(ServicePath):
                ServiceDf = read_summary_services(ServicePath)
                # print(ServiceDf[0:10])
                if ServiceDf is False:
                    print("Service interpretation not found in Summary Services")
                    continue
                # Check LST
                # Select row by key in LST
                site = LST.loc[LST['CTC No.'] == CtcNo]
                if site.empty:
                    print('[' + CtcNo + '] Site not found in LST')
                # Loop through rows in Service
                # Flag for haematology test
                HasHaematology = False
                for index, row in ServiceDf.iterrows():
                    if is_haematology(row[1]) or is_haematology(row[0]):
                        # Update flag
                        HasHaematology = True
                        # Print item name
                        print(row[0])
                        # Clean row[1] using nltk
                        # Remove stop words
                        stop_words = set(stopwords.words('english'))
                        word_tokens = word_tokenize(row[1])
                        filtered_sentence = [w for w in word_tokens if not w in stop_words]
                        # Remove punctuation
                        filtered_sentence = [w for w in filtered_sentence if w.isalnum()]
                        # Interpret tests
                        # Coagulation tests
                        if 'coagulation' in row[0].lower():
                            # Match coagulation tests
                            CTest = haematology.match_coagulation_tests(row[0], filtered_sentence)
                            print(CTest)
                            # Add all tests to HaematologyTests if not exists
                            for test in CTest:
                                # trim test
                                test2 = test.strip()
                                if test2 not in HaematologyTests:
                                    HaematologyTests.append(test2)
                        elif 'examination of blood film' in row[0].lower():
                            BloodFilms = haematology.match_blood_film_examination(row[0], filtered_sentence)
                            print(BloodFilms)
                            # Add all tests to HaematologyTests if not exists
                            for test in BloodFilms:
                                # trim test
                                test2 = test.strip()
                                if test2 not in HaematologyTests:
                                    HaematologyTests.append(test2)
                        elif 'haematology' in row[0].lower() and 'count' in row[1].lower():
                            CBC = haematology.match_complete_blood_picture(row[0], filtered_sentence)
                            print(CBC)
                            # Add all tests to HaematologyTests if not exists
                            for test in CBC:
                                # trim test
                                test2 = test.strip()
                                if test2 not in HaematologyTests:
                                    HaematologyTests.append(test2)
                        else:
                            row[0] = row[0].strip()
                            HaematologyTests.append(row[0])
                            #print(filtered_sentence)
                if HasHaematology is False:
                    print("No haematology test found")
                elif not site.empty:
                    # Check whether column 8 is empty
                    if pd.isnull(site.iloc[0, 8]):
                        print('Haematology lab not set')
            else:
                print("File not found: " + ServicePath)

    # Print all haematology tests
    # print(HaematologyTests)
    # Export into csv
    # with open('HaematologyTests.csv', 'w', newline='') as f:
    #    writer = csv.writer(f)
    #    writer.writerow(HaematologyTests)