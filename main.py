import csv

from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
import haematology
import os
import pandas as pd


# Import LST excel file into pandas dataframe
def read_lst(path):
    df = pd.read_excel(path, header=0)
    return df

# Load HaematologyTestsDB into pandas
def load_haematology_tests_db(path):
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
                    # Check if file is an Excel file
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
    # Load HaematologyTestsDB
    HaematologyTestsDBPath = '.\\HaematologyTestsDB.xlsx'
    HaematologyTestsDB = load_haematology_tests_db(HaematologyTestsDBPath)
    # define local service tracker information
    LSTPath = '\\\\ctc-network.intranet\\dfs\\BIOT\\01 Study Management\\02 Trackers\\Local Services Tracker.xlsm'
    LST = read_lst(LSTPath)
    # define summary services information
    BPATProjectFolder = '\\\\ctc-network.intranet\\dfs\\BPATCR\\Contract-Archive\\Project\\Imago'
    # Define unique test names for haematology
    HaematologyTests = []
    HaematologyTestsForCtcNo = {}
    # Create pandas data frame to store service information
    # ServiceList = pd.DataFrame(columns=['Service', 'Test Interpretation'])
    # Read the first level of BPATProjectFolder
    #for SponsorFolder in os.listdir(BPATProjectFolder):
    #    ServiceFolder = BPATProjectFolder + '\\' + SponsorFolder
    ServiceFolder = BPATProjectFolder
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
            CurrentHaematologyTest = []
            for index, row in ServiceDf.iterrows():
                if is_haematology(row[1]) or is_haematology(row[0]):
                    # Update flag
                    HasHaematology = True
                    # Print item name
                    print(row[0])
                    # Append to ServiceList
                    # ServiceList = ServiceList.append({'Service': row[0], 'Test Interpretation': row[1]}, ignore_index=True)
                    # Clean row[1] using nltk
                    # Remove stop words
                    stop_words = set(stopwords.words('english'))
                    word_tokens = word_tokenize(row[1])
                    filtered_sentence = [w for w in word_tokens if not w in stop_words]
                    # Remove punctuation
                    filtered_sentence = [w for w in filtered_sentence if w.isalnum()]
                    # To lower case
                    filtered_sentence = [w.lower() for w in filtered_sentence]
                    # Interpret tests
                    # Match Coagulation tests
                    CTest = haematology.match_coagulation_tests(row[0], filtered_sentence, CurrentHaematologyTest)
                    # if CTest is not empty
                    if len(CTest) > 0:
                        print(CTest)
                        # Add all tests to HaematologyTests if not exists
                        for test in CTest:
                            # trim test
                            test2 = test.strip()
                            CurrentHaematologyTest.append(test2)
                    # Blood Film Examination
                    BloodFilms = haematology.match_blood_film_examination(row[0], filtered_sentence, CurrentHaematologyTest)
                    if len(BloodFilms) > 0:
                        # Add all tests to HaematologyTests if not exists
                        print(BloodFilms)
                        for test in BloodFilms:
                            # trim test
                            test2 = test.strip()
                            CurrentHaematologyTest.append(test2)
                        continue
                    # Complete Blood Picture
                    CBC = haematology.match_complete_blood_picture(row[0], filtered_sentence, CurrentHaematologyTest)
                    if len(CBC) > 0:
                        # Add all tests to HaematologyTests if not exists
                        print(CBC)
                        for test in CBC:
                            # trim test
                            test2 = test.strip()
                            CurrentHaematologyTest.append(test2)
                        continue
                    # Typing
                    BTyping = haematology.match_blood_typing(row[0], filtered_sentence, CurrentHaematologyTest)
                    if len(BTyping) > 0:
                        # Add all tests to HaematologyTests if not exists
                        print(BTyping)
                        for test in BTyping:
                            # trim test
                            test2 = test.strip()
                            CurrentHaematologyTest.append(test2)
                        continue
                    # None matched
                    if len(CTest) == 0 and len(BloodFilms) == 0 and len(CBC) == 0 and len(BTyping) == 0:
                        # Add test to HaematologyTests if not exists
                        row[0] = row[0].strip()
                        CurrentHaematologyTest.append(row[0])
                        print([row[0]])
                    # Remove duplicates
                    CurrentHaematologyTest = list(dict.fromkeys(CurrentHaematologyTest))
            # Set haematology tests for CTC No.
            HaematologyTestsForCtcNo[CtcNo] = CurrentHaematologyTest
            # Append to HaematologyTests
            for test in CurrentHaematologyTest:
                if test not in HaematologyTests:
                    HaematologyTests.append(test)
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
    # print(HaematologyTests)

    # Export ServiceList into csv
    # ServiceList.to_csv('ServiceList.csv', index=False)

    # Render Test Info for each Ctc No
    print("Rendering Test Info for each Ctc No...")
    for Site, T in HaematologyTestsForCtcNo.items():
        # find test in database
        print(Site)
        TestGroup = []
        for test in T:
            SearchTest = HaematologyTestsDB.loc[HaematologyTestsDB['test'] == test]
            # if found
            if len(SearchTest) > 0:
                thisGroup = SearchTest.iloc[0, 3]
                MatchedTest = False
                # if thisGroup is not empty
                if thisGroup is not None:
                    for index, row in enumerate(TestGroup):
                        if row["TestGroup"] == thisGroup:
                            MatchedTest = True
                            row['Tests'].append(SearchTest.iloc[0])
                            break
                    if not MatchedTest:
                        MatchedTest = True
                        TestGroup.append({'TestGroup': thisGroup, 'Tests': [SearchTest.iloc[0]]})
                else:
                    TestGroup.append({'TestGroup': None, 'Tests': [SearchTest.iloc[0]]})
            else:
                TestGroup.append({'TestGroup': None, 'Tests': [SearchTest.iloc[0]]})
        # Print test group
        print(TestGroup)
        print("End")
        print("---------------------------------")
        # Encode to json
        # json_data = json.dumps(TestGroup, indent=4, default=str)
