from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize

import blood_tube
from haematology import HaematologyLab
from microbiology import MicrobiologyLab
from clinicalchemistry import ClinicalChemistryLab
import os
import pandas as pd
from joblib import load
from docx.shared import Pt, RGBColor


# Local Lab Tests
class LocalLab:
    chemistry = 0x01
    haematology = 0x02
    immunology = 0x03
    microbiology = 0x04
    pathology = 0x05


# Import LST excel file into pandas dataframe
def read_lst(path):
    df = pd.read_excel(path, header=0)
    return df


# Import all Services from budget's Excel file into pandas dataframe
def read_bpat_study_budget(path):
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


# Load budget interpretation of a single study
def get_budget_interpretation_by_ctcno(CtcNo):
    # Define root path
    BPATProjectFolder = '\\\\ctc-network.intranet\\dfs\\BPATCR\\Contract-Archive\\Project'
    CtcNoDigits = ''
    for i in range(len(CtcNo)):
        if CtcNo[i].isdigit():
            CtcNoDigits += CtcNo[i]
        else:
            break
    # Find study folder which is in level 2
    StudyPath = ''
    for SponsorFolder in os.listdir(BPATProjectFolder):
        SponsorPath = BPATProjectFolder + '\\' + SponsorFolder
        for StudyFolder in os.listdir(SponsorPath):
            if StudyFolder == CtcNoDigits:
                StudyPath = os.path.join(SponsorPath, StudyFolder)
                break
    if StudyPath == '':
        return False
    # Define budget file
    latest_budget_file = ''
    # Find budget files
    # Walk through all files in study folder
    for root, dirs, files in os.walk(StudyPath):
        # Loop through all files
        for file in files:
            # Only xlsx file
            if file.endswith('.xlsx') or file.endswith('.xls'):
                if latest_budget_file != '':
                    # Check the modification time of the file
                    if os.path.getmtime(os.path.join(root, file)) > os.path.getmtime(latest_budget_file):
                        # Update the value
                        latest_budget_file = os.path.join(root, file)
                else:
                    # Add new key
                    latest_budget_file = os.path.join(root, file)
    if latest_budget_file == '':
        return False
    # Get sheet names
    xl = pd.ExcelFile(latest_budget_file)
    sheet_names = xl.sheet_names
    # Read sheet with name 'Interpretation'
    if 'Interpretation' in sheet_names:
        df = pd.read_excel(latest_budget_file, 'Interpretation', header=0)
    else:
        return False
    return df


# Read folders to get a list of Excel xls files recursively into a dictionary
# Input: path to folder
# Output: dictionary of folder names and a list containing the path and the name of the file
def read_folders(path):
    folder_dict = {}
    for root, dirs, files in os.walk(path):
        # Open subfolders
        for dir in dirs:
            # Read files in sub-folders
            for root2, dirs2, files2 in os.walk(path + '\\' + dir):
                # Loop through files with index
                for index, file in enumerate(files2):
                    # Check if file is an Excel file
                    if file.endswith('.xls'):
                        # Check if key exists
                        if dir in folder_dict:
                            # Check the modification time of the file
                            if os.path.getmtime(root2 + '\\' + file) > os.path.getmtime(
                                    folder_dict[dir][0] + '\\' + folder_dict[dir][1]):
                                # Update the value
                                folder_dict[dir] = [root2, file]
                        else:
                            # Add new key
                            folder_dict[dir] = [root2, file]
    return folder_dict


# Determine whether Haematology is contained within a text string
# Use a pre-trained random forest model
# Input: text string
# Return Abbreviation of Local Lab
def parse_local_lab(item, description):
    # Import random forest model
    PathToModel = '.\\model.joblib'
    clf = load(PathToModel)
    # Import vectorizer
    PathToVectorizer = '.\\tfidf.joblib'
    vectorizer = load(PathToVectorizer)
    # Use the model
    # Vectorize the text
    text_vectorized = vectorizer.transform([item + ' ' + description])
    # Predict the text
    prediction = clf.predict(text_vectorized)
    if prediction == 'C':
        return LocalLab.chemistry
    elif prediction == 'H':
        return LocalLab.haematology
    elif prediction == 'M':
        return LocalLab.microbiology
    elif prediction == 'I':
        return LocalLab.immunology
    elif prediction == 'P':
        return LocalLab.pathology
    else:
        return False


# Parse all budget interpretation files
def parse_all_budget_interpretation():
    # define local service tracker information
    LSTPath = '\\\\ctc-network.intranet\\dfs\\BIOT\\01 Study Management\\02 Trackers\\Local Services Tracker.xlsm'
    LST = read_lst(LSTPath)
    # define budgets folder
    BPATProjectFolder = '\\\\ctc-network.intranet\\dfs\\BPATCR\\Contract-Archive\\Project\\Imago'
    # define local lab database path
    local_lab_db_path = "\\\\ctc-network.intranet\\dfs\\BIOT\\06 Laboratories and Site Services\\LocalLabTestsDB.xlsx"
    # Define unique test names for haematology
    HaematologyTests = []
    TestsForCtcNo = {}
    # Initialize haematology test object
    HaemaLab = HaematologyLab(LST, db_path=local_lab_db_path)
    # Define DataFrame to store budget information
    AllInterpretations = pd.DataFrame(columns=['CtcNo', 'Item', 'Description'])
    # Create pandas data frame to store budget information
    # AllInterpretations = pd.DataFrame(columns=['Service', 'Test Interpretation'])
    # Read the first level of BPATProjectFolder
    # for SponsorFolder in os.listdir(BPATProjectFolder):
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
            ServiceDf = read_bpat_study_budget(ServicePath)
            if ServiceDf is False:
                print("Service interpretation not found in Summary Services")
                continue
            # Check LST
            # Select row by key in LST
            site = LST.loc[LST['CTC No.'] == CtcNo]
            if site.empty:
                print('[' + CtcNo + '] Site not found in LST')
            TestsForCtcNo[CtcNo] = {"Haematology": []}
            # Loop through rows in Service
            # Flag for haematology test
            HasHaematology = False
            CurrentHaematologyTest = []
            # ThisService = pd.DataFrame({ "CtcNo" : CtcNo, "Item" : ServiceDf.iloc[:, 0], "Description" : ServiceDf.iloc[:, 1] })
            # AllInterpretations = pd.concat([AllInterpretations, ThisService], ignore_index=True)
            # print(len(AllInterpretations))
            for index, row in ServiceDf.iterrows():
                if parse_local_lab(row[1], row[0]) == LocalLab.haematology:
                    # Update flag
                    HasHaematology = True
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
                    CTest = HaemaLab.match_coagulation_tests(row[0], filtered_sentence, CurrentHaematologyTest)
                    # if CTest is not empty
                    if len(CTest) > 0:
                        # Add all tests to HaematologyTests if not exists
                        for test in CTest:
                            # trim test
                            test2 = test.strip()
                            CurrentHaematologyTest.append(test2)
                    # Blood Film Examination
                    BloodFilms = HaemaLab.match_blood_film_examination(row[0], filtered_sentence,
                                                                          CurrentHaematologyTest)
                    if len(BloodFilms) > 0:
                        # Add all tests to HaematologyTests if not exists
                        for test in BloodFilms:
                            # trim test
                            test2 = test.strip()
                            CurrentHaematologyTest.append(test2)
                        continue
                    # Complete Blood Picture
                    CBC = HaemaLab.match_complete_blood_picture(row[0], filtered_sentence, CurrentHaematologyTest)
                    if len(CBC) > 0:
                        # Add all tests to HaematologyTests if not exists
                        for test in CBC:
                            # trim test
                            test2 = test.strip()
                            CurrentHaematologyTest.append(test2)
                        continue
                    # Typing
                    BTyping = HaemaLab.match_blood_typing(row[0], filtered_sentence, CurrentHaematologyTest)
                    if len(BTyping) > 0:
                        # Add all tests to HaematologyTests if not exists
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
                    # Remove duplicates
                    CurrentHaematologyTest = list(dict.fromkeys(CurrentHaematologyTest))
            # Set haematology tests for CTC No.
            TestsForCtcNo[CtcNo]["Haematology"] = CurrentHaematologyTest
            # Append to HaematologyTests
            for test in CurrentHaematologyTest:
                if test not in HaematologyTests:
                    HaematologyTests.append(test)
            if HasHaematology is False:
                print("No haematology test found")
            else:
                print("Haematology test found in " + CtcNo)
                if not site.empty:
                    # Check whether column 8 is empty
                    if pd.isnull(site.iloc[0, 8]):
                        print('Haematology lab not set')

        else:
            print("File not found: " + ServicePath)

    # Export into excel
    # AllInterpretations.to_excel("AllInterpretations.xlsx", index=False)

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
    for Site, T in TestsForCtcNo.items():
        # find test in database
        print(Site)
        # Print test group
        print(HaemaLab.render_haematology_test_group(T["Haematology"]))
        print("End")
        print("---------------------------------")
        # Encode to json
        # json_data = json.dumps(TestGroup, indent=4, default=str)


# Parse a single study
def render_form_for_study(Study):
    # define local service tracker information
    LSTPath = '\\\\ctc-network.intranet\\dfs\\BIOT\\01 Study Management\\02 Trackers\\Local Services Tracker.xlsm'
    LST = read_lst(LSTPath)
    # define local lab database path
    local_lab_db_path = "\\\\ctc-network.intranet\\dfs\\BIOT\\06 Laboratories and Site Services\\LocalLabTestsDB.xlsx"
    chem_ec_db_path = "\\\\ctc-network.intranet\\dfs\\BIOT\\06 Laboratories and Site Services\\01 QMH Division of Chemical Pathology\\Biochem_ECPath_Methodology_TAT 120320_20221221.xlsx"
    rr_path = "\\\\ctc-network.intranet\\dfs\\BIOT\\06 Laboratories and Site Services\\03 QMH Department of Microbiology\\MB_RI_Other Tests_20230317.docx"
    # Initialize labs
    HaemaLab = HaematologyLab(lst=LST, db_path=local_lab_db_path)
    MicroLab = MicrobiologyLab(lst=LST, db_path=local_lab_db_path, rr_path=rr_path)
    ChemLab = ClinicalChemistryLab(lst=LST, db_path=chem_ec_db_path)
    # define export folder path
    UseExportPath = True
    ExportPaths = ['\\\\ctc-network.intranet\\dfs\\BIOTR\\01 Ongoing Studies\\', '\\\\ctc-network.intranet\\dfs\\BIOTR\\02 Closed Studies']
    # Extract digits of CTC No.
    CtcNoDigit = ''
    # Extract first digits
    for c in Study:
        if c.isdigit():
            CtcNoDigit += c
        else:
            break
    # Load budget file
    BudgetDf = get_budget_interpretation_by_ctcno(Study)
    print("Loading budget interpretation for " + Study + "...")
    # Open
    if BudgetDf is False:
        print('Budget interpretation not Found!')
        return
    CurrentHaematologyTest = []
    CurrentMicrobiologyTest = []
    CurrentChemistryTest = []
    for index, row in BudgetDf.iterrows():
        # Predict local lab
        predicted_lab = parse_local_lab(row[1], row[0])
        # Match haematology
        if predicted_lab == LocalLab.haematology:
            # Clean row[1]
            # Remove stop words
            stop_words = set(stopwords.words('english'))
            # Replace / with space
            row[1] = row[1].replace('/', ' ')
            # Tokenize
            word_tokens = word_tokenize(row[1])
            filtered_sentence = [w for w in word_tokens if not w in stop_words]
            # Remove punctuation
            filtered_sentence = [w for w in filtered_sentence if w.isalnum()]
            # To lower case
            filtered_sentence = [w.lower() for w in filtered_sentence]
            # Interpret tests
            matched_tests = HaemaLab.interpret_tests(row[0], filtered_sentence, CurrentHaematologyTest)
            # Append to CurrentHaematologyTest
            if len(matched_tests) > 0:
                for test in matched_tests:
                    # trim test
                    test2 = test.strip()
                    CurrentHaematologyTest.append(test2)
            # Remove duplicates
            CurrentHaematologyTest = list(dict.fromkeys(CurrentHaematologyTest))
        elif predicted_lab == LocalLab.microbiology:
            # Remove stop words
            stop_words = set(stopwords.words('english'))
            # Tokenize
            word_tokens = word_tokenize(row[1])
            filtered_sentence = [w for w in word_tokens if not w in stop_words]
            # Remove punctuation
            # filtered_sentence = [w for w in filtered_sentence if w.isalnum()]
            # To lower case
            filtered_sentence = [w.lower() for w in filtered_sentence]
            # Interpret tests
            matched_tests = MicroLab.interpret_tests(row[0], filtered_sentence, CurrentMicrobiologyTest)
            # Append to CurrentHaematologyTest
            if len(matched_tests) > 0:
                for test in matched_tests:
                    # trim test
                    test2 = test.strip()
                    CurrentMicrobiologyTest.append(test2)
            # Remove duplicates
            CurrentMicrobiologyTest = list(dict.fromkeys(CurrentMicrobiologyTest))
            continue
        elif predicted_lab == LocalLab.chemistry:
            # Remove stop words
            stop_words = set(stopwords.words('english'))
            # Tokenize
            word_tokens = word_tokenize(row[1])
            filtered_sentence = [w for w in word_tokens if not w in stop_words]
            # Remove punctuation
            filtered_sentence = [w for w in filtered_sentence if w.isalnum()]
            # To lower case
            filtered_sentence = [w.lower() for w in filtered_sentence]
            # Interpret tests
            matched_tests = ChemLab.interpret_tests(row[0], filtered_sentence, CurrentChemistryTest)
            # Append to CurrentChemistryTest
            if len(matched_tests) > 0:
                for test in matched_tests:
                    # trim test
                    test2 = test.strip()
                    CurrentChemistryTest.append(test2)
            # Remove duplicates
            CurrentChemistryTest = list(dict.fromkeys(CurrentChemistryTest))
            continue
    # Haematology section
    if len(CurrentHaematologyTest) == 0:
        print('No Haematology Tests Found!')
    else:
        print('Haematology Tests Found!')
        print("Rendering Haematology Test Form...")
        # Remove duplicates
        CurrentHaematologyTest = list(dict.fromkeys(CurrentHaematologyTest))
        if not HaemaLab.render_haematology_test_form(Study, HaemaLab.render_haematology_test_group(CurrentHaematologyTest)):
            print("Haematology Test Form Render Failed!")
    # Microbiology section
    if len(CurrentMicrobiologyTest) == 0:
        print('No Microbiology Test')
    else:
        print('Microbiology Test Found!')
        # Create a file name
        site = LST.loc[LST['CTC No.'] == Study]
        RRExportFileName = ''
        if (len(site) > 0):
            if not isinstance(site.iloc[0, 1], str):
                # Convert to string
                protocol = str(site.iloc[0, 1])
            else:
                protocol = site.iloc[0, 1]
            RRExportFileName = '[AutoGen] ' + site.iloc[0, 0] + '_' + site.iloc[0, 2] + '_' + protocol + '_MicroReferenceRanges.docx'
        else:
            RRExportFileName = '[AutoGen] ' + Study + '_MicroReferenceRanges.docx'
        if not UseExportPath:
            if MicroLab.render_reference_range_request(Study, CurrentMicrobiologyTest, RRExportFileName):
                print("Haematology Test Form Rendered: " + RRExportFileName)
        else:
            # find study folder in export path
            StudyFolder = ''
            for ExportPath in ExportPaths:
                for dirs in os.listdir(ExportPath):
                    if dirs.startswith(CtcNoDigit + '_'):
                        StudyFolder = os.path.join(ExportPath, dirs)
                        # Check sub folder
                        for dirs2 in os.listdir(StudyFolder):
                            if dirs2.startswith('03 '):
                                StudyFolder = os.path.join(StudyFolder, dirs2)
                                break
                        break
                    if StudyFolder != '':
                        break
                if StudyFolder != '':
                    break
            if StudyFolder == '':
                print('Error: Study folder not found')
                print('Export to default path')
                if MicroLab.render_reference_range_request(Study, CurrentMicrobiologyTest, os.path.join(StudyFolder, RRExportFileName)):
                    print('Microbiology Reference Range File Created: ' + RRExportFileName)
            else:
                if MicroLab.render_reference_range_request(Study, CurrentMicrobiologyTest, os.path.join(StudyFolder, RRExportFileName)):
                    print('Microbiology Reference Range File Created: ' + os.path.join(StudyFolder, RRExportFileName))
        print("Rendering Microbiology Test Form...")
        micbio_test_group = MicroLab.render_test_group(CurrentMicrobiologyTest)
        MicroLab.render_test_form(Study, micbio_test_group)
    # Chemistry section
    if len(CurrentChemistryTest) == 0:
        print('No Chemistry Test')
    else:
        print('Chemistry Test Found!')
        if not ChemLab.render_chemistry_test_form(Study, ChemLab.render_test_group(CurrentChemistryTest)):
            print('Error: Chemistry Test Form Render Failed')
    print('Finished')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    render_form_for_study("2355HKU1")
