import csv

from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
import haematology
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


# Import all Summary Services excel file into pandas dataframe
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


# Load summary of service file from a single study
def get_summary_service_by_ctcno(CtcNo):
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
    # Define summary of service file
    SummaryServiceFile = ''
    # Find summary service file
    # Walk through all files in study folder
    for root, dirs, files in os.walk(StudyPath):
        # Loop through all files
        for file in files:
            # Only xlsx file
            if file.endswith('.xlsx') or file.endswith('.xls'):
                if SummaryServiceFile != '':
                    # Check the modification time of the file
                    if os.path.getmtime(os.path.join(root, file)) > os.path.getmtime(SummaryServiceFile):
                        # Update the value
                        SummaryServiceFile = os.path.join(root, file)
                else:
                    # Add new key
                    SummaryServiceFile = os.path.join(root, file)
    if SummaryServiceFile == '':
        return False
    # Get sheet names
    xl = pd.ExcelFile(SummaryServiceFile)
    sheet_names = xl.sheet_names
    # Read sheet with name 'Interpretation'
    if 'Interpretation' in sheet_names:
        df = pd.read_excel(SummaryServiceFile, 'Interpretation', header=0)
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
# Return True if it is, False if it is not
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


# Parse all summary of services
def parse_all_summary_services():
    # define local service tracker information
    LSTPath = '\\\\ctc-network.intranet\\dfs\\BIOT\\01 Study Management\\02 Trackers\\Local Services Tracker.xlsm'
    LST = read_lst(LSTPath)
    # define summary services information
    BPATProjectFolder = '\\\\ctc-network.intranet\\dfs\\BPATCR\\Contract-Archive\\Project\\Imago'
    # Define unique test names for haematology
    HaematologyTests = []
    TestsForCtcNo = {}
    # Define DataFrame to store service information
    AllServices = pd.DataFrame(columns=['CtcNo', 'Item', 'Description'])
    # Create pandas data frame to store service information
    # ServiceList = pd.DataFrame(columns=['Service', 'Test Interpretation'])
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
            TestsForCtcNo[CtcNo] = {"Haematology": []}
            # Loop through rows in Service
            # Flag for haematology test
            HasHaematology = False
            CurrentHaematologyTest = []
            # ThisService = pd.DataFrame({ "CtcNo" : CtcNo, "Item" : ServiceDf.iloc[:, 0], "Description" : ServiceDf.iloc[:, 1] })
            # AllServices = pd.concat([AllServices, ThisService], ignore_index=True)
            # print(len(AllServices))
            for index, row in ServiceDf.iterrows():
                if parse_local_lab(row[1], row[0]) == LocalLab.haematology:
                    # Update flag
                    HasHaematology = True
                    # Print item name
                    # print(row[0])
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
                        # print(CTest)
                        # Add all tests to HaematologyTests if not exists
                        for test in CTest:
                            # trim test
                            test2 = test.strip()
                            CurrentHaematologyTest.append(test2)
                    # Blood Film Examination
                    BloodFilms = haematology.match_blood_film_examination(row[0], filtered_sentence,
                                                                          CurrentHaematologyTest)
                    if len(BloodFilms) > 0:
                        # Add all tests to HaematologyTests if not exists
                        # print(BloodFilms)
                        for test in BloodFilms:
                            # trim test
                            test2 = test.strip()
                            CurrentHaematologyTest.append(test2)
                        continue
                    # Complete Blood Picture
                    CBC = haematology.match_complete_blood_picture(row[0], filtered_sentence, CurrentHaematologyTest)
                    if len(CBC) > 0:
                        # Add all tests to HaematologyTests if not exists
                        # (CBC)
                        for test in CBC:
                            # trim test
                            test2 = test.strip()
                            CurrentHaematologyTest.append(test2)
                        continue
                    # Typing
                    BTyping = haematology.match_blood_typing(row[0], filtered_sentence, CurrentHaematologyTest)
                    if len(BTyping) > 0:
                        # Add all tests to HaematologyTests if not exists
                        # print(BTyping)
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
                        # print([row[0]])
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
    # AllServices.to_excel("AllServices.xlsx", index=False)

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
        print(haematology.render_haematology_test_group(T["Haematology"]))
        print("End")
        print("---------------------------------")
        # Encode to json
        # json_data = json.dumps(TestGroup, indent=4, default=str)


# Parse a single study
def render_form_for_study(Study):
    # define local service tracker information
    LSTPath = '\\\\ctc-network.intranet\\dfs\\BIOT\\01 Study Management\\02 Trackers\\Local Services Tracker.xlsm'
    LST = read_lst(LSTPath)
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
    # Load summary of service file
    ServiceDf = get_summary_service_by_ctcno(Study)
    print("Loading summary of service for " + Study + "...")
    # Open
    if ServiceDf is False:
        print('Summary of Service not Found!')
        return
    CurrentHaematologyTest = []
    for index, row in ServiceDf.iterrows():
        # Match haematology
        if parse_local_lab(row[1], row[0]) == LocalLab.haematology:
            # Update flag
            # Print item name
            # print(row[0])
            # Clean row[1] using nltk
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
            # Match Coagulation tests
            CTest = haematology.match_coagulation_tests(row[0], filtered_sentence, CurrentHaematologyTest)
            # if CTest is not empty
            if len(CTest) > 0:
                # print(CTest)
                # Add all tests to HaematologyTests if not exists
                for test in CTest:
                    # trim test
                    test2 = test.strip()
                    CurrentHaematologyTest.append(test2)
            # Blood Film Examination
            BloodFilms = haematology.match_blood_film_examination(row[0], filtered_sentence, CurrentHaematologyTest)
            if len(BloodFilms) > 0:
                # Add all tests to HaematologyTests if not exists
                # print(BloodFilms)
                for test in BloodFilms:
                    # trim test
                    test2 = test.strip()
                    CurrentHaematologyTest.append(test2)
                continue
            # Complete Blood Picture
            CBC = haematology.match_complete_blood_picture(row[0], filtered_sentence, CurrentHaematologyTest)
            if len(CBC) > 0:
                # Add all tests to HaematologyTests if not exists
                # (CBC)
                for test in CBC:
                    # trim test
                    test2 = test.strip()
                    CurrentHaematologyTest.append(test2)
                continue
            # Typing
            BTyping = haematology.match_blood_typing(row[0], filtered_sentence, CurrentHaematologyTest)
            if len(BTyping) > 0:
                # Add all tests to HaematologyTests if not exists
                # print(BTyping)
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
                # print([row[0]])
            # Remove duplicates
            CurrentHaematologyTest = list(dict.fromkeys(CurrentHaematologyTest))
    # Render Form
    if len(CurrentHaematologyTest) == 0:
        print('No Haematology Tests Found!')
        return
    else:
        print('Haematology Tests Found!')
        # Remove duplicates
        CurrentHaematologyTest = list(dict.fromkeys(CurrentHaematologyTest))
        TestGroup = haematology.render_haematology_test_group(CurrentHaematologyTest)
        # Set up word document
        FormPath = '.\\FormTemplateHaemaLab.docx'
        FormDoc = haematology.OpenHaemaFormTemplate(FormPath)
        # Fill study info
        if len(FormDoc.tables[0].rows[2].cells[1].paragraphs) > 1:
            SiteCell2 = FormDoc.tables[0].rows[2].cells[1].paragraphs[1]
        else:
            SiteCell2 = FormDoc.tables[0].rows[2].cells[1].add_paragraph()
        SiteCell2.add_run('CTC' + CtcNoDigit).bold = True
        # Select row by key in LST
        site = LST.loc[LST['CTC No.'] == Study]
        if len(site) > 0:
            if isinstance(site.iloc[0, 30], str):
                # Para0 = FormDoc.tables[0].rows[3].cells[1].paragraphs[0]
                # Para0.add_run('Rept Locn').font.size = Pt(9)
                if len(FormDoc.tables[0].rows[3].cells[1].paragraphs) > 1:
                    Para1 = FormDoc.tables[0].rows[3].cells[1].paragraphs[1]
                    Para1Run = Para1.add_run(site.iloc[0, 30])
                    Para1Run.bold = True
                    Para1Run.font.size = Pt(12)
                else:
                    Para1 = FormDoc.tables[0].rows[3].cells[1].add_paragraph()
                    Para1Run = Para1.add_run(site.iloc[0, 30])
                    Para1Run.bold = True
                    Para1Run.font.size = Pt(12)
            if isinstance(site.iloc[0, 1], str):
                FormDoc.tables[1].rows[0].cells[0].text = ''
                Prot1Run = FormDoc.tables[1].rows[0].cells[0].paragraphs[0].add_run('Protocol: ' + site.iloc[0, 1])
                Prot1Run.font.size = Pt(10)
                Prot1Run.bold = True
                FormDoc.tables[1].rows[0].cells[0].paragraphs[0].paragraph_format.space_before = Pt(6)
            if isinstance(site.iloc[0, 28], str):
                FormDoc.tables[1].rows[3].cells[0].text = 'Contact Person: ' + site.iloc[0, 28]
                FormDoc.tables[1].rows[3].cells[0].paragraphs[0].runs[0].font.size = Pt(10)
            if isinstance(site.iloc[0, 29], str):
                FormDoc.tables[1].rows[3].cells[1].text = 'Contact Number: ' + site.iloc[0, 29]
                FormDoc.tables[1].rows[3].cells[1].paragraphs[0].runs[0].font.size = Pt(10)
        # Loop through tests
        print("Rendering Haematology Test Form...")
        for TG in TestGroup:
            # Print all test group
            # print(TG)
            # Content
            row1 = FormDoc.tables[2].add_row()
            row1.cells[0].text = u'\u25a1'
            # Apply style
            for paragraph in row1.cells[0].paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'PMingLiU'
            # Loop through test with index
            Content = ''
            CollectionTubes = []
            OptionalTest = []
            # Check at least one test is not optional
            HasNonOptional = False
            for test in TG['Tests']:
                if not test['is_optional']:
                    HasNonOptional = True
            # Loop through test with index
            RowFilled = False
            for index, test in enumerate(TG['Tests']):
                if HasNonOptional and test['is_optional'] is True:
                    OptionalTest.append(test)
                    continue
                if RowFilled:
                    Content += '\n'
                Content += test['test']
                # Add specimen to collection tubes if not exists
                if isinstance(test['specimen'], str) and test['specimen'] != '':
                    if test['specimen'] not in CollectionTubes:
                        CollectionTubes.append(test['specimen'])
                RowFilled = True
            row1.cells[1].text = Content
            # Add collection tubes paragraph
            if len(CollectionTubes) > 0:
                # Loop through collection tubes
                for i, tube in enumerate(CollectionTubes):
                    if isinstance(tube, str) and tube != '':
                        row1.cells[1].add_paragraph('[' + tube + ']')
                    # row1.cells[1].add_paragraph('[' + tube + ']').runs[0].font.color.rgb = RGBColor(0x00, 0xAA, 0x00)
            # Merge cell 1 with 2
            row1.cells[1].merge(row1.cells[2])
            # Move row
            rowA = FormDoc.tables[2].rows[len(FormDoc.tables[2].rows) - 1]
            rowB = FormDoc.tables[2].rows[len(FormDoc.tables[2].rows) - 2]
            rowA._tr.addnext(rowB._tr)
            # Add optional test
            if len(OptionalTest) > 0:
                for i2, t2 in enumerate(OptionalTest):
                    row2 = FormDoc.tables[2].add_row()
                    row2.cells[1].text = u'\u25a1'
                    # Apply style
                    for paragraph in row2.cells[1].paragraphs:
                        for run in paragraph.runs:
                            run.font.name = 'PMingLiU'
                    row2.cells[2].text = t2['test']
                    # Move row
                    rowA = FormDoc.tables[2].rows[len(FormDoc.tables[2].rows) - 1]
                    rowB = FormDoc.tables[2].rows[len(FormDoc.tables[2].rows) - 2]
                    rowA._tr.addnext(rowB._tr)
                # Merge row
                Cell1 = FormDoc.tables[2].rows[len(FormDoc.tables[2].rows) - 2].cells[0]
                Cell2 = FormDoc.tables[2].rows[len(FormDoc.tables[2].rows) - 3].cells[0]
                Cell1.merge(Cell2)
        # Remove placeholder row 2
        Table = FormDoc.tables[2]._tbl
        RemoveRow = FormDoc.tables[2].rows[2]._tr
        Table.remove(RemoveRow)
        try:
            # Create a file name
            site = LST.loc[LST['CTC No.'] == Study]
            ExportFileName = ''
            if (len(site) > 0):
                ExportFileName = '[AutoGen] ' + site.iloc[0,0] + '_' + site.iloc[0,2] + '_' + site.iloc[0,1] +'_HemaForm.docx'
            else:
                ExportFileName = '[AutoGen] ' + Study + ' HemaForm.docx'
            if not UseExportPath:
                FormDoc.save(ExportFileName)
                print("Haematology Test Form Rendered: " + ExportFileName)
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
                    FormDoc.save(ExportFileName)
                    return
                else:
                    FormDoc.save(os.path.join(StudyFolder, ExportFileName))
                    print("Haematology Test Form Rendered: " + StudyFolder + "\\" + ExportFileName)
                    return
        except:
            print('Error: File is open')
            return
        print('OK')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    render_form_for_study("2305HKU1")
