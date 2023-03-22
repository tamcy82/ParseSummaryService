# Clinical Chemistry Module

# Path: clinicalchemistry.py
import pandas as pd
from docx import Document
from nltk.tokenize import word_tokenize
from fuzzywuzzy import fuzz
import re


# Define Clinical Chemistry Class
class ClinicalChemistryLab:

    # Constructor
    def __init__(self, lst, db_path='.\\LocalLabTestsDB.xlsx'):
        # Define paths
        self.study_paths = ['\\\\ctc-network.intranet\\dfs\\BIOTR\\01 Ongoing Studies\\',
                            '\\\\ctc-network.intranet\\dfs\\BIOTR\\02 Closed Studies']
        # Define local service tracker data
        self.lst = lst
        # Define Clinical Chemistry tests
        self.chem_test = []
        # Load ClinicalChemistryTestsDB
        self.ClinicalChemistryTestsDBPath = db_path
        self.ClinicalChemistryTestsDB = self.load_ec_path_code(self.ClinicalChemistryTestsDBPath)
        # Select only clinical chemistry tests
        self.ClinicalChemistryTestsDB = self.ClinicalChemistryTestsDB.loc[self.ClinicalChemistryTestsDB['lab'] == 'chem']
        # Convert all_name to string
        self.ClinicalChemistryTestsDB['alt_name'] = self.ClinicalChemistryTestsDB['alt_name'].astype(str)
        # Expand alt_name column
        self.ClinicalChemistryTestsDB['alt_name'] = self.ClinicalChemistryTestsDB['alt_name'].str.split(',')
        # Define form template path
        self.ClinicalChemistryFormTemplatePath = '.\\templates\\ClinicalChemistryFormTemplate.docx'

    # Load ClinicalChemistryTestsDB into pandas
    def load_chemistry_tests_db(self, path):
        df = pd.read_excel(path, header=0)
        return df

    # Load ECPathCode into pandas
    def load_ec_path_code(self, file_path):
        df = pd.read_excel(file_path, header=0, dtype=str).fillna("")
        # Parse file
        # Create a new dataframe
        df_new = pd.DataFrame(columns=['lab', 'test', 'alt_name', 'specimen', 'code', 'section_order', 'is_optional',
                                       'remarks'])
        # Iterate through each row
        for index, row in df.iterrows():
            # Get lab
            lab = 'chem'
            # Get test
            test = row['Test']
            # Get alt_name
            # alt_name = row['Alt Name']
            alt_name = []
            # Get specimen
            specimen = row['Specimen']
            # Get code
            code = row['ECPath Code']
            # Get section_order
            section_order = row['Section']
            # Get is_optional
            is_optional = ""
            # Get remarks
            remarks = row['Remark on Form'] + " " + row['Internal Remark']
            # Append to new dataframe
            df_new = df_new.append({'lab': lab, 'test': test, 'alt_name': alt_name, 'specimen': specimen,
                                    'code': code, 'section_order': section_order, 'is_optional': is_optional,
                                    'remarks': remarks}, ignore_index=True)
        return df_new

    # Open a given clinical chemistry form template
    def open_chem_form_template(self, path):
        document = Document(path)
        return document

    # Interpret tests
    def interpret_tests(self, item, description, all_tests):
        matched_tests = []
        # Match pregnancy tests
        preg_tests = self.match_pregnancy_tests(item, description, all_tests)
        if len(preg_tests) > 0:
            matched_tests.extend(preg_tests)
        # Match glucose tests
        glucose_tests = self.match_glucose_tests(item, description, all_tests)
        if len(glucose_tests) > 0:
            matched_tests.extend(glucose_tests)
        # Match bilirubin tests
        bilirubin_tests = self.match_bilirubin_tests(item, description, all_tests)
        if len(bilirubin_tests) > 0:
            matched_tests.extend(bilirubin_tests)
        # Match calcium tests
        calcium_tests = self.match_calcium_tests(item, description, all_tests)
        if len(calcium_tests) > 0:
            matched_tests.extend(calcium_tests)
        # Match remaining tests
        remaining_tests = self.match_tests_from_db(item, description, all_tests)
        if len(remaining_tests) > 0:
            matched_tests.extend(remaining_tests)
        # If none of the above tests are matched
        if len(matched_tests) == 0:
            matched_tests.append(item)
        return matched_tests

    # Match proceeding and postceeding words
    def get_proceeding_postceeding_words(self, word_index, texts):
        if word_index < len(texts) - 1:
            # Get proceeding and postceeding words
            if word_index > 0:
                proceeding_word = texts[word_index - 1]
            else:
                proceeding_word = ''
            if word_index < len(texts) - 1:
                postceeding_word = texts[word_index + 1]
            else:
                postceeding_word = ''
            return proceeding_word, postceeding_word
        return None

    # match pregnancy tests
    # To detect incoming item and description that may contain both urine and serum tests
    def match_pregnancy_tests(self, item, description, all_tests):
        # Define matched_tests container
        matched_tests = []
        # Match basic name
        if "pregnancy" in item.lower() or "pregnancy" in description:
            # Pregnancy tests found
            # Check if urine pregnancy test is requested
            if "urine" in item.lower() or "urine" in description:
                # Urine pregnancy test found
                # Check if urine pregnancy test is not already in all_tests
                if "Urine Pregnancy Test" not in all_tests:
                    # Urine pregnancy test not found in all_tests
                    # Add urine pregnancy test to matched_tests
                    matched_tests.append("Urine Pregnancy Test")
            # Check if serum pregnancy test is requested
            if "serum" in item.lower() or "serum" in description:
                # Serum pregnancy test found
                # Check if serum pregnancy test is not already in all_tests
                if "Serum Pregnancy Test" not in all_tests:
                    # Serum pregnancy test not found in all_tests
                    # Add serum pregnancy test to matched_tests
                    matched_tests.append("Serum Pregnancy Test")
        return matched_tests

    # match bilirubin tests
    # To detect incoming item and description that may contain both total and direct bilirubin tests
    # Check preceeding and postceeding words to detect
    # Return matched_tests
    def match_bilirubin_tests(self, item, description, all_tests):
        # Define matched_tests container
        matched_tests = []
        # Define flags
        is_direct_bilirubin = False
        is_indirect_bilirubin = False
        is_total_bilirubin = False
        # Tokenize item
        item_tokens = word_tokenize(item)
        # Remove punctuation
        item_tokens = [token for token in item_tokens if token.isalnum()]
        # To lower
        item_tokens = [token.lower() for token in item_tokens]
        # Simple ECPath code detection first
        if "dbil" in item_tokens or "dbil" in description:
            matched_tests.append("Direct Bilirubin")
            is_direct_bilirubin = True
        if "tbil" in item_tokens or "tbil" in description:
            matched_tests.append("Total Bilirubin")
            is_total_bilirubin = True
        # Match basic name
        if "bilirubin" in item:
            # Bilirubin tests found
            # Loop through test names with index
            for bilirubin_item_index, t in enumerate(item_tokens):
                if t == "bilirubin":
                    # Get proceeding and postceeding words
                    proceeding_word, postceeding_word = self.get_proceeding_postceeding_words(bilirubin_item_index, item_tokens)
                    # Check if total bilirubin is requested
                    if not is_total_bilirubin and (postceeding_word == "total" or postceeding_word == "total"):
                        # Total bilirubin found
                        # Check if total bilirubin is not already in all_tests
                        if "Total Bilirubin" not in all_tests:
                            # Total bilirubin not found in all_tests
                            # Add total bilirubin to matched_tests
                            matched_tests.append("Total Bilirubin")
                            is_total_bilirubin = True
                    # Check if in-direct bilirubin is requested
                    if not is_indirect_bilirubin and (proceeding_word == "in-direct" or postceeding_word == "in-direct"):
                        # Indirect bilirubin found
                        # Check if indirect bilirubin is not already in all_tests
                        if "In-direct Bilirubin" not in all_tests:
                            # Indirect bilirubin not found in all_tests
                            # Add indirect bilirubin to matched_tests
                            matched_tests.append("In-direct Bilirubin")
                            is_indirect_bilirubin = True
                    # Check if direct bilirubin is requested
                    if not is_direct_bilirubin and (proceeding_word == "direct" or postceeding_word == "direct"):
                        # Direct bilirubin found
                        # Check if direct bilirubin is not already in all_tests
                        if "Direct Bilirubin" not in all_tests:
                            # Direct bilirubin not found in all_tests
                            # Add direct bilirubin to matched_tests
                            matched_tests.append("Direct Bilirubin")
                            is_direct_bilirubin = True
        if "bilirubin" in description:
            # Bilirubin tests found
            # Loop through test names with index
            for bilirubin_description_index, t in enumerate(description):
                if t == "bilirubin":
                    # Get proceeding and postceeding words
                    proceeding_word, postceeding_word = self.get_proceeding_postceeding_words(bilirubin_description_index, description)
                    # Check if total bilirubin is requested
                    if not is_total_bilirubin and (postceeding_word == "total" or postceeding_word == "total"):
                        # Total bilirubin found
                        # Check if total bilirubin is not already in all_tests
                        if "Total Bilirubin" not in all_tests:
                            # Total bilirubin not found in all_tests
                            # Add total bilirubin to matched_tests
                            matched_tests.append("Total Bilirubin")
                            is_total_bilirubin = True
                    # Check if in-direct bilirubin is requested
                    if not is_indirect_bilirubin and (proceeding_word == "in-direct" or postceeding_word == "in-direct"):
                        # Indirect bilirubin found
                        # Check if indirect bilirubin is not already in all_tests
                        if "In-direct Bilirubin" not in all_tests:
                            # Indirect bilirubin not found in all_tests
                            # Add indirect bilirubin to matched_tests
                            matched_tests.append("In-direct Bilirubin")
                            is_indirect_bilirubin = True
                    # Check if direct bilirubin is requested
                    if not is_direct_bilirubin and (proceeding_word == "direct" or postceeding_word == "direct"):
                        # Direct bilirubin found
                        # Check if direct bilirubin is not already in all_tests
                        if "Direct Bilirubin" not in all_tests:
                            # Direct bilirubin not found in all_tests
                            # Add direct bilirubin to matched_tests
                            matched_tests.append("Direct Bilirubin")
                            is_direct_bilirubin = True
        # If nothing is found
        if len(matched_tests) == 0 and ("bilirubin" in item or "bilirubin" in description):
            # Add total bilirubin to matched_tests
            matched_tests.append("Bilirubin (UNKNOWN TYPE)")
        return matched_tests

    # match different types of glucose tests
    # To detect incoming item and description that may contain different types of glucose tests
    # Check preceeding and postceeding words to detect fasting and random glucose tests
    # Return matched_tests
    def match_glucose_tests(self, item, description, all_tests):
        # Define matched_tests container
        matched_tests = []
        # Define flags
        fasting_glucose_flag = False
        random_glucose_flag = False
        csf_glucose_flag = False
        # Tokenize item
        item_tokens = word_tokenize(item)
        # Remove punctuation
        item_tokens = [token for token in item_tokens if token.isalnum()]
        # To lower
        item_tokens = [token.lower() for token in item_tokens]
        # Simple ECPath code detection first
        if "fg" in item or "fg" in description:
            matched_tests.append("Glucose (Fasting)")
            fasting_glucose_flag = True
        if "rg" in item or "rg" in description:
            matched_tests.append("Glucose (Random)")
            random_glucose_flag = True
        if "cglu" in item or "cglu" in description:
            matched_tests.append("CSF Glucose")
            csf_glucose_flag = True
        # Match name in item
        if "glucose" in item:
            # Glucose tests found
            # Loop through test names with index
            for glucose_item_index, t in enumerate(item_tokens):
                if t == "glucose":
                    if glucose_item_index > 0:
                        # Get proceeding and postceeding words
                        proceeding_word, postceeding_word = self.get_proceeding_postceeding_words(glucose_item_index, item_tokens)
                        # Check if fasting glucose is requested
                        if not fasting_glucose_flag and (proceeding_word == "fasting" or postceeding_word == "fasting"):
                            # Fasting glucose found
                            # Check if fasting glucose is not already in all_tests
                            if "Glucose (Fasting)" not in all_tests:
                                # Fasting glucose not found in all_tests
                                # Add fasting glucose to matched_tests
                                matched_tests.append("Glucose (Fasting)")
                                fasting_glucose_flag = True
                        # Check if random glucose is requested
                        if not random_glucose_flag and (proceeding_word == "random" or postceeding_word == "random"):
                            # Random glucose found
                            # Check if random glucose is not already in all_tests
                            if "Glucose (Random)" not in all_tests:
                                # Random glucose not found in all_tests
                                # Add random glucose to matched_tests
                                matched_tests.append("Glucose (Random)")
                                random_glucose_flag = True
                        # Check if csf glucose is requested
                        if not csf_glucose_flag and (proceeding_word == "csf" or postceeding_word == "csf"):
                            # CSF glucose found
                            # Check if csf glucose is not already in all_tests
                            if "CSF Glucose" not in all_tests:
                                # CSF glucose not found in all_tests
                                # Add csf glucose to matched_tests
                                matched_tests.append("CSF Glucose")
                                csf_glucose_flag = True
        # Match name in description
        if "glucose" in description:
            # Glucose tests found
            # Loop through test names with index
            for glucose_description_index, t in enumerate(description):
                if t == "glucose":
                    if glucose_description_index > 0:
                        # Get proceeding and postceeding words
                        proceeding_word, postceeding_word = self.get_proceeding_postceeding_words(glucose_description_index, description)
                        # Check if fasting glucose is requested
                        if not fasting_glucose_flag and (proceeding_word == "fasting" or postceeding_word == "fasting"):
                            # Fasting glucose found
                            # Check if fasting glucose is not already in all_tests
                            if "Glucose (Fasting)" not in all_tests:
                                # Fasting glucose not found in all_tests
                                # Add fasting glucose to matched_tests
                                matched_tests.append("Glucose (Fasting)")
                                fasting_glucose_flag = True
                        # Check if random glucose is requested
                        if not random_glucose_flag and (proceeding_word == "random" or postceeding_word == "random"):
                            # Random glucose found
                            # Check if random glucose is not already in all_tests
                            if "Glucose (Random)" not in all_tests:
                                # Random glucose not found in all_tests
                                # Add random glucose to matched_tests
                                matched_tests.append("Glucose (Random)")
                                random_glucose_flag = True
                        # Check if csf glucose is requested
                        if not csf_glucose_flag and (proceeding_word == "csf" or postceeding_word == "csf"):
                            # CSF glucose found
                            # Check if csf glucose is not already in all_tests
                            if "CSF Glucose" not in all_tests:
                                # CSF glucose not found in all_tests
                                # Add csf glucose to matched_tests
                                matched_tests.append("CSF Glucose")
                                csf_glucose_flag = True
        # If nothing is found
        if len(matched_tests) == 0\
                and ("glucose" in item or "glucose" in description):
            # Add glucose unknown type to matched_tests
            matched_tests.append("Glucose (UNKNOWN TYPE)")
        return matched_tests

    # Match different types of calcium tests
    # To detect incoming item and description that may contain different types of calcium tests
    # Check preceeding and postceeding words to detect ionized calcium tests
    # Return matched_tests
    def match_calcium_tests(self, item, description, all_tests):
        # Define matched_tests container
        matched_tests = []
        # Define flags
        ionized_calcium_flag = False
        # Tokenize item
        item_tokens = word_tokenize(item)
        # Remove punctuation
        item_tokens = [token for token in item_tokens if token.isalnum()]
        # To lower
        item_tokens = [token.lower() for token in item_tokens]
        # Simple ECPath code detection first
        if "ca" in item or "ca" in description:
            matched_tests.append("Calcium")
        # Match name in item
        if "calcium" in item:
            # Calcium tests found
            # Loop through test names with index
            for calcium_item_index, t in enumerate(item_tokens):
                if t == "calcium":
                    if calcium_item_index > 0:
                        # Get proceeding and postceeding words
                        proceeding_word, postceeding_word = self.get_proceeding_postceeding_words(calcium_item_index, item_tokens)
                        # Check if ionized calcium is requested
                        if not ionized_calcium_flag and (proceeding_word == "ionized" or postceeding_word == "ionized"):
                            # Ionized calcium found
                            # Check if ionized calcium is not already in all_tests
                            if "Ionized Calcium" not in all_tests:
                                # Ionized calcium not found in all_tests
                                # Add ionized calcium to matched_tests
                                matched_tests.append("Calcium (Ionized)")
                                ionized_calcium_flag = True
                        else:
                            # Ionized calcium not found
                            # Check if calcium is not already in all_tests
                            if "Calcium" not in all_tests:
                                # Calcium not found in all_tests
                                # Add calcium to matched_tests
                                matched_tests.append("Calcium")
        # Match name in description
        if "calcium" in description:
            # Calcium tests found
            # Loop through test names with index
            for calcium_description_index, t in enumerate(description):
                if t == "calcium":
                    if calcium_description_index > 0:
                        # Get proceeding and postceeding words
                        proceeding_word, postceeding_word = self.get_proceeding_postceeding_words(calcium_description_index, description)
                        # Check if ionized calcium is requested
                        if not ionized_calcium_flag and (proceeding_word == "ionized" or postceeding_word == "ionized"):
                            # Ionized calcium found
                            # Check if ionized calcium is not already in all_tests
                            if "Ionized Calcium" not in all_tests:
                                # Ionized calcium not found in all_tests
                                # Add ionized calcium to matched_tests
                                matched_tests.append("Calcium (Ionized)")
                                ionized_calcium_flag = True
                        else:
                            # Ionized calcium not found
                            # Check if calcium is not already in all_tests
                            if "Calcium" not in all_tests:
                                # Calcium not found in all_tests
                                # Add calcium to matched_tests
                                matched_tests.append("Calcium")
        # If nothing is found
        if len(matched_tests) == 0\
                and ("calcium" in item or "calcium" in description):
            # Add calcium unknown type to matched_tests
            matched_tests.append("Calcium (UNKNOWN TYPE)")
        return matched_tests

    # Match tests from DB
    # To detect incoming item and description that may contain tests from DB
    # Return matched_tests
    def match_tests_from_db(self, item, description, all_tests):
        # Define matched_tests container
        matched_tests = []
        # Tokenize item
        item_tokens = word_tokenize(item)
        # Remove punctuation
        item_tokens = [token for token in item_tokens if token.isalnum()]
        # To lower
        item_tokens = [token.lower() for token in item_tokens]
        # Define flags
        # Loop through tests in ClinicalChemistryTestsDB
        for index, test in self.ClinicalChemistryTestsDB.iterrows():
            # Skip special groups
            if test['test'].startswith("["):
                continue
            # This matching sequence is important
            # Match test name
            this_name = test['test'].lower()
            if item.startswith(this_name) or this_name in description:
                # Check if test is not already in all_tests
                if test not in all_tests:
                    # Test not found in all_tests
                    # Add test to matched_tests
                    matched_tests.append(test['test'])
                    continue
            # Match alternative names
            if len(test['alt_name']) > 0:
                for alt_name in test['alt_name']:
                    if alt_name.lower() in item.lower() or alt_name.lower() in description:
                        matched_tests.append(test['test'])
                        continue
            # if test['test'] has space
            if " " in test['test']:
                # Split T into words
                words = test['test'].split(" ")
                # Compare T with description sequentially
                for i in range(len(words)):
                    # If T is not in description, break
                    if words[i].lower() not in description:
                        break
                    # If T is in description, and it is the last word of T
                    if words[i].lower() in description and i == len(words) - 1:
                        matched_tests.append(test['test'])
            if this_name in item.lower():
                # Check if test is not already in all_tests
                if test not in all_tests:
                    # Test not found in all_tests
                    # Add test to matched_tests
                    matched_tests.append(test['test'])
                    continue
        return matched_tests

    # Render microbiology lab test groups
    def render_test_group(self, T):
        TestGroup = []
        for test in T:
            # Use multiple method to find test in HaematologyTestsDB
            # Exact match
            # Search for test in HaematologyTestsDB
            search_test = self.ClinicalChemistryTestsDB.loc[self.ClinicalChemistryTestsDB['test'] == test]
            # Check alt_name column if not found
            if len(search_test) == 0:
                # Loop through alt_name column
                for index, row in self.ClinicalChemistryTestsDB.iterrows():
                    # Check if test in alt_name
                    if test in row['alt_name']:
                        search_test = self.ClinicalChemistryTestsDB.loc[index]
                        break
            # If not found, clean the test name and search again
            if len(search_test) == 0:
                # Extract alphanumeric and '-' characters from search_test with regex
                test_clean = re.search(r'[\w\- ]+', test).group(0)
                # Trim trailing and leading whitespace
                test_clean = test_clean.strip()
                # Redo search
                search_test = self.ClinicalChemistryTestsDB.loc[self.ClinicalChemistryTestsDB['test'] == test_clean]
                if len(search_test) == 0:
                    for index, row in self.ClinicalChemistryTestsDB.iterrows():
                        if test_clean in row['alt_name']:
                            search_test = self.ClinicalChemistryTestsDB.loc[index]
                            break
            # Find test with similar name
            if len(search_test) == 0:
                for index, row in self.ClinicalChemistryTestsDB.iterrows():
                    if len(test) > 5 and fuzz.token_sort_ratio(test, row['test']) > 80:
                        search_test = self.ClinicalChemistryTestsDB.loc[index]
                        break
            # if found
            if len(search_test) > 0:
                thisGroup = search_test.iloc[0, 3]
                # if thisGroup is not empty
                if thisGroup == "General Chemistry":
                    TestGroup.append({'TestGroup': "General Chemistry", 'Tests': [search_test.iloc[0]]})
                elif thisGroup == "Lipid":
                    TestGroup.append({'TestGroup': "Lipid", 'Tests': [search_test.iloc[0]]})
                else:
                    TestGroup.append({'TestGroup': None, 'Tests': [search_test.iloc[0]]})
            else:
                TestGroup.append({'TestGroup': None, 'Tests': [
                    {'lab': 'chem', 'test': test, 'alt_name': "", 'specimen': "", 'code': "", 'section_order': None,
                     'is_optional': False, 'remarks': ''}]})
        return TestGroup
