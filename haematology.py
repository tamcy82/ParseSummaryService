# Haematology lab module
# Interpreting haematology lab test
import pandas as pd
from docx import Document

# Path: haematology.py


class HaematologyLab:

    # Define Haematology tests
    def __init__(self, db_path = '.\\LocalLabTestsDB.xlsx'):
        # Define coagulation tests
        self.CoagulationTests = ["PT", "aPTT", "INR", "Fibrinogen", "D-Dimer"]
        self.TypingTests = ["ABO", "RhD", "Phenotyping", "Antibody Screening", "Direct Antiglobulin Test",
                       "Indirect Antiglobulin Test"]
        # Load HaematologyTestsDB
        self.HaematologyTestsDBPath = db_path
        self.HaematologyTestsDB = self.load_haematology_tests_db(self.HaematologyTestsDBPath)
        # Select only haematology tests
        self.HaematologyTestsDB = self.HaematologyTestsDB.loc[self.HaematologyTestsDB['lab'] == 'haema']
        # Convert all_name to string
        self.HaematologyTestsDB['alt_name'] = self.HaematologyTestsDB['alt_name'].astype(str)
        # Expand alt_name column
        self.HaematologyTestsDB['alt_name'] = self.HaematologyTestsDB['alt_name'].str.split(',')

    # Load HaematologyTestsDB into pandas
    def load_haematology_tests_db(self, path):
        df = pd.read_excel(path, header=0)
        return df

    # Open form word file
    def open_haema_form_template(self, path):
        document = Document(path)
        return document

    # Interpret tests
    def interpret_tests(self, item, description, all_tests):
        matched_tests = []
        CTest = self.match_coagulation_tests(item, description, all_tests)
        # if CTest is not empty
        if len(CTest) > 0:
            # Add all tests to HaematologyTests if not exists
            for test in CTest:
                matched_tests.append(test)
        # Blood Film Examination
        BloodFilms = self.match_blood_film_examination(item, description, all_tests)
        if len(BloodFilms) > 0:
            # Add all tests to HaematologyTests if not exists
            for test in BloodFilms:
                matched_tests.append(test)
        # Complete Blood Picture
        CBC = self.match_complete_blood_picture(item, description, all_tests)
        if len(CBC) > 0:
            # Add all tests to HaematologyTests if not exists
            for test in CBC:
                matched_tests.append(test)
        # Typing
        BTyping = self.match_blood_typing(item, description, all_tests)
        if len(BTyping) > 0:
            # Add all tests to HaematologyTests if not exists
            # print(BTyping)
            for test in BTyping:
                matched_tests.append(test)
        # None matched
        if len(CTest) == 0 and len(BloodFilms) == 0 and len(CBC) == 0 and len(BTyping) == 0:
            # Add test to HaematologyTests if not exists
            item_name = item.strip()
            matched_tests.append(item_name)
        return matched_tests

    # Match case-insensitive coagulation tests from a given list of string
    # Return a list of matched tests
    def match_coagulation_tests(self, item, description, all_tests):
        matched_tests = []
        # Exact match coagulation tests
        for index, T in enumerate(self.CoagulationTests):
            if T.lower() in item.lower() or T.lower() in description:
                matched_tests.append(self.CoagulationTests[index])
        return matched_tests

    # Match blood film examination
    def match_blood_film_examination(self, item, description, all_tests):
        if 'examination of blood film' not in item.lower():
            return []
        matched_tests = ["Blood Film Examination"]
        # Define flag for blast and morphology
        additional_notes = []
        # Check if string 'blast' in description, case-insensitive with index
        for i in range(len(description)):
            if 'blast' in description[i].lower():
                # Add to additional note if not exists
                if "Blasts" not in additional_notes:
                    additional_notes.append("Blasts")
                    continue
            if 'morphology' in description[i].lower():
                # Add to additional note if not exists
                if "Morphology" not in additional_notes:
                    additional_notes.append("General Morphology")
                    continue
            if 'nucleated' in description[i].lower():
                # Add to additional note if not exists
                if "Nucleated" not in additional_notes:
                    additional_notes.append("Nucleated RBC")
                    continue
        if len(additional_notes) > 0:
            additional_notes.sort()
            matched_tests[0] += " (include"
            # Loop through additional notes with index
            for i in range(len(additional_notes)):
                matched_tests[0] += " " + additional_notes[i]
                # if last two items, add 'and'
                if i == len(additional_notes) - 2:
                    matched_tests[0] += " and"
                # else add comma
                elif i < len(additional_notes) - 1:
                    matched_tests[0] += ","
            matched_tests[0] += " evaluation)"
        return matched_tests

    # Mach complete blood picture
    def match_complete_blood_picture(self, item, description, all_tests):
        if 'haematology' in item.lower() or 'count' in description:
            matched_tests = ["Complete Blood Picture"]
            if 'reticulocyte' in item.lower() or 'reticulocyte' in description:
                matched_tests.append("Reticulocytes")
            if 'fibrinogen' in item.lower() or 'fibrinogen' in description:
                matched_tests.append("Fibrinogen")
            return matched_tests
        elif 'reticulocyte' in item.lower() and "Complete Blood Picture" in all_tests:
            return ["Reticulocytes"]
        else:
            return []

    # Match blood typing
    def match_blood_typing(self, item, description, all_tests):
        matched_tests = []
        # Exact match typing tests
        for index, T in enumerate(self.TypingTests):
            if T.lower() in item.lower() or T.lower() in description:
                matched_tests.append(self.TypingTests[index])
        return matched_tests

    # Match Trephine Biopsy IHC Reporting
    def match_trephine_biopsy(self, description):
        return ["Trephine Biopsy IHC Reporting"]

    # Match Bone Marrow Aspirate Reporting
    def match_bone_marrow_aspirate(self, description):
        return ["Bone Marrow Aspirate Reporting"]

    # Match Bone Marrow Smear Preparation
    def match_bone_marrow_smear(self, description):
        return ["Bone Marrow Smear Preparation"]

    # Match Cytogenetics
    def match_cytogenetics(self, description):
        return ["Cytogenetics"]

    # Match Blood Phenotyping
    def match_blood_phenotyping(self, description):
        return ["Blood Phenotyping"]

    # Create a list of Haematology tests from a text string
    def render_haematology_test_group(self, T):
        TestGroup = []
        for test in T:
            SearchTest = self.HaematologyTestsDB.loc[self.HaematologyTestsDB['test'] == test]
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
                        TestGroup.append({'TestGroup': thisGroup, 'Tests': [SearchTest.iloc[0]]})
                else:
                    TestGroup.append({'TestGroup': None, 'Tests': [SearchTest.iloc[0]]})
            else:
                TestGroup.append({'TestGroup': None, 'Tests': [
                    {'test': test, 'description': None, 'specimen': "", 'alt_name': "", 'interpretation': None,
                     'is_optional': False, 'remarks': ''}]})
        return TestGroup
