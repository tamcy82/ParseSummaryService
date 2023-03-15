# Haematology lab module
# Interpreting haematology lab test
import pandas as pd
from docx import Document

# Path: haematology.py

# Define coagulation tests
CoagulationTests = ["PT", "aPTT", "INR", "Fibrinogen", "D-Dimer"]
TypingTests = ["ABO", "RhD", "Phenotyping", "Antibody Screening", "Direct Antiglobulin Test", "Indirect Antiglobulin Test"]


# Load HaematologyTestsDB into pandas
def load_haematology_tests_db(path):
    df = pd.read_excel(path, header=0)
    return df


# Open form word file
def OpenHaemaFormTemplate(path):
    document = Document(path)
    return document


# Match case-insensitive coagulation tests from a given list of string
# Return a list of matched tests
def match_coagulation_tests(item, description, allTests):
    matched_tests = []
    # Exact match coagulation tests
    for index, T in enumerate(CoagulationTests):
        if T.lower() in item.lower() or T.lower() in description:
            matched_tests.append(CoagulationTests[index])
    return matched_tests


# Match blood film examination
def match_blood_film_examination(item, description, allTests):
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
def match_complete_blood_picture(item, description, allTests):
    if 'haematology' in item.lower() or 'count' in description:
        matched_tests = ["Complete Blood Picture"]
        if 'reticulocyte' in item.lower() or 'reticulocyte' in description:
            matched_tests.append("Reticulocytes")
        if 'fibrinogen' in item.lower() or 'fibrinogen' in description:
            matched_tests.append("Fibrinogen")
        return matched_tests
    elif 'reticulocyte' in item.lower() and "Complete Blood Picture" in allTests:
        return ["Reticulocytes"]
    else:
        return []


# Match blood typing
def match_blood_typing(item, description, allTests):
    matched_tests = []
    # Exact match typing tests
    for index, T in enumerate(TypingTests):
        if T.lower() in item.lower() or T.lower() in description:
            matched_tests.append(TypingTests[index])
    return matched_tests


# Match Trephine Biopsy IHC Reporting
def match_trephine_biopsy(description):
    return ["Trephine Biopsy IHC Reporting"]


# Match Bone Marrow Aspirate Reporting
def match_bone_marrow_aspirate(description):
    return ["Bone Marrow Aspirate Reporting"]


# Match Bone Marrow Smear Preparation
def match_bone_marrow_smear(description):
    return ["Bone Marrow Smear Preparation"]


# Match Cytogenetics
def match_cytogenetics(description):
    return ["Cytogenetics"]


# Match Blood Phenotyping
def match_blood_phenotyping(description):
    return ["Blood Phenotyping"]


# Create a list of Haematology tests from a text string
def render_haematology_test_group(T):
    # Load HaematologyTestsDB
    HaematologyTestsDBPath = '.\\HaematologyTestsDB.xlsx'
    HaematologyTestsDB = load_haematology_tests_db(HaematologyTestsDBPath)
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
                    TestGroup.append({'TestGroup': thisGroup, 'Tests': [SearchTest.iloc[0]]})
            else:
                TestGroup.append({'TestGroup': None, 'Tests': [SearchTest.iloc[0]]})
        else:
            TestGroup.append({'TestGroup': None, 'Tests': [
                {'test': test, 'description': None, 'specimen': "", 'interpretation': None, 'is_optional': False}]})
    return TestGroup