# Microbiology lab module
# Path: microbiology.py
import pandas as pd
from docx import Document
from docx.shared import Pt


# Define Microbiology Lab
class MicrobiologyLab:

    # Constructor
    def __init__(self, db_path = '.\\LocalLabTestsDB.xlsx', rr_path = '.\\MB_RI_Other Tests_20230317.docx'):
        self.micbio_tests = []
        # Load MicrobiologyTestsDB
        self.MicrobiologyTestsDBPath = db_path
        self.MicrobiologyTestsDB = self.load_microbiology_tests_db(self.MicrobiologyTestsDBPath)
        # Select only microbiology tests
        self.MicrobiologyTestsDB = self.MicrobiologyTestsDB.loc[self.MicrobiologyTestsDB['lab'] == 'micbio']
        # Convert all_name to string
        self.MicrobiologyTestsDB['alt_name'] = self.MicrobiologyTestsDB['alt_name'].astype(str)
        # Expand alt_name column
        self.MicrobiologyTestsDB['alt_name'] = self.MicrobiologyTestsDB['alt_name'].str.split(',')
        # Define HCV serology tests
        self.HCVSerologyTests = ["HCV DNA", "HCV Ab", "HCV RNA", "HCV viral load", "Anti-HCV"]
        # Define Reference Range template path
        self.MicrobiologyRRPath = rr_path
        return

    # Load MicrobiologyTestsDB into pandas
    def load_microbiology_tests_db(self, path):
        df = pd.read_excel(path, header=0)
        return df

    # Open a given microbiology form template
    def open_micro_form_template(self, path):
        document = Document(path)
        return document

    # Interpret tests
    def interpret_tests(self, item, description, all_tests):
        matched_tests = []
        HCVTest = self.match_hcv_serology(item, description, all_tests)
        # if HCVTest is not empty
        if len(HCVTest) > 0:
            # Add all tests to MicrobiologyTests if not exists
            for test in HCVTest:
                matched_tests.append(test)
            return matched_tests
        MTest = self.match_microbiology_tests(item, description, all_tests)
        # if MTest is not empty
        if len(MTest) > 0:
            # Add all tests to MicrobiologyTests if not exists
            for test in MTest:
                matched_tests.append(test)
            return matched_tests
        # No test matched
        if len(MTest) == 0:
            item_name = item.strip()
            matched_tests.append(item_name)
        return matched_tests

    # Match case-insensitive microbiology tests from a given list of string
    # Return a list of matched tests
    def match_microbiology_tests(self, item, description, all_tests):
        matched_tests = []
        # Exact match microbiology tests
        for index, T in self.MicrobiologyTestsDB.iterrows():
            # Check item
            if T['test'].lower() in item.lower():
                matched_tests.append(T['test'])
                continue
            # Check description
            # if T has space
            if " " in T['test']:
                # Split T into words
                words = T['test'].split(" ")
                # Compare T with description sequentially
                for i in range(len(words)):
                    # If T is not in description, break
                    if words[i].lower() not in description:
                        break
                    # If T is in description, and it is the last word of T
                    if words[i].lower() in description and i == len(words) - 1:
                        matched_tests.append(T['test'])
            else:
                if T['test'].lower() in description:
                    matched_tests.append(T['test'])
            # Match alternative names
            for alt_name in T['alt_name']:
                if alt_name.lower() in item.lower() or alt_name.lower() in description:
                    matched_tests.append(T['test'])
        return matched_tests

    # Match case-insensitive HCV serology tests from a given list of string
    def match_hcv_serology(self, item, description, all_tests):
        matched_tests = []
        if ("hcv" in item.lower() or "hcv" in description) and \
                ("serology" in item.lower() or "serology" in description):
            # Exact match HCV serology tests
            for index, T in enumerate(self.HCVSerologyTests):
                if T.lower() in item.lower():
                    matched_tests.append(T)
                    continue
                # Check if T has space
                if " " in T:
                    # Split T into words
                    words = T.split(" ")
                    # Compare T with description sequentially
                    for i in range(len(words)):
                        # If T is not in description, break
                        if words[i].lower() not in description:
                            break
                        # If T is in description, and it is the last word of T
                        if words[i].lower() in description and i == len(words) - 1:
                            matched_tests.append(T)
                else:
                    if T.lower() in description:
                        matched_tests.append(T)
        return matched_tests

    # Render a reference range request document
    def render_reference_range_request(self, ctcno, tests, file_name):
        # Open a reference range request template
        document = self.open_micro_form_template(self.MicrobiologyRRPath)
        if len(tests) > 0:
            # Clone the test list
            remaining_tests = tests.copy()
            # Set reference table
            RTable = document.tables[0]
            # Loop through each row of the table in reverse order
            for i in range(len(RTable.rows) - 1, 0, -1):
                # Get the test name
                test = RTable.cell(i, 0).text.strip()
                # If test is not in tests, delete the row
                if test not in tests:
                    RemoveRow = RTable.rows[i]._tr
                    RTable._tbl.remove(RemoveRow)
                else:
                    remaining_tests.remove(test)
            # If there are remaining tests, add them to the table
            if len(remaining_tests) > 0:
                for test in remaining_tests:
                    RTable.add_row()
                    RTable.cell(len(RTable.rows) - 1, 0).text = test
                    # Set cell spacing
                    RTable.cell(len(RTable.rows) - 1, 0).paragraphs[0].paragraph_format.space_before = Pt(6)
                    RTable.cell(len(RTable.rows) - 1, 0).paragraphs[0].paragraph_format.space_after = Pt(6)
                    RTable.cell(len(RTable.rows) - 1, 1).paragraphs[0].paragraph_format.space_before = Pt(6)
                    RTable.cell(len(RTable.rows) - 1, 1).paragraphs[0].paragraph_format.space_after = Pt(6)
                    RTable.cell(len(RTable.rows) - 1, 2).paragraphs[0].paragraph_format.space_before = Pt(6)
                    RTable.cell(len(RTable.rows) - 1, 2).paragraphs[0].paragraph_format.space_after = Pt(6)
            # Set footer
            footer = document.sections[0].footer
            footer_para = footer.paragraphs[0]
            footer_para.text = "Reference No.: CTC" + ctcno;
            for r in footer_para.runs:
                r.font.italic = True
            # Save the document
            document.save(file_name)
            return True
        else:
            return False