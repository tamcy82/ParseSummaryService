# Haematology lab module
# Interpreting haematology lab test

# Path: haematology.py

# Define coagulation tests
CoagulationTests = ["pt", "aptt", "inr"]

# Match case-insensitive coagulation tests from a given list of string
# Return a list of matched tests
def match_coagulation_tests(item, description):
    matched_tests = []
    for test in description:
        if test.lower() in CoagulationTests:
            matched_tests.append(test)
    return matched_tests

# Match blood film examination
def match_blood_film_examination(item, description):
    matched_tests = ["Blood Film Examination"]
    # Define flag for blast and morphology
    additional_notes = []
    # Check if string 'blast' in description, case-insensitive with index
    for i in range(len(description)):
        if 'blast' in description[i].lower():
            additional_notes.append("Blasts")
            continue
        if 'morphology' in description[i].lower():
            additional_notes.append("General morphology")
            continue
        if 'nucleated' in description[i].lower():
            additional_notes.append("Nucleated RBC")
            continue
    if len(additional_notes) > 0:
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
        matched_tests[0] += " evulation)"
    return matched_tests

# Mach complete blood picture
def match_complete_blood_picture(item, description):
    matched_tests = ["Complete Blood Picture"]
    if 'reticulocyte' in item.lower() or 'reticulocyte' in description:
        matched_tests.append("Reticulocyte")
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

