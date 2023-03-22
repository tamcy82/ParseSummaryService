# Information about blood tubes

# Path: blood_tube.py


# Define Blood Tube Colour
class BloodTubeColour:
    # Red
    clotted_blood = [0xFF, 0x00, 0x00]
    # Green
    heparin_blood = [0x00, 0xA1, 0x50]
    # Yellow
    sst_blood = [0xFF, 0xFF, 0x00]
    acd_tube = [0xFF, 0xFF, 0x00]
    # Lavender
    edta_blood = [0x70, 0x30, 0xA0]
    # Cyan
    citrate_blood = [0x00, 0x70, 0xA2]
    # Grey
    fluoride_blood = [0x80, 0x80, 0x80]
    # Black
    none = [0x00, 0x00, 0x00]


# Return blood tube colour
def get_blood_tube_colour(tube):
    # Define tube colour
    if 'clotted' in tube.lower():
        tube_colour = BloodTubeColour.clotted_blood
    elif 'heparin' in tube.lower():
        tube_colour = BloodTubeColour.heparin_blood
    elif 'sst' in tube.lower():
        tube_colour = BloodTubeColour.sst_blood
    elif 'edta' in tube.lower():
        tube_colour = BloodTubeColour.edta_blood
    elif 'citrate' in tube.lower():
        tube_colour = BloodTubeColour.citrate_blood
    elif 'fluoride' in tube.lower():
        tube_colour = BloodTubeColour.fluoride_blood
    else:
        tube_colour = BloodTubeColour.none
    return tube_colour
