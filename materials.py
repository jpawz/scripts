#! python3
# Extracts infromations from ProEngineer/Creo materials files.
# Finds all material files within working directory and subdirectories.
# Saves the result into Excell spreadsheet

import os

# find all material files:
materialFiles = []
for root, dirs, files in os.walk('.'):
    for file in files:
        if file.endswith('.mtl'):
            materialFiles.append(file)

# open material file:
# materialFile = open(materialFiles[i])
# materialFile.readlines()
