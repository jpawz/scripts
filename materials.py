#! python3
# Extracts infromations from ProEngineer/Creo materials files.
# Finds all material files within working directory and subdirectories.
# Saves material name, density and standard into Excell spreadsheet

import os, openpyxl

wb = openpyxl.Workbook()

def findMatData(matLines, parameter, offset):
    matData = ''
    for i in range(len(matLines)):
        if parameter in matLines[i]:
            matData = matLines[i+2]
            matData = matData[offset:]
            matData = matData.rstrip('\'\n')
            break
    return matData

def extractData(matFile):
    materialData = {'name': '', 'density': '', 'standard': ''}
    fileLines = matFile.readlines()
    materialData['name'] = fileLines[2][7:].rstrip()
    materialData['density'] = findMatData(fileLines, 'PTC_MASS_DENSITY', 12)
    materialData['standard'] = findMatData(fileLines, 'NORMA_MAT', 13)
    return materialData

def findAllMatFiles():
        materialFiles = []
        for root, dirs, files in os.walk('.'):
            for file in files:
                if file.endswith('.mtl'):
                    materialFiles.append(os.path.join(root, file))

        return materialFiles;

def exportToExcell():
    materialFiles = findAllMatFiles()
    wb = openpyxl.Workbook()
    ws = wb.active # worksheet

    for mFileName in materialFiles:
        mFile = open(mFileName)
        materialData = extractData(mFile)
        print(list(materialData.values()))
        ws.append(list(materialData.values()))

    wb.save('materials.xlsx')

exportToExcell()
