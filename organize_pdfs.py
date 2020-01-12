# Organizes pdf drawings into folders based on sheet format

from math import isclose
from pathlib import Path

from PyPDF2 import PdfFileReader
from PyPDF2.utils import PdfReadError


def create_directory(paper_size):
    Path(paper_size).mkdir(exist_ok=True)


def get_paper_size(width, height):
    area = width * height
    if isclose(area, 210 * 297, rel_tol=0.05):
        return 'A4'
    if isclose(area, 297 * 420, rel_tol=0.05):
        return 'A3'
    if isclose(area, 420 * 594, rel_tol=0.05):
        return 'A2'
    if isclose(area, 594 * 841, rel_tol=0.05):
        return 'A1'
    if isclose(area, 841 * 1189, rel_tol=0.05):
        return 'A0'
    return 'niestandardowy'


pdfs = Path('.').glob('*.pdf')

for file in pdfs:
    try:
        document = PdfFileReader(open(file, 'rb'))
        width = document.getPage(0).mediaBox.getWidth()
        height = document.getPage(0).mediaBox.getHeight()
        paper_size = get_paper_size(width, height)
        create_directory(paper_size)
        Path(file).rename(Path(paper_size).joinpath(file))
    except PdfReadError:
        print(file.name + ' zakodowany?')
