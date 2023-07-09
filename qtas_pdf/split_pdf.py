"""Open pdf and split it into subsections then save as pdfs.

Small script to open a pdf and split it into smaller pdfs. Is
supplied dictionary object containing section names and a list
of section starts and ends.  Iterates through dictionary items
to extract each sub section and save as a pdf under dictionary
key as file name.

Typical Usage:

    python3 split_pdf.py

    OR

    split_pdf.split_pdf("example.pdf", {"section": [1,10]}, "output_dir", 0)
"""

import os

from pdf2docx import Converter
from PyPDF2 import PdfFileReader, PdfFileWriter


def split_pdf(
    filename: str, sections: dict, split_pdf_dir: str = "", offset: int = 0
) -> None:
    """Splits pdf into sub pdfs.

    Open File 'filename' split into sections as determined by sections object.
    Save files as {section_name}.pdf

    Args:
        filename (str):         File namd and path
        sections (dict):        Dict of section and name and list containing
                                start page and end page
        split_pdf_dir (str):    Dir path for output
        offset (int):           offset in case page number does
                                not match actual pages
    """

    input_pdf = PdfFileReader(filename)

    for key in sections.keys():
        output = PdfFileWriter()
        start, end = sections[key]
        for i in range(start + offset, end + offset):
            output.addPage(input_pdf.getPage(i))

        file_path = os.path.join(split_pdf_dir, f"{key}.pdf")
        with open(file_path, "wb") as ouput_stream:
            output.write(ouput_stream)


def convert_to_word(sections: object) -> None:
    """Converts pdf to word document.

    Loops through sections object to open created pdfs and 
    convert to word doc (docx)

    Args:
        sections(object): object of section and name and list 
        containing start page and end page
    """
    # Loop through sections
    for key in sections.keys():

        # Get input + output file
        input = f"data/units/{key}.pdf"
        print(input)
        output = f"data/word_units/{key}.docx"
        print(output)
        # Attempt to convert to docx
        try:
            cv_obj = Converter(input)
            cv_obj.convert(output)
            cv_obj.close()
        except:
            print("Conversion Failed")
        else:
            print("File Converted Successfully")


def discover_sections(filename: str) -> dict:
    """Find and return table sections of document.

    Args:
        filename (str):     filename of pdf to extract

    Returns:
        sect_dict (dict):   dictionary of chap names and start end
    """
    pass


def make_output_dirs() -> None:
    """Make directories for output."""
    pass


def main():
    """The Main function.
    """

    filename = "data/pdf_source/specification-l3dip-occupational-work-supervision-construction.pdf"

    sections = {
        "unit_1": [24, 29],
        "unit_2": [34, 37],
        "unit_3": [35, 42],
        "unit_4": [45, 53],
        "unit_5": [55, 62],
        "unit_6": [64, 69],
        "unit_7": [71, 75],
        "unit_8": [78, 80],
        "unit_9": [82, 85],
        "unit_10": [87, 90],
    }

    split_pdf(filename, sections, "data/units", 0)
    convert_to_word(sections)


if __name__ == "__main__":
    main()
