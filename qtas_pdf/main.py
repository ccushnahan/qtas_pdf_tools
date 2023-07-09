"""Program to covert Pearson unit documents into QTAS format.

Program takes Pearson pdf course documents, splits into specified subsection pdf and extracts tables from subsections, reformats tables, and then exports as a reformatted docx file.

Typical usage example:

    qtaspdf -f filename -o output_dir -s {"unit_1": [10, 15]}
"""

import os
import subprocess
import sys

import export_to_docx
import split_pdf
from loguru import logger


def main():
    """Program main method."""
    args = sys.argv
    # if len(args) <= 1:
    #     logger.info("Invalid number of args: Please add file path")
    #     exit()
    logger.info("Starting PDF table conversion process")
    file_path = "data/pdf_source/specification-l3dip-occupational-work-supervision-construction.pdf"

    split_options = {
        "unit_1": [26, 33],
        "unit_2": [33, 38],
        "unit_3": [38, 44],
        "unit_4": [44, 54],
        "unit_5": [54, 63],
        "unit_6": [63, 70],
        "unit_7": [70, 77],
        "unit_8": [77, 81],
        "unit_9": [81, 86],
        "unit_10": [86, 91],
    }
    home = os.path.expanduser("~")
    output_dir = os.path.join(home, "qtas_unit_output")
    split_files_dir = os.path.join(output_dir, "split_pdf")

    logger.info(f"Creating output directory: {output_dir}")
    try:
        os.mkdir(output_dir)
        os.mkdir(split_files_dir)
    except Exception as e:
        logger.error(e)
        exit()

    logger.info("Splitting PDF into table PDFs")
    split_pdf.split_pdf(file_path, split_options, split_files_dir)

    logger.info("Converting table PDFs to doc files")
    for file in os.listdir(split_files_dir):
        print(file)
        try:
            split_file_path = os.path.join(split_files_dir, file)
            export_to_docx.process_pdf(split_file_path)
            os.remove(split_file_path)
        except Exception as e:
            logger.error(e)

    logger.info("Conversion complete.")
    logger.info("Opening Output")
    try:
        os.startfile(file_path)
    except:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, file_path])


if __name__ == "__main__":
    main()
