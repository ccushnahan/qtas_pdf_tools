"""
Small script to open a pdf and split it into smaller pdfs then convert to docx file

(Hard coded vals currently)
"""


from PyPDF2 import PdfFileReader, PdfFileWriter
from pdf2docx import Converter



def split_pdf(filename: str, sections: object, offset=0):

    """
    Open File 'filename' split into sections as determined by sections object.
    Save files as {section_name}.pdf

    Args:
        filename(str): File namd and path
        sections(object): object of section and name and list containing start page and end page
        offset(int): offset in case page number does not match actual pages 
    """

    input_pdf = PdfFileReader(filename)

    for key in sections.keys():
        output = PdfFileWriter()
        start, end = sections[key]
        for i in range(start + offset, end + offset):
            output.addPage(input_pdf.getPage(i))
        
        with open(f"units/{key}.pdf", "wb") as ouput_stream:
            output.write(ouput_stream)


def convert_to_word(sections: object) -> None:
    """
    Loops through sections object to open created pdfs and convert to word doc (docx)

    Args:
        sections(object): object of section and name and list containing start page and end page


    """


    # Loop through sections
    for key in sections.keys():

        # Get input + output file
        input = f"units/{key}.pdf"
        output = f"word_units/{key}.docx"

        # Attempt to convert to docx
        try:
            cv_obj = Converter(input)
            cv_obj.convert(output)
            cv_obj.close()
        except:
            print('Conversion Failed')
        else:    
            print('File Converted Successfully')



def main():
    """
    Main function
    """

    filename = "L6-NVQ-Site-Management-Spec.pdf"

    sections = {
        "unit_1": [36, 39],
        "unit_2": [40, 46],
        "unit_3": [47, 51],
        "unit_4": [52, 61],
        "unit_5": [62, 69],
        "unit_6": [70, 76],
        "unit_7": [77, 87],
        "unit_8": [88, 96],
        "unit_9": [97, 105],
        "unit_10": [106, 113],
        "unit_11": [114, 121],
        "unit_12": [122, 128],
        "unit_13": [129, 136],
        "unit_14": [137, 144],
        "unit_15": [145, 155],
        "unit_16": [156, 162],
        "unit_17": [163, 174],
        "unit_18": [175, 183],
        "unit_19": [184, 191],
        "unit_20": [192, 196],
        "unit_21": [197, 205],
        "unit_22": [206, 211],
        "unit_23": [212, 226],
        "unit_24": [227, 233],
        "unit_25": [234, 244],
        "unit_26": [245, 262],
        "unit_27": [263, 279],
        "unit_28": [280, 289],
        "unit_29": [290, 295],
        "unit_30": [296, 305],
        "unit_31": [306, 317],
    }

    split_pdf(filename, sections, 8)
    converttoword(sections)




if __name__ == "__main__":
    main()