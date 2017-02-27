import glob
import os.path
import re
from docx import Document

""" Folder to search for DOCX files """
input_folder = './docx/'

""" Folder to write the TXT files to """
output_folder = './txt/'

""" Search for all the DOCX files in a given folder """
for filename in glob.glob(os.path.join(input_folder, "*.docx")):
    """ Extract the document name, without the path"""
    m = re.match(r"(.*)\.docx", os.path.basename(filename))
    document_name = m.group(1)
    """ Open the DOCX document """
    document = Document(filename)
    """ Output filename """
    output_filename = os.path.join(output_folder,
            "{:s}.txt".format(document_name))
    """ Iterate through all the paragraphs in the document and output the text
    to a TXT file """
    with open(output_filename, 'w') as output_file:
        for paragraph in document.paragraphs:
            output_file.write(paragraph.text)
