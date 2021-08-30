#Script to parse trough multiple pdf files and convert them to raw text

from numpy import column_stack
import tika, pandas as pd, os, glob
from tika import parser


#create function to parse through data with tika and return metadata and text
def pdf_to_list(file_path):
    """loops through a file and converts pdfs to its metadata and raw texts
        input: file_path given by user
        output: metadata list and raw text list"""
    os.chdir(file_path)
    metadata_l = []
    content_l = []
    for file in glob.glob('*.pdf'):
        parsed = parser.from_file(file)
        metadata_l.append(parsed['metadata'])
        content_l.append(parsed['content'])
    return metadata_l, content_l
            
#TODO: save Dataframe to csv file for future use
def list_to_df(metadata_list, text_list):
    """takes two lists and converts them to a Pandas DataFrame"""
    df = pd.DataFrame(metadata_list, text_list, columns=['Metadata', 'Text'])
    return df

fpath = input('Enter file path: ')

metadata_list, text_list = pdf_to_list(fpath)
print(metadata_list[0])
print(text_list[0])

