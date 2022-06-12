# import required module
import os
import fitz
# assign directory

def searchPdf(f):
    file=fitz.open(f)
    for pageNumber, page in enumerate(file.pages(), start=1):
        # print(page.getText())
        print('\n',pageNumber)


directory = 'C:\TMP\Python_PDF_Search'
 
# iterate over files in
# that directory
for filename in os.listdir(directory):
    f = os.path.join(directory, filename)
    # checking if it is a file
    if os.path.isfile(f):
        name, extension = os.path.splitext(f)
        if extension=='.pdf':
            searchPdf(f)
        else:
            print('file not supported yet')
        # print(extension)
        # print(f)