from PIL import Image
import pytesseract
import pdf2image
import os
import pandas as pd

os.system('cls' if os.name == 'nt' else 'clear')

# Providing the tesseract 
# executable location to pytesseract library
path_to_tesseract = r"C:\Users\daniel.lester\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = path_to_tesseract

def returnequipmentids(img):

    # Passing the image object to  image_to_string() function
    # This function will extract the text from the image.
    text = pytesseract.image_to_string(img)

    output = []

    #Look for 4 consectutive numbers preceded by a hyphen then a letter then a space
    for char in range(6,len(text)):
            if text[char].isnumeric() and text[char-1].isnumeric() and text[char-2].isnumeric() and text[char-3].isnumeric() and text[char-4] == "-" and text[char-5].isalpha() and text[char-6] == " ":
                stringtemp = text[(char-5):(char+1)]
                if not output.__contains__(stringtemp):
                    output.append(stringtemp)

    #Return all results 
    return output

Outputpath = r"C:\Users\daniel.lester\AppData\Local\Programs\Python\Projects\OCR\OUTPUT"
fulloutput = {}
popplerpath = r"C:\Users\daniel.lester\AppData\Local\Programs\Python\Projects\OCR\Release-23.01.0-0\poppler-23.01.0\Library\bin"
count = 0
numfiles = len(os.listdir(r"C:\Users\daniel.lester\AppData\Local\Programs\Python\Projects\OCR\PDF"))

for filename in os.listdir(r"C:\Users\daniel.lester\AppData\Local\Programs\Python\Projects\OCR\PDF"):
    count += 1
    print(str(count) + " out of " + str(numfiles) + " - " + str(filename))
    path = os.path.join(r"C:\Users\daniel.lester\AppData\Local\Programs\Python\Projects\OCR\PDF", filename)
    img = pdf2image.convert_from_path(path, dpi=200, poppler_path = popplerpath)
    x = returnequipmentids(img[0])
    fulloutput[filename] = x


df = pd.DataFrame(dict([(key, pd.Series(value)) for key, value in fulloutput.items()]))

df.to_excel("Results.xlsx")
