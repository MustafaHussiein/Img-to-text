# Importing required libraries
import numpy as np
import cv2
from PIL import Image
import cv2
from IPython.display import display
from random import randrange
import numpy as np
import string
import pandas as pd
from collections import Counter
from itertools import groupby
import tensorflow as tf
tf.compat.v1.logging.set_verbosity(tf.compat.v1.logging.ERROR)
import os
import re
import sys
import comtypes.client
import pytesseract
import pytesseract as tess
tess.pytesseract.tesseract_cmd = "C:/Program Files/Tesseract-OCR/tesseract.exe"
from craft_text_detector import (
    read_image,
    load_craftnet_model,
    load_refinenet_model,
    get_prediction,
    export_detected_regions,
    export_extra_results,
    empty_cuda_cache
)
import docx
from docx.shared import Pt
import glob
# load models
refine_net = load_refinenet_model(cuda=True)
craft_net = load_craftnet_model(cuda=True)
data_folder ='Crops'
output_dir = 'Crop/'
custom_config = r'--oem 3 --psm 6'
txt = ""

#Preprocessing and sorting the extracted text
def sorted_alphanumeric(data):
    convert = lambda text: int(text) if text.isdigit() else text.lower()
    alphanum_key = lambda key: [ convert(c) for c in re.split('([0-9]+)', key) ] 
    return sorted(data, key=alphanum_key) 

#Detecting the text from the image
def detector_recognizer():
    txt = ""
    counter = 0
    isExist1 = os.path.exists(data_folder)
    if not isExist1:
       # Create a new directory because it does not exist
       os.makedirs(data_folder)
    # read image
    images_path = input("please enter path to directory of images: ")
    images_path = images_path.replace(os.sep, '/')
    lst_names = os.listdir(images_path)
    lst_Images = glob.glob(images_path + "/*.jpeg")
    for i in range(len(lst_Images)):
        img = read_image(lst_Images[7])
        prediction_result = get_prediction(
          image=img,
          craft_net=craft_net,
          refine_net=refine_net,
          text_threshold=0.5,
          link_threshold=0.2,
          low_text=0.2,
          cuda=True,
          long_size=1280
        )
        output_dir_new = output_dir + '/' + lst_names[7].split('.')[0]
        isExist2 = os.path.exists(output_dir)
        if not isExist2:

           # Create a new directory because it does not exist
           os.makedirs(output_dir)
        # export detected text regions
        exported_file_paths = export_detected_regions(
           image=img,
           regions=prediction_result["boxes"],
           output_dir=output_dir_new,
           rectify=True
        )

        empty_cuda_cache()
        counter += 1
        lst_names_crops = os.listdir(output_dir_new + '/image_crops')
        lst_names_crops.sort(key=natural_keys)
        for crop in lst_names_crops:
            txt += Recognize_Text(output_dir_new + '/image_crops/' + crop)
        print(counter)
        break
    SaveTo_W(txt)
    SaveTo_pdf(os.path.abspath("docx_file.docx"))

#Recognizing the text
def Recognize_Text(image_path):
    img = cv2.imread(image_path)
    text = pytesseract.image_to_string(img, config=custom_config, output_type=pytesseract.Output.DICT,lang='eng')
    txt = text.get('text')+"\n"
    return txt

#Sorting crops names
def atof(text):
    try:
        retval = float(text)
    except ValueError:
        retval = text
    return retval

def natural_keys(text):
    '''
    alist.sort(key=natural_keys) sorts in human order
    http://nedbatchelder.com/blog/200712/human_sorting.html
    (See Toothy's implementation in the comments)
    float regex comes from https://stackoverflow.com/a/12643073/190597
    '''
    return [ atof(c) for c in re.split(r'[+-]?([0-9]+(?:[.][0-9]*)?|[.][0-9]+)', text) ]

#Save to MS Word
def SaveTo_W(text):
    doc = docx.Document()
    lst = text.split("\n")
    for line in lst:
        line = line.rstrip()
        idx = length = len(line)
        for c in line[::-1]:
            if ord(c) > 126:
                # character with ASCII value out of printable range
                idx -= 1
            else:
                # valid-character found
                break
        if idx < length:
            # strip non-printable characters from the end of the line
            line = line[0:idx]
        doc.add_paragraph(line)
    doc.save('docx_file.docx')
    
#Convert to PDF
def SaveTo_pdf(in_file):
    wdFormatPDF = 17
    out_file = os.path.abspath("pdf_file.pdf")

    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


detector_recognizer()

