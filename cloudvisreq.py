from base64 import b64encode
from os import makedirs
from os.path import join, basename

from flask import Flask, request, redirect, url_for, send_from_directory, get_template_attribute, flash, send_file
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from werkzeug.utils import secure_filename
from flask import render_template
from collections import defaultdict

import json
import os
import datefinder
import re
import xlsxwriter
import glob2
import time
import shutil
import requests
import ntpath
import io



ENDPOINT_URL = 'https://vision.googleapis.com/v1/images:annotate'
RESULTS_DIR = 'jsons'
UPLOAD_DIR = 'uploads'
makedirs(RESULTS_DIR, exist_ok=True)
makedirs(UPLOAD_DIR, exist_ok=True)
alldone= False
# Initialize the Flask application
app = Flask(__name__)
filenames = []
dispfilenames = []

# This is the path to the upload directory
app.config['UPLOAD_FOLDER'] = 'uploads/'

# These are the extension that we are accepting to be uploaded
app.config['ALLOWED_EXTENSIONS'] = set([ 'pdf', 'png', 'jpg', 'jpeg','PDF'])


# For a given file, return whether it's an allowed type or not
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in app.config['ALLOWED_EXTENSIONS']


# This route will show a form to perform an AJAX request
# jQuery is loaded to execute the request and update the
# value of the operation
@app.route('/')
def index():
    return render_template('index.html')

# Route that will process the file upload
@app.route('/upload', methods=['POST'])



def upload():
    # Get the name of the uploaded file

    uploaded_files = request.files.getlist("file[]")
    totalfiles=len(glob2.glob(app.config['UPLOAD_FOLDER']+'*'))
    global filenames
    global dispfilenames
    if(totalfiles==0):
        makedirs(UPLOAD_DIR, exist_ok=True)
        filenames = []
        dispfilenames = []
    print("tell me the num of totalfiles" +str(totalfiles))
    for index, file in enumerate(uploaded_files):
        # Check if the file is one of the allowed types/extensions
        if file and allowed_file(file.filename):
            # Make the filename safe, remove unsupported chars
            file_extension = os.path.splitext(file.filename)[1]
            filename = str(secure_filename(file.filename))
            print("tell me the filename" + filename + " index "+str(index)+" totalfiles "+str(totalfiles))
            # Move the filetotalfiles form the temporal folder to the upload
            # folder we setup
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            # Save the filename into a list, we'll use it later
            filenames.append(app.config['UPLOAD_FOLDER']+filename)
            dispfilenames.append(filename)
            # Redirect the user to the uploaded_file route, which
            # will basicaly show on the browser the uploaded file
            # Load an html page with a link to each uploaded file

    return render_template("index.html", filenames=dispfilenames)



# Route that will process the uploaded files
@app.route('/process', methods=['POST'])
def process():
   try:

        new_file= processfile()
        print("done")
        print("Path" + str('root/' + new_file + '.zip'))

        checkfile(new_file)

        @app.route('/return-files/', methods=["GET"])

        def return_files_tut():
                try:
                    return send_file('root/' + new_file + '.zip', attachment_filename=new_file + '.zip', as_attachment=True)
                except Exception as e:
                    return str(e)
        return render_template('index.html', processed=True);

   except:
       shutil.rmtree(app.config['UPLOAD_FOLDER'])
       return render_template('index.html', error=True);


def checkfile(new_file):
    while True:
        bool = os.path.isfile('root/' + new_file + '.zip')
        if bool:
            print("is it here" + str(os.path.isfile('root/' + new_file + '.zip')))
            bool = os.path.isfile('root/' + new_file + '.zip')
            return bool
        time.sleep(10)


def processfile():
    image_filenames =[]
    pdf_filenames=[]
    global i
    global d
    d = defaultdict(list)
    i= 1
    for file in filenames:
        if (os.path.splitext(file)[1]=='.pdf')or(os.path.splitext(file)[1]=='.PDF'):

            pdf_filenames.append(file)

        else:
            image_filenames.append(file)
    print("tell me the files in there  " + str(pdf_filenames))
    if(image_filenames):
       d=processImage(image_filenames,d,i)
    if(pdf_filenames):
      d= processPdf(pdf_filenames,d,i)
    if d:
        writeToExcel(d)
        new_file=zipContent()
    shutil.rmtree(app.config['UPLOAD_FOLDER'])
    return new_file


def processImage(image_filenames,d,i):
<<<<<<< HEAD
        #api_key = 'AIzaSyCrOfaHEfR1yN7G5DTlGg4OsBgI_6viz-sRJ'
        api_key = '1234'
=======
        api_key = 'AIzaSyCrOfaHEfR1yN7G5DTlGg4OsBgI_6viz-srj'
        #api_key = '1234'
>>>>>>> 17ae9955e25598d401f39009cc7e955fd3fea435
        response = request_ocr(api_key, image_filenames)
        if response.status_code != 200 or response.json().get('error'):
            print(response.text)
        else:
            for idx, resp in enumerate(response.json()['responses']):
                # save to JSON file
                imgname = image_filenames[idx]
                jpath = join(RESULTS_DIR, basename(imgname) + '.json')
                with open(jpath, 'w') as f:
                    datatxt = json.dumps(resp, indent=2)
                    print("Wrote", len(datatxt), "bytes to", jpath)
                    f.write(datatxt)

                # print the plaintext to screen for convenience
                print("---------------------------------------------")
                t = resp['textAnnotations'][0]
                #print("    Bounding Polygon:")
                #print(t['boundingPoly'])
                #print("    Text:")
                print(t['description'])
                desc = t['description']
                # for date
                matches = list(datefinder.find_dates(t['description']))
                pattern = re.findall(r'([£$€$R])[\s]?(\d+(?:\.\d{2})?)', desc)
                if 'OOLA' in desc:
                    d[i].append('taxi')
                    d[i].append('OOLA')
                    d[i].append(str(matches[0].date()))
                    if 'R' in pattern[0][0]:
                        d[i].append('INR')
                        d[i].append(pattern[0][1])
                    elif '$' in pattern[0][0]:
                        d[i].append('USD')
                        d[i].append(pattern[0][1])
                    elif '€' in pattern[0][0]:
                        d[i].append('USD')
                        d[i].append(pattern[0][1])
                elif 'UBER' in desc:
                    d[i].append('taxi')
                    d[i].append('UBER')
                    d[i].append(str(matches[1].date()))
                    if 'R' in pattern[0][0]:
                        d[i].append('INR')
                        d[i].append(pattern[0][1])
                    elif '$' in pattern[0][0]:
                        d[i].append('USD')
                        d[i].append(pattern[0][1])
                    elif '€' in pattern[0][0]:
                        d[i].append('USD')
                        d[i].append(pattern[0][1])
                elif 'Room' in desc:
                    d[i].append('hotel')
                    words = desc.split("\n")
                    d[i].append(words[0])
                    d[i].append(str(matches[0].date()))
                    d[i].append('INR')
                    money = re.findall(r'[\d,]+(?:\.\d{2})[^%]',desc)
                    sum = 0.00
                    for m in money:
                        print(m)
                        m = m.replace("\n", "")
                        m = m.replace(",", "")
                        sum = sum + float(m)
                        print(sum)
                    credit=money[-2].replace("\n","")
                    credit=credit.replace(",","")
                    d[i].append(sum- float(credit))
                else:
                    d[i].append('Unrecognizable Bill')
                    d[i].append('--')
                    d[i].append('--')
                    d[i].append('--')
                    d[i].append('--')
                d[i].append(imgname)
                i += 1
            print(d)
        return d

def processPdf(pdf_filenames,d,i):
    i=len(d)+1
    print("in process pdf   "+str(d)+"  value i "+str(i))
    for file in pdf_filenames:
        path=file
        rsrcmgr = PDFResourceManager()
        retstr = io.StringIO()
        codec = 'utf-8'
        laparams = LAParams()
        device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
        fp = open(path, 'rb')
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        password = ""
        maxpages = 0
        caching = True
        pagenos = set()

        for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages,
                                  password=password,
                                  caching=caching,
                                  check_extractable=True):
            interpreter.process_page(page)

            text = retstr.getvalue()

        fp.close()
        device.close()
        retstr.close()

        matches = list(datefinder.find_dates(text))
        pattern = re.findall(r'([£$€$₹])[\s]?(\d+(?:\.\d{2})?)', text)
        print(matches)
        print(pattern)
        if 'Ola' in text:
            d[i].append('taxi')
            d[i].append('OOLA')
            d[i].append(str(matches[3].date()))
            if '₹' in pattern[0][0]:
                d[i].append('INR')
                d[i].append(pattern[0][1])
            elif '$' in pattern[0][0]:
                d[i].append('USD')
                d[i].append(pattern[0][1])
            elif '€' in pattern[0][0]:
                d[i].append('USD')
                d[i].append(pattern[0][1])
        elif 'Uber' in text:
            d[i].append('taxi')
            d[i].append('UBER')
            d[i].append(str(matches[3].date()))
            if '₹' in pattern[0][0]:
                d[i].append('INR')
                d[i].append(pattern[0][1])
            elif '$' in pattern[0][0]:
                d[i].append('USD')
                d[i].append(pattern[0][1])
            elif '€' in pattern[0][0]:
                d[i].append('USD')
                d[i].append(pattern[0][1])
        elif 'Invoice' in text:
            d[i].append('flight')
            d[i].append('international')
            d[i].append(str(matches[0].date()))
            if '₹' in pattern[0][0]:
                d[i].append('INR')
                d[i].append(pattern[0][1])
            elif '$' in pattern[0][0]:
                d[i].append('USD')
                d[i].append(pattern[0][1])
            elif '€' in pattern[0][0]:
                d[i].append('USD')
                d[i].append(pattern[0][1])
        else:
            d[i].append('flight')
            d[i].append('domestic')
            d[i].append(str(matches[3].date()))
            d[i].append('INR')
            words = text.split("\n")
            if 'Total Fare' in words:
                val = words.index('Total Fare') + 1
                d[i].append(words[val])
        d[i].append(file)
        i += 1
    print("after all "+str(d))
    return d

def writeToExcel(d):
    print(d)
    workbook = xlsxwriter.Workbook('EXPENSE_REPORT.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 'SL. No')
    worksheet.write(0, 1, 'Expense Medium')
    worksheet.write(0, 2, 'Medium Name')
    worksheet.write(0, 3, 'Date')
    worksheet.write(0, 4, 'Currency')
    worksheet.write(0, 5, 'Currency Value')
    worksheet.write(0, 6, 'Image')

    row = 1
    img_no = 0
   # for key in d.keys():
    #    worksheet.write(row, 0, key)
    #   worksheet.write_row(row, 1, d[key-2])
    #  worksheet.write(row, 6, '=HYPERLINK("'+'output\\' + ntpath.basename(d.keys()[-1])+'",'+'"'+ntpath.basename(d.keys()[-1])+'")')
    # row += 1
    #    img_no += 1
    #workbook.close()

    for dic in d.keys():
        worksheet.write(row, 0, dic)
        worksheet.write_row(row, 1, d[dic])
        index=len(d[dic])
        worksheet.write(row, 6, '=HYPERLINK("'+'output\\' + ntpath.basename(d[dic][index-1])+'",'+'"'+ntpath.basename(d[dic][index-1])+'")')
        row += 1
        img_no += 1
    workbook.close()

def zipContent():
    st = int(time.time())
    new_file = "EX_" + str(st)

    if not os.path.exists('root'):
        os.makedirs('root')

    if not os.path.exists(new_file):
        os.makedirs(new_file)

    shutil.copy('EXPENSE_REPORT.xlsx', new_file)

    def copytree(src, dst, symlinks=False, ignore=None):
        if not os.path.exists(new_file + '\\output'):
            os.makedirs(new_file + '\\output')
        for item in os.listdir(src):
            s = os.path.join(src, item)
            d = os.path.join(dst, item)
            if os.path.isdir(s):
                shutil.copytree(s, d, symlinks, ignore)
            else:
                shutil.copy2(s, d)

    copytree(app.config['UPLOAD_FOLDER'], new_file + '\\output')

    shutil.make_archive("root\\" + new_file, "zip", new_file)

    shutil.rmtree(new_file)
    os.remove('EXPENSE_REPORT.xlsx')
    return new_file


def make_image_data_list(image_filenames):
    img_requests = []
    for imgname in image_filenames:
        with open(imgname, 'rb') as f:
            ctxt = b64encode(f.read()).decode()
            img_requests.append({
                'image': {'content': ctxt},
                'features': [{
                    'type': 'TEXT_DETECTION',
                    'maxResults': 1
                }]
            })
    return img_requests


def make_image_data(image_filenames):
    """Returns the image data lists as bytes"""
    imgdict = make_image_data_list(image_filenames)
    return json.dumps({"requests": imgdict}).encode()


def request_ocr(api_key, image_filenames):
    response = requests.post(ENDPOINT_URL,
                             data=make_image_data(image_filenames),
                             params={'key': api_key},
                             headers={'Content-Type': 'application/json'})
    return response


if __name__ == '__main__':

    app.run(

    )
