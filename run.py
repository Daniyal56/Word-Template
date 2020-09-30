import os
from werkzeug.utils import secure_filename
from flask import Flask, flash, request, redirect, send_file, render_template, url_for
from docxtpl import DocxTemplate
import requests
from io import StringIO
from docx2pdf import convert
import os
from flask_caching import Cache
import win32com.client as win32
from os import path
import pythoncom


word = win32.DispatchEx("Word.Application")

UPLOAD_FOLDER = './'
app = Flask(__name__, template_folder='templates')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
cache = Cache(app, config={'CACHE_TYPE': 'simple'})


@cache.cached(timeout=3)
@app.route('/garibsons', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':

        # check if the post request has the file part
        if 'file' not in request.files:
            print('no file')
            return redirect(request.url)
        file = request.files['file']

        # if user does not select file, browser also
        # submit a empty part without filename
        if file.filename == '':
            print('no filename')
            return redirect(request.url)
        else:
            filename = secure_filename(file.filename)
            file.save(os.path.join(
                app.config['UPLOAD_FOLDER'], filename))
            print("saved file successfully")
      # send file name as parameter to
        pythoncom.CoInitialize()
        return redirect('/uploadfile/' + filename)
        # return send_file('static/GSAGROPAK.docx', as_attachment=True, attachment_filename='GSAGROPAK.docx')
    return render_template('upload_file.html')

# Download API


@ app.route("/uploadfile/<filename>", methods=['GET'])
def download_file(filename):
    doc = DocxTemplate(
        "C:/Users/Danyal/Desktop/Arwentech/mainweb/upload/{}".format(
            filename))

    if request.method == "POST":
        var_input = request.form['invoice']

        data = requests.get(
            'http://151.80.237.86:1251/ords/zkt/exprt_doc/doc?pi_no={}'.format(var_input))
        data = data.json()
#   take_input = int(input('Please enter your invoice: '))
        pythoncom.CoInitialize()
        for x in data['items']:
            # if x['pi_no'].strip() == 'GSAGROPAK- {}'.format(str()):  # 17865
            doc.render(x)
            file_stream = StringIO()
    # time.sleep(1)

            doc.save('./static/{}.docx'.format(str(file_stream)))
            convert('./static/{}.docx'.format(str(file_stream)),
                    './static/{}.pdf'.format(str(file_stream)))
    return render_template('upload_file.html')

# @ app.route('/return-files/<filename>')
# def return_files_tut(filename, method=["GET", "POST"]):
#     return send_file('static/GSAGROPAK.docx', as_attachment=True, attachment_filename='GSAGROPAK.docx')


if __name__ == "__main__":
    app.run(port=5000, debug=True, threaded=True)
