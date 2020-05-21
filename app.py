# -*- coding: utf-8 -*-
"""
Created on Thu Apr  7 23:31:09 2020
@author: Mohit
"""
import os
import sqlite3

import pandas as pd
from cv_parser import extract_info
from flask import (Flask, flash, redirect, render_template, request,
                   send_from_directory, url_for)
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = './Uploaded_Files'  # PATH TO STORE THE UPLOADED RESUMES
DOWNLOAD_FOLDER = './Output_Files'  # PATH TO STORE THE OUTPUT EXCEL FILES
ALLOWED_EXTENSIONS = {'pdf', 'docx'}
app = Flask(__name__, static_folder='static')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.secret_key = 'super secret key'

# METHOD TO CHECK IF THE FILE EXTENSION IS PDF OR DOCX ONLY


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ROUTE TO ALLOW USER TO UPLOAD A FILE


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            saved_file = open(os.path.join(
                app.config['UPLOAD_FOLDER'], filename), "rb")
            bin_data = saved_file.read()
            database(name=filename, data=bin_data)
            return redirect(url_for('uploaded_file', filename=filename))
    return render_template("template.html")

# ROUTE TO PROCESS THE UPLOADED FILE AND CALL cv_parser.py MODULE


@app.route('/uploads/<filename>')
def uploaded_file(filename):
    path = UPLOAD_FOLDER + "/" + filename
    fname_without_ext, file_extension = os.path.splitext(filename)
    output_fname = "Data_" + fname_without_ext + ".xlsx"
    df = extract_info(path)
    writer = pd.ExcelWriter(
        "./Output_Files/" + output_fname, engine='xlsxwriter')
    df.to_excel(writer, sheet_name="Sheet 1", index=False)
    writer.save()
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename=output_fname, as_attachment="True")

# FUNCTION TO STORE THE FILE NAME AND FILE DATA (BLOB) INTO A SQLite3 DATABASE


def database(name, data):
    conn = sqlite3.connect("Resume_parser.db")
    cursor = conn.cursor()
    cursor.execute(
        """CREATE TABLE IF NOT EXISTS resume_files (name TEXT,data BLOB) """)
    cursor.execute(
        """INSERT INTO resume_files (name, data) VALUES(?,?) """, (name, data))
    conn.commit()
    cursor.close()
    conn.close()



if __name__ == "__main__":
    app.run(debug=True)
