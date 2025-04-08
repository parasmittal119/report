from flask import Flask, render_template, request, send_file
import pandas as pd
import numpy as np
import os
import zipfile
import shutil
from werkzeug.utils import secure_filename
import tempfile

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def upload_files():
    if request.method == "POST":
        uploaded_files = request.files.getlist("files")
        if not uploaded_files:
            return "No files uploaded."

        temp_dir = tempfile.mkdtemp()

        for file in uploaded_files:
            filename = secure_filename(file.filename)
            filepath = os.path.join(temp_dir, filename)
            file.save(filepath)

            # Create subfolder for output
            output_folder = os.path.join(temp_dir, os.path.splitext(filename)[0])
            os.makedirs(output_folder, exist_ok=True)

            try:
                # Fault Summary
                df1 = pd.read_excel(filepath, sheet_name=[x for x in range(1, 32)], header=None, skiprows=range(0, 35), usecols=range(18))
                df3 = pd.concat(df1, sort=False)
                df3.columns = ['S.NO', 'DATE', 'STATION', 'PART NO.', 'PROJECT', 'System S.No.', 'CABLE', 'ASSY', 'CARD', 'ATS',
                               'ENGG/R&D', 'MATERIAL', 'Time Loss (minutes)', 'FAULT DESCRIPTION',
                               'ACTION TAKEN (TESTING TEAM)', 'REPAIRING ACTION', 'FAULTY CARD SERIAL NOS.', 'SYSTEM RESULT']
                df3['FAULT DESCRIPTION'].replace('', np.nan, inplace=True)
                df3.dropna(subset=['FAULT DESCRIPTION'], inplace=True)
                df3 = df3.iloc[:, 1:18]
                df3 = df3.drop('Time Loss (minutes)', axis=1)
                df3.to_excel(os.path.join(output_folder, 'Fault_Summary_DCT.xlsx'), index=False)

                # Product wise
                df2 = pd.read_excel(filepath, sheet_name=[x for x in range(1, 32)], header=None, skiprows=range(16, 76), usecols=range(0, 7))
                df4 = pd.concat(df2, sort=False)
                df4.columns = ['STATION', 'PART NO', 'PRODUCT ', 'FTY % ', 'ATTEMPTED', 'FTY NO.', '']
                df4['STATION'].replace('', np.nan, inplace=True)
                df4.dropna(subset=['STATION'], inplace=True)
                df4 = df4[df4["STATION"].str.contains("STATION") == False]
                df4.to_excel(os.path.join(output_folder, 'Product_wise.xlsx'), index=False)

            except Exception as e:
                print(f"Error processing {filename}: {e}")

        # Create zip
        zip_filename = os.path.join(temp_dir, "processed_files.zip")
        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    if file != "processed_files.zip":
                        zipf.write(os.path.join(root, file),
                                   arcname=os.path.relpath(os.path.join(root, file), temp_dir))

        return send_file(zip_filename, as_attachment=True)

    return render_template("index.html")
