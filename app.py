#from flask import Flask, redirect, render_template, request, url_for
from flask import Flask, render_template, request, send_file

import zipfile
import pandas as pd
from io import BytesIO


app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    uploaded_file = request.files['file']
    if uploaded_file.filename.endswith('.zip'):
        z = zipfile.ZipFile(uploaded_file)
        dataframes = []

        content = z.read(z.namelist()[0])
        content_as_file = BytesIO(content)
        tmp = pd.read_excel(content_as_file)
        e_tmp1 = tmp.iloc[1::2,:]
        e_tmp1.reset_index(drop=True , inplace=True)
        e_tmp2 = tmp.iloc[0::2,:]
        e_tmp2 = e_tmp2.rename(columns=e_tmp2.iloc[0])
        e_tmp2 = e_tmp2.drop(e_tmp2.columns[0], axis=1)
        e_tmp2 = e_tmp2.drop(e_tmp2.index[0])
        e_tmp2.reset_index(drop=True , inplace=True)
        result2 = pd.concat([e_tmp1,e_tmp2], axis=1)
        result2['ACC_NUM'] = z.namelist()[0][-28:-5]

        for i in range(1, len(z.namelist())):
            content = z.read(z.namelist()[i])
            content_as_file = BytesIO(content)
            tmp = pd.read_excel(content_as_file)

            e_tmp1 = tmp.iloc[1::2,:]
            e_tmp1.reset_index(drop=True , inplace=True)
            e_tmp2 = tmp.iloc[0::2,:]
            e_tmp2 = e_tmp2.rename(columns=e_tmp2.iloc[0])
            e_tmp2 = e_tmp2.drop(e_tmp2.columns[0], axis=1)
            e_tmp2 = e_tmp2.drop(e_tmp2.index[0])
            e_tmp2.reset_index(drop=True , inplace=True)
            result1 = pd.concat([e_tmp1,e_tmp2], axis=1)
            result1['ACC_NUM'] = z.namelist()[i][-28:-5]
            result2 = pd.concat([result2,result1], ignore_index=True)
        
        combined_df = result2
        buffer = BytesIO()
        combined_df.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)
        return send_file(buffer, as_attachment=True, download_name='merge_excel.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        return "Invalid file. Please upload a zip file containing .xlsx files.", 400

if __name__ == '__main__':
    app.run(debug=True)
