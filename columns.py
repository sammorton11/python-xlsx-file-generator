from flask import Flask, render_template, request, send_file
import pandas as pd
import uuid
from datetime import datetime


app = Flask(__name__)

def contacts_sheet(num_rows):
    # Generate unique identifiers for the "ID" column
    ids = [str(uuid.uuid4()) for _ in range(num_rows)]

    # Create a DataFrame with the specified format
    df = pd.DataFrame({
        'Action': ['new'] * num_rows,
        'ID': ids,
        'Nick Name': [f'Nick Name {i}' for i in range(1, num_rows + 1)],
        'First Name': [f'First Name {i}' for i in range(1, num_rows + 1)],
        'Last Name': [f'Last Name {i}' for i in range(1, num_rows + 1)],
        'Email': [''] * num_rows,
        'Description': [''] * num_rows,
        'Mobile User': ['FALSE'] * num_rows,
        'Updated At': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')] * num_rows,
        'Updated By': ['administrator@rev1.co'] * num_rows,
        'Updated By Name': ['Admin Admin'] * num_rows
    })

    filename = f'output_with_uuid_{datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx'
    df.to_excel(filename, index=False)

    return filename


def work_pack_sheet(num_rows):
    # Generate unique identifiers for the "ID" column
    ids = [str(uuid.uuid4()) for _ in range(num_rows)]


    df_work_pack = pd.DataFrame({
        'Action': ['new'] * num_rows,
        'ID': ids,
        'Work Pack': [f'test {i}' for i in range(1, num_rows + 1)],
        'Description': [''] * num_rows,
        'Updated At': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')] * num_rows,
        'Updated By': ['administrator@rev1.co'] * num_rows,
        'Updated By Name': ['Admin Admin'] * num_rows
    })

    filename = f'output_workpack_{datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx'
    df_work_pack.to_excel(filename, index=False)

    return filename


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    choice = request.form['choice']
    num_rows = int(request.form['num_rows'])

    if choice == 'contacts':
        filename = contacts_sheet(num_rows)
    elif choice == 'work_pack':
        filename = work_pack_sheet(num_rows)

    return send_file(filename, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
