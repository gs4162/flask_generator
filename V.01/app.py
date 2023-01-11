from flask import Flask, request, render_template, send_file
import openpyxl

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def home():
    if request.method == 'POST':
        input1 = request.form['input1']
        input2 = request.form['input2']
        input3 = request.form['input3']

        # Write the data to an Excel spreadsheet
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 'Input 1'
        ws['B1'] = 'Input 2'
        ws['C1'] = 'Input 3'
        ws['A2'] = input1
        ws['B2'] = input2
        ws['C2'] = input3
        wb.save("test_data.xlsx")

        return render_template('success.html', input1=input1, input2=input2, input3=input3)
    return render_template('index.html')

@app.route('/download')
def download_file():
    return send_file("test_data.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name="test_data.xlsx")

if __name__ == '__main__':
    app.run(debug=True)
