from flask import Flask, request, render_template, send_file
import openpyxl

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def home():
    if request.method == 'POST':
        input1 = request.form['input1']
        input2 = request.form['input2']
        input3 = request.form['input3']
        input4 = request.form['input4']
        input5 = request.form['input5']
        input6 = request.form['input6']
        input7 = request.form['input7']
        input8 = request.form['input8']
        input9 = request.form['input9']
        input10 = request.form['input10']
        input11 = request.form['input11']
        input12 = request.form['input12']                            
        # Write the data to an Excel spreadsheet
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 'Input 1'
        ws['B1'] = 'Input 2'
        ws['C1'] = 'Input 3'
        ws['D1'] = 'Input 4'
        ws['E1'] = 'Input 5'
        ws['F1'] = 'Input 6'     
        ws['G1'] = 'Input 7'
        ws['H1'] = 'Input 8'
        ws['I1'] = 'Input 9'
        ws['J1'] = 'Input 10'
        ws['K1'] = 'Input 11'
        ws['L1'] = 'Input 12'      

        ws['A2'] = input1
        ws['B2'] = input2
        ws['C2'] = input3
        ws['D2'] = input4
        ws['E2'] = input5
        ws['F2'] = input6        
        ws['G2'] = input7
        ws['H2'] = input8
        ws['I2'] = input9
        ws['J2'] = input10
        ws['K2'] = input11
        ws['L2'] = input12     
        
      
        
        
        
        wb.save(input1+"test_data.xlsx")

        return render_template('success.html', input1=input1, input2=input2, input3=input3,input4=input4, input5=input5, input6=input6,input7=input7, input8=input8, input9=input9,input10=input10, input11=input11, input12=input12, )
        
    return render_template('index.html')

@app.route('/download')
def download_file():
    return send_file("test_data.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,
                     download_name="test_data.xlsx")

if __name__ == '__main__':
    app.run(debug=True)
