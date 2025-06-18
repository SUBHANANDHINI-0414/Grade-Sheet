from flask import Flask, request
import pythoncom
import win32com.client as win32

app = Flask(__name__)

@app.route('/', methods=['GET'])
def home():
    # Simple welcome page you can open in browser
    return '''
    <h2>Welcome to Excel Data Submitter</h2>
    <p>Use POST /submit to send form data.</p>
    <p>You can test POST requests using a form, Postman, or curl.</p>
    '''

@app.route('/submit', methods=['GET', 'POST'])
def submit():
    if request.method == 'GET':
        # Show a message when accessed via browser GET (avoid 405 error)
        return '''
        <h3>This endpoint accepts POST requests with form data.</h3>
        <p>Use a POST form or tool to send data here.</p>
        '''

    # If POST request, process the form data
    try:
        pythoncom.CoInitialize()  # Needed for Excel COM in threads

        data = request.form

        # Open Excel and worksheet
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(r"C:\Users\ASUS\Desktop\ExcelWebConnect\GradeBook.xlsm")
        sheet = wb.Sheets("GradeCalculator")

        # Find next empty row in column A
        row = 2
        while sheet.Cells(row, 1).Value is not None:
            row += 1

        # Safely get values from form (default 0 for floats)
        values = {
            'reg': data.get('regno', ''),
            'name': data.get('name', ''),
            'dept': data.get('department', ''),
            'year': data.get('year', ''),
            'a1': float(data.get('a1', 0)),
            'a2': float(data.get('a2', 0)),
            'quiz': float(data.get('quiz', 0)),
            'cia1': float(data.get('cia1', 0)),
            'cia2': float(data.get('cia2', 0)),
            'att': float(data.get('att', 0)),
            'internal': float(data.get('internal', 0)),
            'external': float(data.get('external', 0))
        }

        # Clamp inputs to max allowed values
        cia1 = min(max(values['cia1'], 0), 50)
        cia2 = min(max(values['cia2'], 0), 50)
        internal = min(max(values['internal'], 0), 25)
        external = min(max(values['external'], 0), 50)

        # Calculate total as per frontend formula
        cia_final = ((cia1 + cia2) / 100) * 25
        total = cia_final + internal + external

        # Cap total at 100
        total = min(total, 100)

        # Determine pass/fail
        result = "Pass" if total >= 35 else "Fail"

        # Write data to Excel
        sheet.Cells(row, 1).Value = row - 1  # Serial No.
        sheet.Cells(row, 2).Value = values['reg']
        sheet.Cells(row, 3).Value = values['name']
        sheet.Cells(row, 4).Value = values['dept']
        sheet.Cells(row, 5).Value = values['year']
        sheet.Cells(row, 6).Value = values['a1']
        sheet.Cells(row, 7).Value = values['a2']
        sheet.Cells(row, 8).Value = values['quiz']
        sheet.Cells(row, 9).Value = cia1
        sheet.Cells(row, 10).Value = cia2
        sheet.Cells(row, 11).Value = values['att']
        sheet.Cells(row, 12).Value = internal
        sheet.Cells(row, 13).Value = external
        sheet.Cells(row, 14).Value = round(total, 2)
        sheet.Cells(row, 15).Value = result

        # Save and close workbook
        wb.Save()
        wb.Close()
        excel.Quit()

        return f"✅ Data added for {values['name']} (Total: {total:.2f}, Result: {result})"

    except Exception as e:
        return f"❌ Error: {str(e)}"

if __name__ == '__main__':
    app.run(port=8080, debug=True)
