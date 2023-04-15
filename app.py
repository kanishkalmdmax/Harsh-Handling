from flask import Flask, request, send_file
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill

app = Flask(__name__)

@app.route('/')
def index():
    return '''
        <html>
            <head>
                <title>Netradyne Harsh Handling E-Mail</title>
                <style>
                    #instructions {
                        border: 1px solid black;
                        padding: 10px;
                        position: absolute;
                        top: 50%;
                        right: 10px;
                        transform: translateY(-50%);
                    }
                </style>
            </head>
            <body>
        <form method="post" action="/upload" enctype="multipart/form-data">
            <input type="file" name="file">
            <input type="submit" value="Upload">
        </form>
    '''
                <div id="instructions">
                    Steps:<br>(Upload Reports in the Performance App first)<br>1. Open Performance App, then select Reports and then Driver Report.<br>2. Click on the Daily button and select yesterday's date, and click Download.<br>3. Once the file is downloaded, open this URL, click on Choose File, select the file, and click Upload.<br>4. Then a new page will open with the Download button, click on that and download the file.<br>5. Open the file, it'll be the same file, with a new sheet added with the name of "Extracted Data", open that sheet.<br>6. This sheet will have the email grid, you can copy this grid and paste it in the email.
                </div>
            </body>
        </html>
    '''

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    file_path = file.filename

    # Check if the uploaded file is in XLSX format
    if not file_path.endswith('.xlsx'):
        return '''
            <h3>Error: This is not a XLSX file</h3>
            <p>Please download &amp; upload the file from Performance App &gt; Driver Report.</p>
            <a href="/">Return to upload form</a>
            '''

    file.save(file_path)

    # Load the excel file
    try:
        df = pd.read_excel(file_path, sheet_name='Sheet1')
    except:
        return '''
            <h3>Error: This is not a valid Driver Report file</h3>
            <p>Please download the correct file from Performance App &gt; Driver Report.</p>
            <a href="/">Return to upload form</a>
            '''

    # Check if the uploaded file has the required columns
    cols_to_check = ['Following Distance', 'Camera Obstruction', 'U Turn', 'Driver Distraction', 'Seatbelt Compliance', 'Sign Violations', 'Hard Turn', 'Speeding Violations', 'Hard Braking', 'Hard Acceleration', 'Traffic Light Violation']


    if not all(col in df.columns for col in cols_to_check):
        return '''
            <h3>Error: This is not a valid Driver Report file</h3>
            <p>Please download the correct file from Performance App &gt; Driver Report.</p>
            <a href="/">Return to upload form</a>
            '''

    # Create a new dataframe to store the extracted data
    extracted_data = pd.DataFrame(columns=['Name', 'Violations'])

    # Iterate over each row in the dataframe
    for index, row in df.iterrows():
        # Check if any of the columns have a value greater than 0
        if row[cols_to_check].gt(0).any():
            # Get the name from column B
            name = row['Name']
            # Get the headers and counts of the columns with values greater than 0
            violations = ', '.join([f'{col} ({int(row[col])})' for col in cols_to_check if row[col] > 0])
            # Calculate the total number of violations for this row
            violations_count = row[cols_to_check].sum()
            # Append the data to the extracted_data dataframe using pandas.concat()
            extracted_data = pd.concat([extracted_data, pd.DataFrame({'Name': [name], 'Violations': [violations], 'Violations Count': [violations_count]})], ignore_index=False)

    # Write the extracted data to a new sheet in the same excel file using the openpyxl engine
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        extracted_data.to_excel(writer, sheet_name='Extracted Data', index=False)

    # Load the workbook and select the new sheet
    wb = load_workbook(file_path)
    ws = wb['Extracted Data']

    # Set the background color of cells A1:C1 to B8CCE4
    for row in ws['A1:C1']:
        for cell in row:
            cell.fill = PatternFill(fill_type='solid', fgColor='B8CCE4')

    # Align all data in the new sheet to the center and middle
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Adjust the column widths to fit the data
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column].width = adjusted_width

    # Apply borders to all cells with data
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border

    # Move "Extracted Data" sheet to first position and "Sheet1" to second position
    wb.active = wb.sheetnames.index('Extracted Data')
    wb.active = wb.sheetnames.index('Sheet1')
    
    # Save the changes to the workbook
    wb.save(file_path)

    return '''
        <form method="post" action="/download">
            <input type="hidden" name="file_path" value="''' + file_path + '''">
            <input type="submit" value="Download">
        </form>
    '''

@app.route('/download', methods=['POST'])
def download():
    file_path = request.form
    file_path = request.form['file_path']
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=os.environ.get('PORT', 5000),debug=True)
