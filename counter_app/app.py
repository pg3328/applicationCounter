from flask import Flask, render_template, redirect, url_for, request
from datetime import datetime
import openpyxl

app = Flask(__name__)

# Global counter
counter = 0
# Excel file path
EXCEL_FILE_PATH = "data/counter_data.xlsx"

# Create an Excel file if it doesn't exist
def create_excel_file():
    try:
        workbook = openpyxl.load_workbook(EXCEL_FILE_PATH)
    except FileNotFoundError:
        # If the file doesn't exist, create it
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Counter Data"
        sheet.append(["Timestamp", "Counter Value"])  # Header row
        workbook.save(EXCEL_FILE_PATH)

# Function to append timestamp and counter value to the Excel file
def write_to_excel(counter_value):
    create_excel_file()  # Ensure the Excel file exists
    workbook = openpyxl.load_workbook(EXCEL_FILE_PATH)
    sheet = workbook.active
    
    # Get the current timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Append the timestamp and counter value to a new row
    sheet.append([timestamp, counter_value])
    
    # Save the workbook
    workbook.save(EXCEL_FILE_PATH)

@app.route('/')
def index():
    return render_template('index.html', counter=counter)

@app.route('/increase', methods=['POST'])
def increase():
    global counter
    counter += 1
    return redirect(url_for('index'))

@app.route('/decrease', methods=['POST'])
def decrease():
    global counter
    counter -= 1
    return redirect(url_for('index'))

@app.route('/submit', methods=['POST'])
def submit():
    global counter
    write_to_excel(counter)  # Write current counter and timestamp to Excel
    return redirect(url_for('index'))

@app.route('/end_session', methods=['POST'])
def end_session():
    global counter
    counter = 0  # Reset the counter
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
