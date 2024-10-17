from flask import Flask, render_template, request
import pandas as pd
import asyncio, os, openpyxl
from twilio.rest import Client
import mysql.connector

# Load environment variables
account_sid = os.getenv('TWILIO_ACCOUNT_SID')
auth_token = os.getenv('TWILIO_AUTH_TOKEN')
twilio_client = Client(account_sid, auth_token)

# Function to read Excel file and convert to array
def read_excel_to_array(file_path):
    df = pd.read_excel(file_path)
    return df.values.tolist()

def header_read(file_path):
    df = pd.read_excel(file_path)
    return df.columns

def columns_read():
    wb = openpyxl.load_workbook('Marks1.xlsx')
    ws = wb.active
    return len(list(ws.iter_cols(values_only=True)))

def after_process():
    wb = openpyxl.load_workbook('Marks1.xlsx') 
    ws = wb.active
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None
            wb.save('Marks1.xlsx')
    return None

async def login_main(login,email,password):
    if str(login) == "HOD" and str(email) == "IThod123@gmail.com" and str(password) == "hodit@123":
        stat = "hod"
        return True
    elif str(login) == "Staff" and str(email) == "jaishreekruthika12@gmail.com" and str(password) == "kruthi!12@":
        stat = False
        return stat
    else:
        stat = "none"
        return stat

async def send_sms_message(ph_no, message):
    try:
        message = twilio_client.messages.create(
            from_='+18472428909',  # Twilio number, you can hardcode or make this an env variable too
            to=f"{ph_no}",
            body=message
        )
        print(f"Message sent to {ph_no} regarding arrears.")
    except Exception as e:
        print(f"Failed to send message to {ph_no}: {str(e)}")

def process_hod_data(year, sem, exam, arrear):
    username = os.getenv('DB_USERNAME')
    password = os.getenv('DB_PASSWORD')
    host = os.getenv('DB_HOST')

    cnx = mysql.connector.connect(user=username, password=password, host=host)
    cursor = cnx.cursor()
    data = None

    try:
        if arrear == 'three_arrear':
            cursor.execute("USE 3_arrear_data")
            query = "SELECT name, arrear_count, year, sem, exam FROM 3_arrear WHERE year = %s AND sem = %s AND exam = %s"
            cursor.execute(query, (year, sem, exam))
            data = cursor.fetchall()
        elif arrear == 'two_arrear':
            cursor.execute("USE 2_arrear_data")
            query = "SELECT name, arrear_count, year, sem, exam FROM 2_arrear WHERE year = %s AND sem = %s AND exam = %s"
            cursor.execute(query, (year, sem, exam))
            data = cursor.fetchall()
        elif arrear == 'one_arrear':
            cursor.execute("USE 1_arrear_data")
            query = "SELECT name, arrear_count, year, sem, exam FROM 1_arrear WHERE year = %s AND sem = %s AND exam = %s"
            cursor.execute(query, (year, sem, exam))
            data = cursor.fetchall()
        elif arrear == 'nil_arrear':
            cursor.execute("USE nil_arrear_data")
            query = "SELECT name, arrear_count, year, sem, exam FROM nil_arrear WHERE year = %s AND sem = %s AND exam = %s"
            cursor.execute(query, (year, sem, exam))
            data = cursor.fetchall()
        else:
            print("Invalid arrear type")

    finally:
        cursor.close()
        cnx.close()

    return data

def clear_data(arrear, year, exam, sem):
    username = os.getenv('DB_USERNAME')
    password = os.getenv('DB_PASSWORD')
    host = os.getenv('DB_HOST')

    cnx = mysql.connector.connect(user=username, password=password, host=host)
    cursor = cnx.cursor()

    try:
        if arrear == 'three_arrear':
            cursor.execute("USE 3_arrear_data")
            query = 'DELETE FROM 3_arrear WHERE year = %s AND exam = %s AND sem = %s'
            values = (year, exam, sem)
            cursor.execute(query, values)
        elif arrear == 'two_arrear':
            cursor.execute("USE 2_arrear_data")
            query = 'DELETE FROM 2_arrear WHERE year = %s AND exam = %s AND sem = %s'
            values = (year, exam, sem)
            cursor.execute(query, values)
        elif arrear == 'one_arrear':
            cursor.execute("USE 1_arrear_data")
            query = 'DELETE FROM 1_arrear WHERE year = %s AND exam = %s AND sem = %s'
            values = (year, exam, sem)
            cursor.execute(query, values)
        elif arrear == 'nil_arrear':
            cursor.execute("USE nil_arrear_data")
            query = 'DELETE FROM nil_arrear WHERE year = %s AND exam = %s AND sem = %s'
            values = (year, exam, sem)
            cursor.execute(query, values)
        else:
            print("Invalid arrear type")

    finally:
        cnx.commit()
        cursor.close()
        cnx.close()

    return None

async def main(file_path, exam, year, sem):
    print("Process started")
    cols = columns_read()
    data = read_excel_to_array(file_path)
    header = header_read(file_path)
    tasks = []
    output_file = os.path.join(os.getcwd(), 'templates', 'newsheet.xlsx')

    wb = openpyxl.load_workbook(output_file)
    ws = wb.active

    ws.delete_cols(1, ws.max_column)
    ws.delete_rows(1, ws.max_row)
    ws.append(list(header))
    max_column = ws.max_column + 1
    ws.cell(row=1, column=max_column).value = "Arrear count"

    for i in range(len(data)):
        ws.append(data[i])
        cnx = mysql.connector.connect(user=os.getenv('DB_USERNAME'), password=os.getenv('DB_PASSWORD'), host=os.getenv('DB_HOST'))
        count = 0
        subject = []

        for j in range(2, cols-1):
            if int(data[i][j]) < 25:
                subject.append(header[j] + '-' + str(data[i][j]))
                count += 1

        ws.cell(row=i+2, column=max_column).value = count

        student_data = {
            "name": data[i][1],
            "phone_number": str(data[i][cols-1]),
            "subjects": subject,
            "arrear_count": count
        }

        if count >= 3:
            phone_number = "+91" + student_data['phone_number']
            message = f"Dear {student_data['name']}, you have {count} arrears in {exam.upper()}. Please take necessary action."
            for subject_detail in subject:
                message += f"\n{subject_detail}"
            tasks.append(send_sms_message(phone_number, message))

        # Insert data into MySQL based on arrear count
        cursor = cnx.cursor()
        if count >= 3:
            cursor.execute("USE 3_arrear_data")
            cursor.execute("INSERT INTO 3_arrear (name, arrear_count, sem, exam, year) VALUES (%s, %s, %s, %s, %s)",
                           (data[i][1], count, sem, exam, year))
        elif count == 2:
            cursor.execute("USE 2_arrear_data")
            cursor.execute("INSERT INTO 2_arrear (name, arrear_count, sem, exam, year) VALUES (%s, %s, %s, %s, %s)",
                           (data[i][1], count, sem, exam, year))
        elif count == 1:
            cursor.execute("USE 1_arrear_data")
            cursor.execute("INSERT INTO 1_arrear (name, arrear_count, sem, exam, year) VALUES (%s, %s, %s, %s, %s)",
                           (data[i][1], count, sem, exam, year))
        else:
            cursor.execute("USE nil_arrear_data")
            cursor.execute("INSERT INTO nil_arrear (name, arrear_count, sem, exam, year) VALUES (%s, %s, %s, %s, %s)",
                           (data[i][1], count, sem, exam, year))
        cnx.commit()
        cursor.close()
        cnx.close()

    wb.save(output_file)
    after_process()
    await asyncio.gather(*tasks)
    print("Process completed")

# Flask web application setup
app = Flask(__name__)

def get_or_create_eventloop():
    try:
        return asyncio.get_event_loop()
    except RuntimeError as ex:
        if "There is no current event loop in thread" in str(ex):
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            return asyncio.get_event_loop()

@app.route('/')
def index():
    return render_template('login.html')

@app.route('/back', methods=['POST'])
def back():
    return render_template('login.html')

@app.route('/login', methods=['POST'])
def login():
    login = request.form['login']
    email = request.form['email']
    password = request.form['password']
    loop = get_or_create_eventloop()
    stat = loop.run_until_complete(login_main(login, email, password))

    if stat == "hod":
        return render_template('hod.html')
    elif stat == False:
        return render_template('staff.html')
    else:
        return render_template('login.html')

@app.route('/data', methods=['POST'])
def data():
    exam = request.form['exam']
    sem = request.form['sem']
    year = request.form['year']
    file = request.files['file']
    file.save(os.path.join(os.getcwd(), 'Marks1.xlsx'))
    loop = get_or_create_eventloop()
    loop.run_until_complete(main("Marks1.xlsx", exam, year, sem))
    return render_template('staff.html')

@app.route('/hod_data', methods=['POST'])
def hod_data():
    arrear = request.form['arrear']
    sem = request.form['sem']
    year = request.form['year']
    exam = request.form['exam']
    data = process_hod_data(year, sem, exam, arrear)
    return render_template('arrear_table.html', arrear_data=data)

@app.route('/hod_clear', methods=['POST'])
def hod_clear():
    arrear = request.form['arrear']
    sem = request.form['sem']
    year = request.form['year']
    exam = request.form['exam']
    clear_data(arrear, year, exam, sem)
    return render_template('hod.html')

if __name__ == '__main__':
    app.run(debug=True)
