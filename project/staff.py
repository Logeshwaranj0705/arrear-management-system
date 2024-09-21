from flask import Flask, render_template, request
import pandas as pd
import asyncio,os,openpyxl
from twilio.rest import Client


# Twilio account credentials
account_sid =""
auth_token =""
client = Client(account_sid, auth_token)
# Function to read Excel file and convert to array
def read_excel_to_array(file_path):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path)
    # Convert the DataFrame into a list of lists
    data = df.values.tolist()

    return data
def header_read(file_path):
    df = pd.read_excel(file_path)
    header=df.columns
    return header
def columns_read():
    wb = openpyxl.load_workbook('Marks1.xlsx')
    ws = wb.active
    num=list(ws.iter_cols(values_only=True))
    cols=len(num)
    return cols
def after_process():
    wb = openpyxl.load_workbook('Marks1.xlsx') 
    ws = wb.active
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None
            wb.save('Marks1.xlsx')
    return None
async def send_sms_message(ph_no, message):
    try:
        print(f"Sending message to {ph_no}...")
        message = client.messages.create(
            from_='',
            to=f"{ph_no}",
            body=message
        )
        print(f"Message sent to {ph_no} regarding arrears.")
    except Exception as e:
        print(f"Failed to send message to {ph_no}: {str(e)}")
async def login_main(login,email,password):
    if str(login)=="HOD" and str(email)=="IThod123@gmail.com" and str(password)=="hodit@123":
        stat="hod"
        return True
    elif str(login)=="Staff" and  str(email)=="jaishreekruthika12@gmail.com" and str(password)=="kruthi!12@":
        stat=False
        return stat
    else:
        stat="none"
        return stat

async def main(file_path, exam):
    print("Process started")
    cols = columns_read()
    data = read_excel_to_array(file_path)
    header = header_read(file_path)
    tasks = []
    load_count = []
    output_file = os.path.join(os.getcwd(), 'templates', 'newsheet.xlsx')
    
    # Create a new Excel file or load an existing one
    wb = openpyxl.load_workbook(output_file)
    ws = wb.active
    
    # Clear existing data in the output file
    ws.delete_cols(1, ws.max_column)
    ws.delete_rows(1, ws.max_row)
    
    # Write header to the output file
    ws.append(list(header))  # Convert header to a list
    
    # Write data to the output file
    max_column=ws.max_column+1
    ws.cell(row=1,column=max_column).value="Arrear count"
    for i in range(0, len(data)):
        ws.append(data[i])  # Append each row of data as a list
        
        # Calculate load count
        count = 0
        subject = []  
        for j in range(2, cols-1):
            if int(data[i][j]) < 25:  # Assuming scores below 25 are considered arrears
                subject.append(header[j]+'-'+str(data[i][j]))
                count += 1
        load_count.append(count)
        
        # Add load count to the last column
        ws.cell(row=i+2, column=max_column).value = count
        
        # Send SMS if load count is 3 or more
        if count >= 3:
            student_name = data[i][1]  # Assuming student name is in the second column
            phone_number = str(data[i][cols-1])  # Ensure phone number is a string
            ph_no = "+91" + phone_number
            message = f"Dear {student_name}, you have {count} arrears in {exam.upper()}. Please take necessary action."
            for k in range(0, len(subject)):
                message = message + "\n" + str(subject[k])
            tasks.append(send_sms_message(ph_no, message))
    
    # Save the output file
    wb.save(output_file)
    
    await asyncio.gather(*tasks)
    after_process()
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
@app.route('/logout',methods=['POST'])
def logout():
    return render_template('login.html')
@app.route('/back',methods=['POST'])
def back_page():
    return render_template('Staff.html')
@app.route('/login',methods=['POST'])
def login_page():
    login_user=request.form['login_user']
    email=request.form['email_user']
    password=request.form['password_user']
    loop=get_or_create_eventloop()
    stat=loop.run_until_complete(login_main(login_user,email,password))
    if stat==True:
        return render_template('hod.html')
    elif stat==False:
        return render_template('Staff.html')
    elif stat=="none":
        return render_template('login.html')

@app.route('/upload', methods=['POST'])
def my_link1():
    if request.method == 'POST':
        exam=request.form['form_sheet']
        file = request.files['file']
        file.save(os.path.join(os.getcwd(), 'Marks1.xlsx'))
        loop = get_or_create_eventloop()
        loop.run_until_complete(main('Marks1.xlsx',exam))
        return render_template('message.html')
    return "Messages not sent successfully"
# Run the Flask application
if __name__ == '__main__':
    app.run(debug=True)
