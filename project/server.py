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

async def main(file_path,exam):
    print("Process started")
    cols=columns_read()
    data = read_excel_to_array(file_path)
    tasks = []
    count_num=0

    for i in range(0, len(data)):
        count = 0  
        for j in range(2,cols-1):  
            if int(data[i][j]) < 25:  # Assuming scores below 25 are considered arrears
                count += 1
        if count >= 3:
            student_name = data[i][1]  # Assuming student name is in the second column
            phone_number = str(data[i][cols-1])  # Ensure phone number is a string
            ph_no = "+91" + phone_number
            message = f"Dear {student_name}, you have {count} arrears in {exam.upper()}. Please take necessary action."
            tasks.append(send_sms_message(ph_no, message))
            count_num += 1 
    await asyncio.gather(*tasks)
    after_process()
    print("Process completed")
    return count_num

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
    return render_template('Staff.html')
# Route to trigger the SMS message sending
@app.route('/upload', methods=['POST'])
def my_link1():
    if request.method == 'POST':
        exam=request.form['form_sheet']
        file = request.files['file']
        file.save(os.path.join(os.getcwd(), 'Marks1.xlsx'))
        loop = get_or_create_eventloop()
        loop.run_until_complete(main('Marks1.xlsx',exam))
# Run the Flask application
if __name__ == '__main__':
    app.run(debug=True)