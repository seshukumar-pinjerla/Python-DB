import mysql.connector
from mysql.connector import Error
import pandas as pd
from openpyxl import Workbook
import zipfile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def connect_to_database(host, user, password, database):
    try:
        connection = mysql.connector.connect(
            host=host,
            user=user,
            password=password,
            database=database
        )
        if connection.is_connected():
            print("Successfully connected to the database")
            return connection
    except Error as e:
        print(f"Error: {e}")
        return None

def fetch_data(connection, query):
    try:
        cursor = connection.cursor(dictionary=True)
        cursor.execute(query)
        rows = cursor.fetchall()
        cursor.close()
        return rows
    except Error as e:
        print(f"Error: {e}")
        return None

def save_to_excel(data, filename):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)

def zip_file(file_name, zip_name):
    with zipfile.ZipFile(zip_name, 'w') as zipf:
        zipf.write(file_name)

def send_email(subject, body, to_email, attachment_path):
    from_email = 'seshukumar.pinjerla@gmail.com'
    from_password = 'aaiq qggv pddf vbkk'

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    attachment = MIMEBase('application', 'octet-stream')
    try:
        with open("C:/Users/SESHU/Desktop/Python-DB/data.zip", 'rb') as file:
            attachment.set_payload(file.read())
        encoders.encode_base64(attachment)
        attachment.add_header(
            'Content-Disposition',
            f'attachment; filename={"data.zip"}',
        )
        msg.attach(attachment)
    except FileNotFoundError:
        print(f"The file {attachment_path} was not found.")
        return

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(from_email, from_password)
            server.sendmail(from_email, to_email, msg.as_string())
            print("Email sent successfully")
    except Exception as e:
        print(f"Error: {e}")

# Main execution
if __name__ == "__main__":
    # Database connection parameters
    host = 'localhost'
    user = 'seshu'
    password = 'Seshu468$'
    database = 'attendance_tracker'
    query = 'select student_id,date, status FROM attendance'

    # Connect to the database
    connection = connect_to_database(host, user, password, database)

    if connection:
        # Fetch data from the database
        data = fetch_data(connection, query)
        if data:
            excel_file = 'data.xlsx'
            zip_file_name = 'data.zip'

            # Save data to Excel
            save_to_excel(data, excel_file)

            # Zip the Excel file
            zip_file(excel_file, zip_file_name)

            # Send email
            email_subject = 'Database Query Results'
            email_body = 'Please find the attached Excel file containing the query results.'
            recipient_email = 'seshukumar.pinjerla@gmail.com'
            send_email(email_subject, email_body, recipient_email, zip_file_name)

        # Close the database connection
        connection.close()