import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import smtplib
from email.message import EmailMessage
import os
import time


# Load participant data
df = pd.read_csv('data_final.csv', delimiter=",")  # Assumes a column 'email'

# Read the PDF
reader = PdfReader('final cetificates.pdf')

# SMTP setup
SMTP_SERVER = 'mail.ai4science.si'  # Replace with your SMTP server
SMTP_PORT = 465                      # Or 465 for SSL
EMAIL = 'team@ai4science.si'
PASSWORD = '4wVSnD_Gxzc5Zw4'           # Use app password if needed


message = """Dear participants,

Please find your AI4Science certificate of attendance attached.

Best regards,
AI4Science team
"""



BATCH_SIZE = 30
SLEEP_SECONDS = 3600  # 1 hour

with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
    server.login(EMAIL, PASSWORD)

    for batch_start in range(0 + BATCH_SIZE * 2, len(df), BATCH_SIZE):
        batch = df.iloc[batch_start:batch_start + BATCH_SIZE]
        for idx, row in batch.iterrows():
            email = row['Email']
            name = row.get('Name', '')

            # Extract the corresponding PDF page
            writer = PdfWriter()
            writer.add_page(reader.pages[idx])
            pdf_filename = f'Certificate_of_attendance_AI4Science_{name}.pdf'
            with open(pdf_filename, 'wb') as f:
                writer.write(f)

            # Compose email
            msg = EmailMessage()
            msg['Subject'] = "AI4Science certificate of attendance"
            msg['From'] = EMAIL
            msg['To'] = email
            msg.set_content(message)

            # Attach PDF
            with open(pdf_filename, 'rb') as f:
                msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=pdf_filename)

            # Send email
            server.send_message(msg)

            # Clean up
            os.remove(pdf_filename)

        print(f"Batch {batch_start // BATCH_SIZE + 1} sent. Waiting for next batch...")
        if batch_start + BATCH_SIZE < len(df):
            time.sleep(SLEEP_SECONDS)

print("All emails sent!")
