import os
import subprocess
import streamlit as st

try:
    from pyvirtualdisplay import Display
    # Check if Xvfb is installed
    if subprocess.call(['which', 'Xvfb']) == 0:
        # Start a virtual display
        display = Display(visible=0, size=(1024, 768))
        display.start()
        using_virtual_display = True
    else:
        st.warning("Xvfb is not installed. Skipping virtual display setup.")
        using_virtual_display = False
except ImportError:
    st.warning("pyvirtualdisplay is not installed. Skipping virtual display setup.")
    using_virtual_display = False

import pywhatkit as kit
import pyautogui as pg
import mouseinfo
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from openpyxl import load_workbook
import time
import warnings
import webbrowser

warnings.simplefilter(action='ignore', category=FutureWarning)

# SMTP configuration
your_name = "Sekolah Harapan Bangsa"
your_email = "shsmodernhill@shb.sch.id"
your_password = "jvvmdgxgdyqflcrf"

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def send_whatsapp_messages(data, announcement=False, invoice=False, proof_payment=False):
    webbrowser.open("https://web.whatsapp.com")
    st.info("Please scan the QR code in the opened WhatsApp Web window.")
    time.sleep(45)
    for index, row in data.iterrows():
        phone_number = str(row['Phone Number'])
        if not phone_number.startswith('+62'):
            phone_number = f'+62{phone_number.lstrip("0")}'
        
        if announcement:
            message = f"""..."""
        elif invoice:
            message = f"""..."""
        elif proof_payment:
            message = f"""..."""
        else:
            continue

        while True:
            try:
                kit.sendwhatmsg_instantly(phone_number, message, wait_time=20)
                time.sleep(20)
                st.success(f"Message sent successfully to {phone_number}")
                break
            except Exception as e:
                st.error(f"Failed to send message to {phone_number}: {str(e)}. Retrying...")
                time.sleep(20)

def send_emails(email_list, announcement=False, invoice=False, proof_payment=False):
    for idx, entry in enumerate(email_list):
        if announcement:
            subject = entry['Subject']
            name = entry['Nama_Siswa']
            email = entry['Email']
            description = entry['Description']
            link = entry['Link']
            message = f"""..."""
        elif invoice:
            subject = entry['Subject']
            grade = entry['Grade']
            va = entry['virtual_account']
            name = entry['customer_name']
            email = entry['customer_email']
            nominal = "{:,.2f}".format(entry['trx_amount'])
            expired_date = entry['expired_date']
            expired_time = entry['expired_time']
            description = entry['description']
            link = entry['link']
            message = f"""..."""
        elif proof_payment:
            subject = entry['Subject']
            va = entry['virtual_account']
            name = entry['Nama_Siswa']
            email = entry['Email']
            grade = entry['Grade']
            sppbuljal = "{:,.2f}".format(entry['bulan_berjalan'])
            ket1 = entry['Ket_1']
            spplebih = "{:,.2f}".format(entry['SPP_30hari'])
            ket2 = entry['Ket_2']
            denda = "{:,.2f}".format(entry['Denda'])
            ket3 = entry['Ket_3']
            ket4 = entry['Ket_4']
            total = "{:,.2f}".format(entry['Total'])
            message = f"""..."""
        else:
            continue

        msg = MIMEMultipart()
        msg['From'] = your_email
        msg['To'] = email
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'html'))

        try:
            server.sendmail(your_email, email, msg.as_string())
            st.success(f'Email {idx + 1} to {email} successfully sent!')
        except Exception as e:
            st.error(f'Failed to send email {idx + 1} to {email}: {e}')

def handle_file_upload(announcement=False, invoice=False, proof_payment=False):
    uploaded_file = st.file_uploader("Upload Excel file", type="xlsx")
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        df = df.astype(str)
        email_list = df.to_dict(orient='records')
        st.dataframe(df)

        if st.button("Send Emails"):
            send_emails(email_list, announcement, invoice, proof_payment)
        
        if st.button("Send WhatsApp Messages"):
            send_whatsapp_messages(df, announcement, invoice, proof_payment)

def main():
    st.title('Communication Sender for SHB')
    menu = ["Home", "Invoice", "Send Reminder", "Announcement"]
    choice = st.sidebar.selectbox("Menu", menu)

    if choice == "Home":
        st.subheader("Home")
        st.write("Welcome to the Communication Sender App!")

    elif choice == "Announcement":
        st.subheader("Announcement")
        handle_file_upload(announcement=True)

    elif choice == "Invoice":
        st.subheader("Invoice")
        handle_file_upload(invoice=True)

    elif choice == "Send Reminder":
        st.subheader("Send Reminder")
        handle_file_upload(proof_payment=True)

    st.markdown("[Download Template Excel file](https://drive.google.com/drive/folders/1Pnpmacr7n3rS1Uht8eUI8A75KFrSA7rt?usp=sharing)")

if __name__ == '__main__':
    main()

if using_virtual_display:
    display.stop()
