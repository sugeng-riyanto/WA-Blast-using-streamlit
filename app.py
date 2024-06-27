import os
from pyvirtualdisplay import Display

# Start a virtual display
display = Display(visible=0, size=(1024, 768))
display.start()

# Now you can import your modules
import pywhatkit as kit
import pyautogui as pg
# Your Streamlit code
import streamlit as st

import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from openpyxl import load_workbook
import pywhatkit as kit
import time
import warnings
import webbrowser



# Suppress specific warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

# SMTP configuration
your_name = "Sekolah Harapan Bangsa"
your_email = "shsmodernhill@shb.sch.id"
your_password = "jvvmdgxgdyqflcrf"

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

# Utility function to check allowed file extensions
ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def send_whatsapp_messages(data, announcement=False, invoice=False, proof_payment=False):
    # Open WhatsApp Web once
    webbrowser.open("https://web.whatsapp.com")
    st.info("Please scan the QR code in the opened WhatsApp Web window.")
    
    # Wait for user to scan QR code and login (increase if needed)
    time.sleep(45)

    for index, row in data.iterrows():
        phone_number = str(row['Phone Number'])
        if not phone_number.startswith('+62'):
            phone_number = f'+62{phone_number.lstrip("0")}'
        
        if announcement:
            message = f"""
            Kepada Yth. Orang Tua/Wali Murid *{row['Nama_Siswa']}*,
            Kami hendak menyampaikan info mengenai:
            *Subject:* {row['Subject']}
            *Description:* {row['Description']}
            *Link:* {row['Link']}
            Terima kasih atas kerjasamanya.
            Admin Sekolah
            
            Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:
            • Ibu Penna (Kasir): https://bit.ly/mspennashb
            • Bapak Supatmin (Admin SMP & SMA): https://bit.ly/wamrsupatminshb4
            """
        elif invoice:
            message = f"""
            Kepada Yth. Orang Tua/Wali Murid *{row['customer_name']}* (Kelas *{row['Grade']}*),
            Kami hendak menyampaikan info mengenai:
            • *Subject:* {row['Subject']}
            • *Batas Tanggal Pembayaran:* {row['expired_date']}
            • *Sebesar:* Rp. {row['trx_amount']:,.2f}
            • Pembayaran via nomor *virtual account* (VA) BNI/Bank: *{row['virtual_account']}*
        Terima kasih atas kerjasamanya.
        Admin Sekolah
        Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:
            • Ibu Penna (Kasir): https://bit.ly/mspennashb
            • Bapak Supatmin (Admin SMP & SMA): https://bit.ly/wamrsupatminshb4
            """
        elif proof_payment:
            message = f"""
            Kepada Yth. Orang Tua/Wali Murid *{row['Nama_Siswa']}* (Kelas *{row['Grade']}*),
            Kami hendak menyampaikan info mengenai SPP:
            • *SPP yang sedang berjalan:* {row['bulan_berjalan']:,.2f} ({row['Ket_1']})
            • *Denda:* {row['Denda']:,.2f} ({row['Ket_3']})
            • *SPP bulan-bulan sebelumnya:* {row['SPP_30hari']:,.2f} ({row['Ket_2']})
            • *Keterangan:* {row['Ket_4']}
            • *Total tagihan:* {row['Total']:,.2f}
            Terima kasih atas kerjasamanya.
            Admin Sekolah
            
            Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:
            • Ibu Penna (Kasir): https://bit.ly/mspennashb
            • Bapak Supatmin (Admin SMP & SMA): https://bit.ly/wamrsupatminshb4
            """
        else:
            continue

        while True:
            try:
                # Send WhatsApp message using the existing session
                kit.sendwhatmsg_instantly(phone_number, message, wait_time=20)
                
                # Wait for 20 seconds to ensure the message is sent
                time.sleep(20)
                st.success(f"Message sent successfully to {phone_number}")
                break  # Exit the loop if the message is sent successfully
            except Exception as e:
                st.error(f"Failed to send message to {phone_number}: {str(e)}. Retrying...")
                time.sleep(20)  # Wait before retrying

def send_emails(email_list, announcement=False, invoice=False, proof_payment=False):
    for idx, entry in enumerate(email_list):
        if announcement:
            subject = entry['Subject']
            name = entry['Nama_Siswa']
            email = entry['Email']
            description = entry['Description']
            link = entry['Link']
            message = f"""
            Kepada Yth.<br>Orang Tua/Wali Murid <span style="color: #007bff;">{name}</span><br>
            <p>Salam Hormat,</p>
            <p>Kami hendak menyampaikan info mengenai:</p>
            <ul>
                <li><strong>Subject:</strong> {subject}</li>
                <li><strong>Description:</strong> {description}</li>
                <li><strong>Link:</strong> {link}</li>
            </ul>
            <p>Terima kasih atas kerjasamanya.</p>
            <p>Admin Sekolah</p>
            <p>Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:</p>
            <strong>Ibu Penna (Kasir):</strong> <a href='https://bit.ly/mspennashb' style="color: #007bff;">https://bit.ly/mspennashb</a><br>
            <strong>Bapak Supatmin (Admin SMP & SMA):</strong> <a href='https://bit.ly/wamrsupatminshb4' style="color: #007bff;">https://bit.ly/wamrsupatminshb4</a>
            """
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
            message = f"""
            Kepada Yth.<br>Orang Tua/Wali Murid <span style="color: #007bff;">{name}</span> (Kelas <span style="color: #007bff;">{grade}</span>)<br>
            <p>Salam Hormat,</p>
            <p>Kami hendak menyampaikan info mengenai:</p>
            <ul>
                <li><strong>Subject:</strong> {subject}</li>
                <li><strong>Batas Tanggal Pembayaran:</strong> {expired_date}</li>
                <li><strong>Sebesar:</strong> Rp. {nominal}</li>
                <li><strong>Pembayaran via nomor virtual account (VA) BNI/Bank:</strong> {va}</li>
            </ul>
            <p>Terima kasih atas kerjasamanya.</p>
            <p>Admin Sekolah</p>
            <p>Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:</p>
            <strong>Ibu Penna (Kasir):</strong> <a href='https://bit.ly/mspennashb' style="color: #007bff;">https://bit.ly/mspennashb</a><br>
            <strong>Bapak Supatmin (Admin SMP & SMA):</strong> <a href='https://bit.ly/wamrsupatminshb4' style="color: #007bff;">https://bit.ly/wamrsupatminshb4</a>
            """
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
            message = f"""
            Kepada Yth.<br>Orang Tua/Wali Murid <span style="color: #007bff;">{name}</span> (Kelas <span style="color: #007bff;">{grade}</span>)<br>
            <p>Salam Hormat,</p>
            <p>Kami hendak menyampaikan info mengenai SPP:</p>
            <ul>
                <li><strong>SPP yang sedang berjalan:</strong> {sppbuljal} ({ket1})</li>
                <li><strong>Denda:</strong> {denda} ({ket3})</li>
                <li><strong>SPP bulan-bulan sebelumnya:</strong> {spplebih} ({ket2})</li>
                <li><strong>Keterangan:</strong> {ket4}</li>
                <li><strong>Total tagihan:</strong> {total}</li>
            </ul>
            <p>Terima kasih atas kerjasamanya.</p>
            <p>Admin Sekolah</p>
            <p>Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:</p>
            <strong>Ibu Penna (Kasir):</strong> <a href='https://bit.ly/mspennashb' style="color: #007bff;">https://bit.ly/mspennashb</a><br>
            <strong>Bapak Supatmin (Admin SMP & SMA):</strong> <a href='https://bit.ly/wamrsupatminshb4' style="color: #007bff;">https://bit.ly/wamrsupatminshb4</a>
            """
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
        
        # Ensure all columns are of type string
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

    # Add link to download template Excel file
    st.markdown("[Download Template Excel file](https://drive.google.com/drive/folders/1Pnpmacr7n3rS1Uht8eUI8A75KFrSA7rt?usp=sharing)")

if __name__ == '__main__':
    main()
