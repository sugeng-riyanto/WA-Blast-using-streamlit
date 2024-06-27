import os
import streamlit as st
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
                <li><strong>SPP yang sedang berjalan:</strong> Rp. {sppbuljal} ({ket1})</li>
                <li><strong>SPP bulan-bulan sebelumnya:</strong> Rp. {spplebih} ({ket2})</li>
                <li><strong>Denda:</strong> Rp. {denda} ({ket3})</li>
                <li><strong>Keterangan:</strong> {ket4}</li>
                <li><strong>Total tagihan:</strong> Rp. {total}</li>
            </ul>
            <p>Terima kasih atas kerjasamanya.</p>
            <p>Admin Sekolah</p>
            <p>Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:</p>
            <strong>Ibu Penna (Kasir):</strong> <a href='https://bit.ly/mspennashb' style="color: #007bff;">https://bit.ly/mspennashb</a><br>
            <strong>Bapak Supatmin (Admin SMP & SMA):</strong> <a href='https://bit.ly/wamrsupatminshb4' style="color: #007bff;">https://bit.ly/wamrsupatminshb4</a>
            """
        else:
            continue

        # Email MIME setup
        msg = MIMEMultipart()
        msg['From'] = your_name
        msg['To'] = email
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'html'))

        try:
            server.sendmail(your_email, email, msg.as_string())
            st.success(f'Email successfully sent to {email}')
        except Exception as e:
            st.error(f'Failed to send email to {email}: {str(e)}')

# Streamlit UI
st.title("School Communication System")
st.header("Send Announcements, Invoices, and Proof of Payment")

option = st.selectbox("Choose the type of message you want to send", 
                      ("Announcement", "Invoice", "Proof of Payment"))

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    if allowed_file(uploaded_file.name):
        data = pd.read_excel(uploaded_file)
        
        if st.button("Send WhatsApp Messages"):
            if option == "Announcement":
                send_whatsapp_messages(data, announcement=True)
            elif option == "Invoice":
                send_whatsapp_messages(data, invoice=True)
            elif option == "Proof of Payment":
                send_whatsapp_messages(data, proof_payment=True)

        if st.button("Send Emails"):
            email_list = data.to_dict('records')
            if option == "Announcement":
                send_emails(email_list, announcement=True)
            elif option == "Invoice":
                send_emails(email_list, invoice=True)
            elif option == "Proof of Payment":
                send_emails(email_list, proof_payment=True)
    else:
        st.error("Please upload a valid Excel file.")

st.warning("Please ensure you are logged in to WhatsApp Web and have scanned the QR code before sending messages.")
