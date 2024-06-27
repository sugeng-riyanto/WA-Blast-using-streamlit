import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from openpyxl import load_workbook
import time
import warnings
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

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
    # Setup Selenium WebDriver with headless Chrome
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.get("https://web.whatsapp.com")
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

        try:
            # Open chat for the phone number
            search_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
            search_box.send_keys(phone_number + Keys.ENTER)
            time.sleep(10)
            
            # Send message
            message_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"][@data-tab="6"]')
            message_box.send_keys(message + Keys.ENTER)
            
            # Wait for a few seconds to ensure the message is sent
            time.sleep(5)
            st.success(f"Message sent successfully to {phone_number}")
        except Exception as e:
            st.error(f"Failed to send message to {phone_number}: {str(e)}")

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
            name = entry['Nama_Siswa']
            email = entry['Email']
            message = f"""
            Kepada Yth.<br>Orang Tua/Wali Murid <span style="color: #007bff;">{name}</span><br>
            <p>Salam Hormat,</p>
            <p>Kami hendak menyampaikan info mengenai:</p>
            <ul>
                <li>SPP yang sedang berjalan: {entry['bulan_berjalan']}, {entry['Ket_1']}</li>
                <li>Denda: {entry['Denda']}, {entry['Ket_3']}</li>
                <li>SPP bulan-bulan sebelumnya: {entry['SPP_30hari']}, {entry['Ket_2']}</li>
                <li>Keterangan: {entry['Ket_4']}</li>
                <li>Total tagihan: {entry['Total']}</li>
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
            st.success(f'Email sent to {email}')
        except Exception as e:
            st.error(f'Failed to send email to {email}: {str(e)}')

st.title('Sekolah Harapan Bangsa Messaging System')

uploaded_file = st.file_uploader('Upload your Excel file', type=['xlsx'])

if uploaded_file is not None:
    file_name = uploaded_file.name
    if allowed_file(file_name):
        data = pd.read_excel(uploaded_file)

        st.subheader('Choose your action')
        announcement = st.checkbox('Send Announcement')
        invoice = st.checkbox('Send Invoice')
        proof_payment = st.checkbox('Send Proof of Payment')

        if st.button('Send WhatsApp Messages'):
            send_whatsapp_messages(data, announcement, invoice, proof_payment)
        if st.button('Send Emails'):
            email_list = data.to_dict('records')
            send_emails(email_list, announcement, invoice, proof_payment)
    else:
        st.error('Invalid file type. Please upload an Excel file.')

server.quit()
