import smtplib
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from tkinter import messagebox as mb
def GenerateOTP(email):
    global otp, OTP_entry, window_otp
    #send email
    try:
        s = smtplib.SMTP("smtp.gmail.com", 587)
        s.starttls()
        s.login("hdt.pos.noreply@gmail.com", "trinhduchieu1424")
        otp = str(random.randint(1000, 9999))
        msg = MIMEMultipart()
        msg['From'] = "POS by HDT"
        msg['To'] = email
        msg['Subject'] = 'Mã OTP - POS by HDT ©2021'
        message = 'Mã OTP của bạn là: ' + otp
        msg.attach(MIMEText(message))
        s.sendmail("hdt.pos.noreply@gmail.com", email, msg.as_string())
        s.quit()
        return otp
    except:
        mb.showerror("Thông báo","Vui lòng kiểm tra lại mạng Internet hoặc email đã nhập.")