import smtplib
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
def SendUserProfile(name,dob,gender,email,phone,username,position):
    try:
        s = smtplib.SMTP("smtp.gmail.com", 587)
        s.starttls()
        s.login("hdt.pos.noreply@gmail.com", "trinhduchieu1424")
        msg = MIMEMultipart()
        msg['From'] = "POS by HDT"
        msg['To'] = email
        msg['Subject'] = 'Email chào mừng - POS by HDT ©2021'
        message = 'Chào mừng bạn đến với POS by HDT - Phần mềm bán hàng miễn phí. Sau đây là thông tin tài khoản của bạn.' + '\n' + '- Họ và tên: ' + name + '\n' + '- DOB : ' + dob + '\n' + '- Giới tính : ' + gender + '\n' + '- Email : ' + email+ '\n' + '- Di động : ' + phone +'\n' + '- Tên đăng nhập : ' + username+ '\n' + '- Chức vụ : ' + position
        msg.attach(MIMEText(message))
        s.sendmail("hdt.pos.noreply@gmail.com", email, msg.as_string())
        s.quit()
    except:
        pass
# SendUserProfile("Trịnh Đức Hiếu","06/14/05","Nam","hieu140625@gmail.com","0855860886","hieuductrinh","Admin")