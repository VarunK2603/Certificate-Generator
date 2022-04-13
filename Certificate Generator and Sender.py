from PIL import Image, ImageDraw, ImageFont
import pandas
import os
import smtplib
from os.path import basename
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

font = ImageFont.truetype('Fonts/Playball-Regular.ttf',300)
if not os.path.exists(r'Certificates'):
    os.makedirs(r'Certificates')
#from_add = input("Enter Sender's Email Address :")
#password = input("Enter Password :")

def get_name_and_email():
    df = pandas.read_excel('Walkathon_Attendance_for_Certificates.xlsx')
    name_list = df['Name']
    email_list = df['Email']
    return name_list, email_list

def name_and_email_in_correct_format():
    nl,el = get_name_and_email()
    for i in range(len(nl)):
        split_word = nl[i].split()
        for j in range(len(split_word)):
            if len(split_word[j]) !=2:
                name_ = split_word[j]
            if len(split_word[j]) == 2:
                for k in split_word[j]:
                    name_ = name_ + " " + k
                nl[i] = name_
        nl[i] = nl[i].title()
        if '. ' in nl[i]:
            nl[i] = nl[i].replace('. ',' ')
        if '.' in nl[i]:
            nl[i] = nl[i].replace('.',' ')
    return(nl,el)

def generate_certificate(name):
    img = Image.open('Template/Certificate of Participation.png')
    draw = ImageDraw.Draw(img)
    w,h = font.getsize(name)
    draw.text(xy=((3508-w)/2,(2480-h)/2),text=name,fill=(0,107,45),font=font)
    img1 = img.convert('RGB')
    img1.save(f"Certificates/{name}.pdf")

def genearate_all_certificates(nl):
    for i in nl:
        generate_certificate(i)

name_, email_id = name_and_email_in_correct_format()
genearate_all_certificates(name_)

""" def send_all_certificates(nl,el):
    mail = MIMEMultipart()
    mail['From'] = from_add
    mail['Subject'] = 'Walkathon Certificates'
    body = MIMEText('PFA Certificates','plain')
    for i in range(len(nl)):
        to_add = f"{el[i]}"
        mail.attach(body)
        filename=f"Certificates/{nl[i]}.pdf"
        f=open(filename,'rb')
        certificate = MIMEApplication(f.read(),_subtype="pdf")
        f.close()
        certificate.add_header('Content-Disposition','attachment',filename=nl[i])
        mail.attach(certificate)
        server = smtplib.SMTP('smtp.gmail.com')
        server.starttls()
        server.login(mail['From'], password)
        server.send_message(mail, from_add, to_add)
        server.quit()
send_all_certificates(name_,email_id) """