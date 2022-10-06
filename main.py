from asyncio import events
from secrets import choice
from email import message
from email.mime import image
from lib2to3.pgen2 import driver
from multiprocessing.connection import wait
from os import link
from unittest import result
from selenium import webdriver
import time
import pandas
import pyautogui
import configparser
import imghdr
import os
import smtplib
import ssl
import sys
import time
from email.message import EmailMessage
import pandas
from datetime import date, datetime, timedelta

def menu():
    print("Vui Lòng Chọn Nền Tảng By Hoàng Hữu Phúc")
    print("      1--Facebook ")
    print("      2--Linkedin")
    print("      3--Gmail")
    print("---------------------------")
def menu_facebook():
    excel_data_df = pandas.read_excel('fb.xlsx')
    message=excel_data_df['Noidung'].tolist()
    LinkGroup=excel_data_df['UID'].tolist()
    image=excel_data_df['image'].tolist()
    time.sleep(2)
    driver=webdriver.Chrome()
    driver.get("https://www.facebook.com/")
    baseDir = os.path.dirname(os.path.realpath(sys.argv[0])) + os.path.sep
    config = configparser.RawConfigParser()
    config.read(baseDir + 'mail_acount.cfg')
    driver.find_element("name","email").send_keys(config.get('CREDS', 'facebook_username'))
    driver.find_element("name","pass").send_keys(config.get('CREDS', 'facebook_password'))
    driver.find_element("name","login").click()
    for i in range (len(LinkGroup)):
        try:
            time.sleep(2)
            driver.get(LinkGroup[i])
            time.sleep(2)
            element= driver.find_element("xpath","/html/body/div[1]/div/div[1]/div/div[3]/div/div/div/div[1]/div[1]/div[4]/div/div[2]/div/div/div/div[1]/div/div/div/div[1]/div/div[1]/span").click()
            time.sleep(3)
            element = driver.find_element("xpath","/html/body/div[1]/div/div[1]/div/div[4]/div/div/div[1]/div/div[2]/div/div/div/div/div[1]/form/div/div[1]/div/div/div[1]/div/div[2]/div[1]/div[1]/div[1]/div[1]/div/div/div/div/div[2]/div/div/div/div").send_keys(message[i])
            time.sleep(1)
            clickanh=driver.find_element("xpath","/html/body/div[1]/div/div[1]/div/div[4]/div/div/div[1]/div/div[2]/div/div/div/div/div[1]/form/div/div[1]/div/div/div[1]/div/div[3]/div[1]/div[2]/div[1]/div/span/div/div/div[1]/div/div/div[1]/i").click()
            time.sleep(3)
            clickAnh=driver.find_element("xpath","/html/body/div[1]/div/div[1]/div/div[4]/div/div/div[1]/div/div[2]/div/div/div/div/div[1]/form/div/div[1]/div/div/div[1]/div/div[2]/div[1]/div[1]/div[2]/div/div[1]/div/div[1]/div/div[1]/div/div/div/div[1]/div/i").click()
            time.sleep(3)
            pyautogui.write(image[i])
            time.sleep(2)
            pyautogui.keyDown("enter")
            time.sleep(2)
            dangbai=driver.find_element("xpath","/html/body/div[1]/div/div[1]/div/div[4]/div/div/div[1]/div/div[2]/div/div/div/div/div[1]/form/div/div[1]/div/div/div[1]/div/div[3]/div[2]/div/div/div/div[1]").click()
            time.sleep(5)
        except:
            thamgia=driver.find_element("xpath","/html/body/div[1]/div[1]/div[1]/div/div[3]/div/div/div/div[2]/div[1]/div[1]/div[2]/div/div/div/div/div[2]/div/div[1]").click()
def menu_linkedin():
    baseDir = os.path.dirname(os.path.realpath(sys.argv[0])) + os.path.sep
    config = configparser.RawConfigParser()
    config.read(baseDir + 'mail_acount.cfg')
    
    tk=config.get('CREDS', 'linkedin_username')
    mk=config.get('CREDS', 'linkedin_password')
    excel_data_df = pandas.read_excel('linkedin.xlsx')
    page=excel_data_df['Link'].tolist()
    noi_dung=excel_data_df['NoiDung'].tolist()
    hinhanh=excel_data_df['hinhanh'].tolist()
    driver=webdriver.Chrome()
    driver.maximize_window()
    driver.get("https://www.linkedin.com/login")
    tai_khoan=driver.find_element("id","username").send_keys(tk)
    # time.sleep(2)
    mat_khau=driver.find_element("id","password").send_keys(mk)
    dang_nhap=driver.find_element("xpath","/html/body/div/main/div[2]/div[1]/form/div[3]/button").click()
    for i in range(len(page)):
        try:
            gr=driver.get(page[i])
            time.sleep(2)
            click_dangbai=driver.find_element("xpath","/html/body/div[5]/div[3]/div/div[2]/div/div/main/div/div[4]/div/div[1]/button/span").click()
            time.sleep(2)
            pyautogui.write(noi_dung[i])
            time.sleep(5)
            hinh_anh=driver.find_element("xpath","/html/body/div[3]/div/div/div[2]/div/div/div[2]/div[2]/div[1]/span[1]/button").click()
            time.sleep(2)
            pyautogui.write(hinhanh[i])
            pyautogui.keyDown("enter")
            time.sleep(2)
            clickk=driver.find_element("xpath","/html/body/div[3]/div/div/div[2]/div/div/div[2]/div/button").click()
            dang_bai=driver.find_element("xpath","/html/body/div[3]/div/div/div[2]/div/div/div[2]/div[2]/div[3]/button/span").click()
            time.sleep(3)
        except:
            thamgia=driver.find_element("xpath","/html/body/div[5]/div[3]/div/div[2]/div/div/main/div/div[1]/section/div/button").click()
def menu_gmail():
    hnay = datetime.now()
    x = hnay.strftime('%X' + ' ' + ' ''%x')
    baseDir = os.path.dirname(os.path.realpath(sys.argv[0])) + os.path.sep
    config = configparser.RawConfigParser()
    config.read(baseDir + 'mail_acount.cfg')
    username = config.get('CREDS', 'mail_username')
    password = config.get('CREDS', 'mail_password')
    df = pandas.read_excel(r"Danhsach.xlsx")
    for i in range(0, len(df)):
        content = df['Email Người Nhận'][i]
        title = df['Tiêu Đề (Subject)'][i]
        noidung = df['Nội Dung (Body)'][i]
        hinhanh = df['Hình Ảnh'][i]
        message = EmailMessage()
        subject = title
        body = noidung
        df.loc[i, 'Date & time'] = datetime.now()
        sender_email = username
        receiver_email = content
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject
        message.set_content(body)
        with open(hinhanh, 'rb') as m:
            file_data = m.read()
            file_type = imghdr.what(m.name)
            file_name = m.name
        message.add_attachment(file_data, maintype = 'image', subtype = file_type, filename = file_name)
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, message.as_string())
            print("Đã gửi Gmail", i + 1)
            time.sleep(10)
    df.drop(df.filter(regex="Unnamed"),axis=1, inplace=True)
    with pandas.ExcelWriter('Danhsach.xlsx', engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name="Sheet1")
        writer.save()
    print("Đã gửi tất cả mail")
def main():
    while True:
        menu()
        try: 
            chon=int(input("Nền Tảng Của Bạn Là Số Mấy:"))
            if chon ==1:
                menu_facebook()
                print("Đã xong")
            elif chon ==2:
                menu_linkedin()
                print("Đã xong")
            elif chon ==3:
                menu_gmail()
                print("Đã xong")
            else:
                print("-----Không Có trong chức năng vui lòng nhập lại-----")
        except:
            print("Không Đúng Định Dạng")
main()

