#需要的东西：对应系统的chrome、chromedriver、python3.11.5，和下面的依赖：xlrd、xlwt、ddddocr、xlutils、selenium
# Name: while_try_InvitationCode
# Author: massorant
# MakeTime: 2024-04-20
# -----------------开始-------------------
#GUI
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import threading
import json # 保存配置
#核心逻辑
import os
import time
import xlrd
import xlwt
import ddddocr
from xlutils.copy import copy
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
#邮箱（为了远程提醒解码成功）
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

app = tk.Tk()
app.title("无标题") #gui标题

# 全局变量
email_config = {"sender": "", "password": "", "receiver": ""} #邮箱参数
excel_file_path = "" #全局xls文件路径
driver = None
CONFIG_FILE_PATH = "app_config.json"

#配置文件的读写
def save_config(config_data):
    with open(CONFIG_FILE_PATH, 'w') as f:
        json.dump(config_data, f)

def load_config():
    try:
        with open(CONFIG_FILE_PATH, 'r') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

#读取配置文件
def initialize_app():
    config_data = load_config()
    global email_config, excel_file_path
    email_config = config_data.get("email_config", {"sender": "", "password": "", "receiver": ""})
    excel_file_path = config_data.get("excel_file_path", "")
    if excel_file_path:
        file_path_entry.insert(0, excel_file_path)

#发送邮箱函数，接受参数，发送内容
def send_email(sender_email, password, receiver_email, subject, body):
    try:
        smtp_server = 'smtp.qq.com'
        server = smtplib.SMTP(smtp_server, 587)
        server.starttls()
        server.login(sender_email, password)
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver_email
        message['Subject'] = subject
        message.attach(MIMEText(body, 'plain'))
        server.sendmail(sender_email, receiver_email, message.as_string())
        server.quit()
        print("邮件发送成功")
    except Exception as e:
        print("邮件发送失败:", e)

#初始化参数：chrome驱动路径、chrome路径、网址、检查表格路径
def init_webdriver():
    global driver, book, sheet, write_book, write_sheet, log_file
    chrome_path = './chrome-win64/chrome.exe'
    options = Options()
    options.binary_location = chrome_path
    service = Service('./chromedriver.exe')
    driver = webdriver.Chrome(service=service, options=options)
    driver.get("https://www.google.com")

    if excel_file_path:
        log_file = open("automation_log.txt", "a")
        book = xlrd.open_workbook(excel_file_path)
        sheet = book.sheet_by_index(0)
        write_book = copy(book)
        write_sheet = write_book.get_sheet(0)
    else:
        messagebox.showerror("错误", "未选择Excel文件")
        return
    
#保存验证码截图到本地
def save_captcha():
    if not os.path.exists('vcode'):
        os.makedirs('vcode')
    captcha_element = driver.find_element(By.ID, "codeImg")
    captcha_path = 'vcode/captcha.png'
    captcha_element.screenshot(captcha_path)
    return captcha_path

#打开本地验证码截图进行识别
def recognize_captcha(image_path):
    ocr = ddddocr.DdddOcr()
    with open(image_path, 'rb') as f:
        image = f.read()
    return ocr.classification(image)

#处理网站验证码部分：刷新--调用保存--调用识别--输入结果
def handle_captcha():
    try:
        WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='codeImg']"))).click()
        time.sleep(2)
        captcha_path = save_captcha()
        captcha = recognize_captcha(captcha_path)
        driver.find_element(By.XPATH, "//*[@id='validate']").send_keys(captcha)
    except Exception as e:
        driver.find_element(By.XPATH, "//*[@id='codeImgText']").click()
        print("错误的验证码", e)

#创造子线程，防止阻塞主gui活动
def process():
    threading.Thread(target=run_process).start()

#主体逻辑：处理验证码部分--填写表格内容进输入框--点击检查按钮--检查返回信息
def run_process():
    init_webdriver()
    if not excel_file_path:
        return

    start_row = int(start_row_entry.get())
    rx = start_row

    while rx < sheet.nrows:
        try:
            inv_code = sheet.cell(rx, 0).value
            driver.find_element(By.ID, "invcode").clear()
            driver.find_element(By.ID, "invcode").send_keys(inv_code)
            handle_captcha()
            driver.find_element(By.XPATH, "//*[@id='main']/form[3]/div[1]/table/tbody/tr[7]/th[2]/input[2]").click()
            time.sleep(2)

            try:
                result_element = driver.find_element(By.XPATH, "//*[@id='check_info_invcode']/span")
                result_text = result_element.text
            except:
                result_element = driver.find_element(By.XPATH, "//*[@id='check_info_invcode']")
                result_text = result_element.text
                continue

            log_text.insert(tk.END, f"Row {rx + 1}: {result_text}\n")
            log_file.write(f"Row {rx + 1}: {result_text}\n")

            if "驗證碼不正確" in result_text:
                continue
            elif "恭喜" in result_text:
                write_sheet.write(rx, 2, 'yes')
                if email_var.get():
                    send_email(email_config["sender"], email_config["password"], email_config["receiver"], "注册码通知", f"恭喜！您的注册码 {inv_code} 可用。")

            rx += 1 #验证码通过了才能进行下一个循环
        except Exception as e:
            print(f"处理第 {rx + 1} 行时发生错误: {e}")
            messagebox.showinfo("错误", f"处理第 {rx + 1} 行时发生错误: {e}")
            write_sheet.write(rx, 2, 'no')
            rx += 1 

    updated_excel_path = excel_file_path.replace(".xls", "_updated.xls")
    write_book.save(updated_excel_path)
    driver.quit()
    log_file.close()
    messagebox.showinfo("完成", "处理完成，更新的文件已保存至 " + updated_excel_path)

#选择文件路径
def select_excel_file():
    global excel_file_path, browse_button
    browse_button.config(state='disabled', text='正在选择...')
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
    if filepath:
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, filepath)
        excel_file_path = filepath
        config_data = load_config()
        config_data.update({"excel_file_path": excel_file_path}) 
        save_config(config_data) #配置文件存表格路径配置
    browse_button.config(state='normal', text='浏览')

#配置邮箱按钮的事件逻辑：把输入框内容放全局变量
def configure_email():
    email_window = tk.Toplevel(app)
    email_window.title("配置邮箱服务")

    tk.Label(email_window, text="邮箱账户:").grid(row=0, column=0)
    email_entry = tk.Entry(email_window)
    email_entry.grid(row=0, column=1)

    tk.Label(email_window, text="密码:").grid(row=1, column=0)
    password_entry = tk.Entry(email_window, show="*")
    password_entry.grid(row=1, column=1)

    tk.Label(email_window, text="接收邮箱:").grid(row=2, column=0)
    receiver_entry = tk.Entry(email_window)
    receiver_entry.grid(row=2, column=1)

    #配置文件存邮箱配置
    def save_email_info():
        global email_config
        email_config = {
            "sender": email_entry.get(),
            "password": password_entry.get(),
            "receiver": receiver_entry.get()
        }
        config_data = load_config()
        config_data.update({"email_config": email_config})
        save_config(config_data)
        messagebox.showinfo("信息", "邮箱配置已保存")
        email_window.destroy()
    #触发赋值的按钮
    tk.Button(email_window, text="保存配置", command=save_email_info).grid(row=3, column=1)

#--------------------------------可以看见的GUI界面布局设置--------------------------------------

file_path_entry = tk.Entry(app, width=50)
file_path_entry.grid(row=0, column=1)

browse_button = tk.Button(app, text="浏览", command=select_excel_file)
browse_button.grid(row=0, column=2)

tk.Button(app, text="配置邮箱服务", command=configure_email).grid(row=1, columnspan=3)

email_var = tk.BooleanVar(value=False)
tk.Checkbutton(app, text="发送邮箱", variable=email_var).grid(row=2, columnspan=3)

start_row_entry = tk.Entry(app)
start_row_entry.grid(row=3, column=1)
tk.Button(app, text="确定", command=lambda: start_row_entry.config(state='disabled')).grid(row=3, column=2)
tk.Button(app, text="修改", command=lambda: start_row_entry.config(state='normal')).grid(row=3, column=3)

tk.Button(app, text="启动", command=process).grid(row=4, column=1)

log_text = tk.Text(app, height=30, width=75)
log_text.grid(row=5, columnspan=3)

initialize_app()
app.mainloop()
