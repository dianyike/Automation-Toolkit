import smtplib
import schedule
import time
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime
import os
import sys
import pandas as pd

# 設置控制台輸出編碼
sys.stdout.reconfigure(encoding='utf-8')

# 配置日誌
logging.basicConfig(
    filename='email_sender.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

def read_recipients_from_file(file_path):
    """
    從文件讀取收件人列表
    支持 .xlsx, .xls, .txt 格式
    """
    try:
        if file_path.endswith(('.xlsx', '.xls')):
            # 讀取Excel文件
            df = pd.read_excel(file_path)
            # 假設郵箱在第一列，如果列名不同請修改
            return df.iloc[:, 0].tolist()
        elif file_path.endswith('.txt'):
            # 讀取txt文件
            with open(file_path, 'r', encoding='utf-8') as f:
                return [line.strip() for line in f if line.strip()]
        else:
            logging.error(f"不支持的文件格式: {file_path}")
            return []
    except Exception as e:
        logging.error(f"讀取收件人列表時出錯: {str(e)}")
        return []

class EmailSender:
    def __init__(self, smtp_server, smtp_port, sender_email, password):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.sender_email = sender_email
        self.password = password

    def send_email(self, recipient_email, subject, body, attachments=None):
        try:
            # 創建信箱對象
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = recipient_email
            msg['Subject'] = subject

            # 添加信箱正文
            msg.attach(MIMEText(body, 'plain'))

            # 添加附件
            if attachments:
                for file_path in attachments:
                    if os.path.exists(file_path):
                        with open(file_path, 'rb') as f:
                            attachment = MIMEApplication(f.read(), _subtype=os.path.splitext(file_path)[1][1:])
                            attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
                            msg.attach(attachment)
                    else:
                        logging.error(f"附件不存在: {file_path}")

            # 連接MTP伺服器并發送信箱
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, self.password)
                server.send_message(msg)

            logging.info(f"信箱成功發送給 {recipient_email}")
            return True

        except Exception as e:
            logging.error(f"發送信箱時出錯: {str(e)}")
            return False

def send_email_with_options(sender, recipients, subject, body, attachments=None, schedule_time=None):
    """
    發送郵件的整合函數
    :param sender: EmailSender實例
    :param recipients: 收件人列表
    :param subject: 郵件主題
    :param body: 郵件內容
    :param attachments: 附件列表
    :param schedule_time: 定時發送時間（格式：'HH:MM'），如果為None則立即發送
    """
    def send_to_all():
        success_count = 0
        fail_count = 0
        for recipient in recipients:
            print(f"正在發送郵件給 {recipient}...")
            if sender.send_email(recipient, subject, body, attachments):
                success_count += 1
                print(f"成功發送給 {recipient}")
            else:
                fail_count += 1
                print(f"發送給 {recipient} 失敗")
        print(f"\n發送完成！成功：{success_count}，失敗：{fail_count}")

    if schedule_time:
        # 定時發送
        schedule.every().day.at(schedule_time).do(send_to_all)
        print(f"已設置定時發送，將在每天 {schedule_time} 發送郵件")
        
        # 運行定時任務
        while True:
            schedule.run_pending()
            time.sleep(60)
    else:
        # 立即發送
        send_to_all()

def main():
    # 配置信箱伺服器信息
    smtp_server = "smtp.gmail.com"  # 修改為正確的SMTP伺服器
    smtp_port = 587
    sender_email = "example@gmail.com"  # 替換為您的信箱
    password = "123456789"  # 替換為您的密碼或應用專用密碼

    # 創建EmailSender實例
    email_sender = EmailSender(smtp_server, smtp_port, sender_email, password)

    # 讀取收件人列表
    print("請選擇收件人列表文件格式：")
    print("1. Excel文件 (.xlsx, .xls)")
    print("2. 文本文件 (.txt)")
    file_choice = input("請輸入選項（1或2）：")
    
    if file_choice == "1":
        file_path = input("請輸入Excel文件路徑：")
    elif file_choice == "2":
        file_path = input("請輸入txt文件路徑：")
    else:
        print("無效的選項！")
        return

    recipients = read_recipients_from_file(file_path)
    if not recipients:
        print("未能讀取到有效的收件人列表！")
        return

    print(f"成功讀取到 {len(recipients)} 個收件人")

    # 設置信箱內容
    subject = input("請輸入郵件主題：")
    body = input("請輸入郵件內容：")
    
    # 添加附件
    attachments = []
    while True:
        attachment = input("請輸入附件路徑（直接按Enter結束添加）：")
        if not attachment:
            break
        attachments.append(attachment)

    # 用戶可以選擇發送方式
    print("\n請選擇發送方式：")
    print("1. 立即發送")
    print("2. 定時發送")
    choice = input("請輸入選項（1或2）：")

    if choice == "1":
        # 立即發送
        send_email_with_options(email_sender, recipients, subject, body, attachments)
    elif choice == "2":
        # 定時發送
        schedule_time = input("請輸入定時發送時間（格式：HH:MM，例如 09:00）：")
        send_email_with_options(email_sender, recipients, subject, body, attachments, schedule_time)
    else:
        print("無效的選項！")

if __name__ == "__main__":
    main() 