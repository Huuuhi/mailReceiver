# -*- coding:UTF-8 -*-
import imaplib
import email
import os
import ssl
from email.header import decode_header

outlook_address = "outlook.office365.com"

# Load system's trusted SSL certificates

# 连接到邮件服务器
# mail = imaplib.IMAP4_SSL(outlook_address, port=993)
mail = imaplib.IMAP4_SSL(outlook_address,port=993)
mail.login('xxx', 'xxx')

code, mailboxes = mail.list()
for mailbox in mailboxes:
    print(mailbox.decode("utf-8"))

mail.select('INBOX')  # 选择收件箱或其他文件夹

# 搜索特定条件的邮件
# message = '第七章'
# print(message.encode('utf-8'))
# 获取所有邮件的UID列表
# 转换搜索关键字为字符串并编码为utf-8

# 搜索邮件
status, data = mail.search(None, 'ALL')  # 搜索所有邮件
email_ids = data[0].split()

# 遍历每个邮件并下载附件
for email_id in email_ids:
    status, data = mail.fetch(email_id, '(RFC822)')
    raw_email = data[0][1]
    msg = email.message_from_bytes(raw_email)

    subject = decode_header(msg['Subject'])[0][0]
    if isinstance(subject, bytes):
        if isinstance(subject, bytes):
            try:
                subject = subject.decode('gbk')
            except UnicodeDecodeError:
                subject = subject.decode('utf-8')
        print(subject)

    if '第5次' or '第五次' in subject:
        for part in msg.walk():
            if part.get_content_type() == 'application/octet-stream' and part.get_filename():
                # 检查附件名是否包含"第一次实验"或"第1次实验"这几个字
                filename = part.get_filename()
                decoded_filename = decode_header(filename)[0][0]
                if isinstance(decoded_filename, bytes):
                    try:
                        decoded_filename = decoded_filename.decode('utf-8')
                    except UnicodeDecodeError:
                        decoded_filename = decoded_filename.decode('gb18030')
                if '第五次' in decoded_filename or '第5次' in decoded_filename:
                    # 保存附件
                    with open(decoded_filename, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    print(f"已导出附件: {decoded_filename}")

# 关闭连接
mail.logout()
