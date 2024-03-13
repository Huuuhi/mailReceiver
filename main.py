# 这是一个示例 Python 脚本。
import imaplib
import email
import os
from email.header import decode_header

outlook_address="outlook.office365.com"
sina_address="imap.sina.com"
# 连接到邮件服务器
mail = imaplib.IMAP4_SSL('imap.sina.com') #邮件服务器地址
mail.login('邮件地址->改', '邮件授权码->改')#邮箱地址，授权码
mail.select('INBOX')  # 选择收件箱或其他文件夹

# 搜索特定条件的邮件
# message = '第七章'
# print(message.encode('utf-8'))
# 获取所有邮件的UID列表


# result, data = mail.uid('search', charset='UTF-8',
#                                'SUBJECT', f'"{keyword.decode()}*"')
# status, data = mail.search(None, 'SINCE 05-May-2023)'.encode('imap-utf7'))

status, data = mail.uid('search', None, 'ALl')
email_uids = data[0].split()


# 遍历每个邮件并下载附件
for email_uid in email_uids:
    status, data = mail.uid('fetch', email_uid, '(RFC822)')
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
    if '第九章' or '第9章' in subject:
        # 检查邮件标题是否包含目标关键字
        for part in msg.walk():
            if part.get_content_type() == 'application/octet-stream' and part.get_filename():
                # 检查附件标题中是否包含"第七章"三个字
                filename = part.get_filename()
                decoded_filename = decode_header(filename)[0][0]
                if isinstance(decoded_filename, bytes):
                    try:
                        decoded_filename = decoded_filename.decode('gb18030')
                    except UnicodeDecodeError:
                        decoded_filename = decoded_filename.decode('utf-8')
                # 保存附件
                if '第9章' in decoded_filename or '第九章' in decoded_filename:
                    with open(decoded_filename, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    print(f"已导出附件: {decoded_filename}")

# 关闭连接
mail.logout()
