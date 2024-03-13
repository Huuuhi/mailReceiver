from imapclient import IMAPClient
import email
import os

# 邮件服务器的登录凭据
email = 'xxx@sina.com'
password = 'xxx' #授权码

# 连接到IMAP服务器
with IMAPClient('imap.sina.com') as imap_server:
    # 登录到邮箱账号
    imap_server.login(email, password)

    # 选择邮箱文件夹
    imap_server.select_folder('INBOX')

    # 搜索符合条件的邮件
    search_criteria1 = b'SUBJECT "第八"'
    search_criteria2 = b'SUBJECT "第8"'
    messages = imap_server.search([search_criteria1, search_criteria2], 'OR')

    # 遍历邮件
    for message_id in messages:
        # 获取邮件内容
        raw_message = imap_server.fetch([message_id], ['RFC822'])[message_id][b'RFC822']
        msg = email.message_from_bytes(raw_message)

        # 检查是否有附件
        if msg.get_content_maintype() == 'multipart':
            for part in msg.iter_attachments():
                # 获取附件文件名
                filename = part.get_filename()
                if filename:
                    # 下载附件到本地
                    filepath = os.path.join('attachments', filename)
                    with open(filepath, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    print(f"已下载附件: {filename}")

# 登出邮箱账号
imap_server.logout()
