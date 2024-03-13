import poplib
import email
import datetime
import os
import shutil

# 连接到 POP3 服务器
# pop_server = poplib.POP3_SSL('outlook.office365.com',port=995)
# pop_server.user('@qq.com')
# pop_server.pass_('')
pop_server = poplib.POP3_SSL('pop.sina.com')
pop_server.user('xxx@sina.com')
pop_server.pass_('xxx')



# 创建保存附件的文件夹
attachment_dir = './attachments'  # 指定附件保存的文件夹名称
os.makedirs(attachment_dir, exist_ok=True)

# 获取邮件数量和大小
num_emails = len(pop_server.list()[1])
email_ids = [i + 1 for i in range(num_emails)]

# 遍历每个邮件并处理
for email_id in email_ids:
    _, email_data, _ = pop_server.retr(email_id)
    raw_email = b'\r\n'.join(email_data)
    msg = email.message_from_bytes(raw_email)

    # 遍历每个消息部分并检查是否是附件
    for part in msg.walk():
        if part.get_content_disposition() == 'attachment':
            attachment_filename = part.get_filename()
            print(attachment_filename)
            if attachment_filename:
                try:
                    attachment_filename = attachment_filename.encode('iso-8859-1').decode('gb18030')
                except UnicodeDecodeError:
                    continue

                print(attachment_filename)
                print('...')

                if '第八章' in attachment_filename or '第8章' in attachment_filename:
                    # 下载附件
                    attachment_data = part.get_payload(decode=True)

                    # 获取附件的完整保存路径
                    attachment_path = os.path.join(attachment_dir, attachment_filename)

                    # 保存附件到本地
                    with open(attachment_path, 'wb') as file:
                        file.write(attachment_data)

# 关闭连接
pop_server.quit()
