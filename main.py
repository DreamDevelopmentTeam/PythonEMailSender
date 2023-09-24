# 导入所需的模块
import smtplib
import email
import tkinter as tk
import json

# 定义一个函数，用于发送邮件
def send_mail():
    # 获取用户输入的信息
    server = server_entry.get()
    port = port_entry.get()
    ssl = ssl_var.get()
    sender = sender_entry.get()
    password = password_entry.get()
    receiver = receiver_entry.get()
    cc = cc_entry.get()
    bcc = bcc_entry.get()
    subject = subject_entry.get()
    content = content_text.get("1.0", "end")
    html = html_var.get()

    # 创建一个邮件对象
    msg = email.message.EmailMessage()

    # 设置邮件的基本信息
    msg["Subject"] = subject # 邮件主题
    msg["From"] = sender # 发件人
    msg["To"] = receiver # 收件人
    if cc: # 如果有抄送人
        msg["Cc"] = cc # 抄送人
    if bcc: # 如果有密送人
        msg["Bcc"] = bcc # 密送人

    # 设置邮件的内容
    if html: # 如果是HTML格式的邮件
        msg.set_content(content, subtype="html") # 设置内容为HTML格式
    else: # 如果是文本格式的邮件
        msg.set_content(content) # 设置内容为文本格式

    # 尝试连接服务器并发送邮件
    try:
        if ssl: # 如果启用了SSL加密
            smtp = smtplib.SMTP_SSL(server, port) # 创建一个SSL加密的SMTP对象
        else: # 如果没有启用SSL加密
            smtp = smtplib.SMTP(server, port) # 创建一个普通的SMTP对象

        smtp.login(sender, password) # 登录发件人的账号和密码
        smtp.send_message(msg) # 发送邮件对象
        smtp.quit() # 退出SMTP连接

        status_label["text"] = "邮件发送成功！" # 显示发送成功的提示信息

        # 保存配置信息到json文件中，不保存密码
        config = {
            "server": server,
            "port": port,
            "ssl": ssl,
            "sender": sender,
            "receiver": receiver,
            "cc": cc,
            "bcc": bcc,
            "html": html,
        }
        with open("config.json", "w") as f:
            json.dump(config, f)

    except Exception as e: # 如果出现异常
        status_label["text"] = f"邮件发送失败！错误信息：{e}" # 显示发送失败的提示信息


# 创建一个Tkinter窗口对象
window = tk.Tk()

# 设置窗口的标题和大小
window.title("邮件发送程序")
window.geometry("600x600")

# 创建一个标签，显示程序的标题
title_label = tk.Label(window, text="邮件发送程序", font=("Arial", 20))
title_label.pack()

# 创建一个标签，显示服务器地址的提示信息
server_label = tk.Label(window, text="请输入服务器地址：")
server_label.pack()

# 创建一个输入框，用于输入服务器地址
server_entry = tk.Entry(window)
server_entry.pack()

# 创建一个标签，显示端口号的提示信息
port_label = tk.Label(window, text="请输入端口号：")
port_label.pack()

# 创建一个输入框，用于输入端口号
port_entry = tk.Entry(window)
port_entry.pack()

# 创建一个复选框，用于选择是否启用SSL加密
ssl_var = tk.IntVar() # 创建一个整数变量，用于存储复选框的状态
ssl_checkbutton = tk.Checkbutton(window, text="启用SSL加密", variable=ssl_var)
ssl_checkbutton.pack()

# 创建一个标签，显示发件人的提示信息
sender_label = tk.Label(window, text="请输入发件人（账号）：")
sender_label.pack()

# 创建一个输入框，用于输入发件人
sender_entry = tk.Entry(window)
sender_entry.pack()

# 创建一个标签，显示密码的提示信息
password_label = tk.Label(window, text="请输入密码：")
password_label.pack()

# 创建一个输入框，用于输入密码，显示为星号
password_entry = tk.Entry(window, show="*")
password_entry.pack()

# 创建一个标签，显示收件人的提示信息
receiver_label = tk.Label(window, text="请输入收件人：")
receiver_label.pack()

# 创建一个输入框，用于输入收件人
receiver_entry = tk.Entry(window)
receiver_entry.pack()

# 创建一个标签，显示抄送人的提示信息
cc_label = tk.Label(window, text="请输入抄送人（可选）：")
cc_label.pack()

# 创建一个输入框，用于输入抄送人
cc_entry = tk.Entry(window)
cc_entry.pack()

# 创建一个标签，显示密送人的提示信息
bcc_label = tk.Label(window, text="请输入密送人（可选）：")
bcc_label.pack()

# 创建一个输入框，用于输入密送人
bcc_entry = tk.Entry(window)
bcc_entry.pack()

# 创建一个标签，显示邮件主题的提示信息
subject_label = tk.Label(window, text="请输入邮件主题：")
subject_label.pack()

# 创建一个输入框，用于输入邮件主题
subject_entry = tk.Entry(window)
subject_entry.pack()

# 创建一个标签，显示邮件内容的提示信息
content_label = tk.Label(window, text="请输入邮件内容：")
content_label.pack()

# 创建一个文本框，用于输入邮件内容
content_text = tk.Text(window)
content_text.pack()

# 创建一个复选框，用于选择是否发送HTML格式的邮件
html_var = tk.IntVar() # 创建一个整数变量，用于存储复选框的状态
html_checkbutton = tk.Checkbutton(window, text="发送HTML格式的邮件", variable=html_var)
html_checkbutton.pack()

# 创建一个按钮，用于发送邮件
send_button = tk.Button(window, text="发送邮件", command=send_mail)
send_button.pack()

# 创建一个标签，用于显示发送状态的提示信息
status_label = tk.Label(window, text="")
status_label.pack()

# 尝试从json文件中读取配置信息，并填充到对应的输入框或复选框中
try:
    with open("config.json", "r") as f:
        config = json.load(f) # 读取json文件中的数据

    # 填充配置信息到对应的输入框或复选框中
    server_entry.insert(0, config["server"])
    port_entry.insert(0, config["port"])
    ssl_var.set(config["ssl"])
    sender_entry.insert(0, config["sender"])
    receiver_entry.insert(0, config["receiver"])
    cc_entry.insert(0, config["cc"])
    bcc_entry.insert(0, config["bcc"])
    html_var.set(config["html"])

except Exception as e: # 如果出现异常
    status_label["text"] = f"配置文件读取失败！错误信息：{e}" # 显示读取失败的提示信息

# 进入Tkinter主循环
window.mainloop()
