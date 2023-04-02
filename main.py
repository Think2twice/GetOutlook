import sys
import tkinter as tk
import win32com.client

def get_latest_email_subject():
    # 创建Outlook对象
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # 访问收件箱（Inbox），注：收件箱的索引为6
    inbox = outlook.GetDefaultFolder(6)

    # 获取收件箱中的所有邮件
    messages = inbox.Items

    # 按接收时间对邮件进行排序，最新的邮件在最前面
    messages.Sort("[ReceivedTime]", True)

    # 获取最新邮件
    latest_email = messages.GetFirst()

    # 返回邮件标题
    return latest_email.Subject

def on_button_click(text_widget):
    # 获取最新邮件标题
    subject = get_latest_email_subject()

    # 更新文本视图的内容
    text_widget.delete(1.0, tk.END)
    text_widget.insert(tk.END, subject)

# 创建主窗口
root = tk.Tk()
root.title("获取邮件内容")

# 设置自定义图标
if getattr(sys, 'frozen', False):
    # 如果是打包状态，使用exe文件中的嵌入图标
    root.iconbitmap(sys.executable)
else:
    # 如果不是打包状态，则使用外部图标文件
    root.iconbitmap('title.ico')# 将'custom_icon.ico'替换为您自己的图标文件名

# 创建文本视图
text_widget = tk.Text(root, wrap=tk.WORD, width=50, height=10)
text_widget.pack()

# 创建按钮
button = tk.Button(root, text="获取邮件内容", command=lambda: on_button_click(text_widget))
button.pack()

# 运行GUI事件循环
root.mainloop()
