# coding:utf-8   #强制使用utf-8编码格式

import smtplib  # 加载smtplib模块
from email.utils import formataddr
from email.mime.text import MIMEText
from email.header import Header
from email import encoders
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase

# 一，
# 代表你发送邮箱的用户名
sender = '********@163.com'

# 对方邮箱的名字
receiver = '**********@qq.com'

# 发送标题
subject = '学习使我快乐'

smtpserver = 'smtp.163.com'

# 邮箱对应用户名密码，密码不是邮箱密码是你开通smtp时设定的密码
username = '********@163.com'
password = '*******'

# 邮件内容，记得一定要用plain传入
msg = MIMEText('学习使我快乐', 'plain', 'utf-8')  # 中文需参数‘utf-8'，单字节字符不需要
msg['Subject'] = Header(subject, 'utf-8')

# 自己邮箱用户名
msg['From'] = sender

# 发送给哪些邮箱
msg['To'] = receiver

# 发送邮箱方法被调用传入参数
smtp = smtplib.SMTP()
smtp.connect(smtpserver)
smtp.login(username, password)
smtp.sendmail(sender, receiver, msg.as_string())
smtp.quit()

print("================================================")


# 二
my_sender = '**********@163.com'  # 发件人邮箱账号，为了后面易于维护，所以写成了变量
my_user = '**********@qq.com'  # 收件人邮箱账号，为了后面易于维护，所以写成了变量


def mail():
    ret = True
    try:
        msg = MIMEText('填写邮件内容', 'plain', 'utf-8')
        msg['From'] = formataddr(["张****", my_sender])  # 括号里的对应发件人邮箱昵称、发件人邮箱账号
        msg['To'] = formataddr(["张三", my_user])  # 括号里的对应收件人邮箱昵称、收件人邮箱账号
        msg['Subject'] = "主题测试邮件"  # 邮件的主题，也可以说是标题

        server = smtplib.SMTP("smtp.163.com", 25)  # 发件人邮箱中的SMTP服务器，端口是25
        server.login(my_sender, "****")  # 括号中对应的是发件人邮箱账号、邮箱密码(这里的密码是授权码密钥)
        server.sendmail(my_sender, [my_user, ], msg.as_string())  # 括号中对应的是发件人邮箱账号、收件人邮箱账号、发送邮件
        server.quit()  # 这句是关闭连接的意思

    except Exception:  # 如果try中的语句没有执行，则会执行下面的ret=False
        ret = False

    return ret


ret = mail()
if ret:
    print("ok")  # 如果发送成功则会返回ok，稍等20秒左右就可以收到邮件
else:
    print("filed")  # 如果发送失败则会返回filed



print("================================================")



# 三
sender = 'xxxxx@163.com'
receiver = list()#接收者列表
receiver.append('yyyyy@163.com')
copyReceive = list()#抄送者列表
copyReceive.append(sender)#将发件人添加到抄送列表
username = 'xxxxx@163.com'#发件人邮箱账号
password = '****'#授权密码
mailall=MIMEMultipart()
mailall['Subject'] = "测试邮件主题" #记住一定要设置，并且要稍微正式点
mailall['From'] = sender #发件人邮箱
mailall['To'] = ';'.join(receiver) #收件人邮箱,不同收件人邮箱之间用;分割
mailall['CC'] = ';'.join(copyReceive)  #抄送邮箱
mailcontent = '测试邮件正文'
mailall.attach(MIMEText(mailcontent, 'plain', 'utf-8'))
mailAttach = '测试邮件附件内容'
contype = 'application/octet-stream'
maintype, subtype = contype.split('/', 1)
filename = '附加文件.txt'#附件文件所在路径
attfile = MIMEBase(maintype, subtype)
attfile.set_payload(open(filename, 'rb').read())
attfile.add_header('Content-Disposition', 'attachment',filename=('utf-8', '', filename))#必须加上第三个参数，用于格式化输出
mailall.attach(attfile)
fullmailtext = mailall.as_string()
smtp = smtplib.SMTP()
smtp.connect('smtp.163.com')
smtp.login(username, password)
smtp.sendmail(sender, receiver+copyReceive, fullmailtext)#发送的时候需要将收件人和抄送者全部添加到函数第二个参数里
smtp.quit()


print("================================================")


# 四.
# 用Python 在 AWS 调用邮件接口：
import smtplib
import email.utils
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

class SendEmail:
    def send_email(self, content):
        SENDER = '*********@easyaca.cn'  # 发送邮箱的用户名
        SENDERNAME = '发件人姓名可修改'  # 发件人姓名
        RECIPIENT = ['*******@foxmail.com', "*********@qq.com"]  # 发送到的邮件，可是list, 但注意时msg['To']值必须是字符串用,隔开
        USERNAME_SMTP = "AKIAIR32MMJV4XGRYENA"  # 带有邮件权限的 IAM 帐号
        PASSWORD_SMTP = "Ai5tfn5Lsm1Pi/9PJovJr9PxnrSltb56Wo/RmdeAhsHb"  # 带有邮件权限的 IAM 密码
        HOST = "email-smtp.us-east-1.amazonaws.com"  # Amazon SES SMTP 终端节点
        PORT = 25  # SMTP 客户端连接到STARTTLS 端口 25、587 或 2587 上的 Amazon SES SMTP 终端节点
        SUBJECT = '主题测试使用可修改'  # 主题

        # 测试邮件正文
        BODY_TEXT = ("Amazon SES Test\r\n"
                     "This email was sent through the Amazon SES SMTP "
                     )

        BODY_HTML = """<html>
        <head></head>
        <body>
          <h1>5th International Conference on Civil Engineering</h1>
          <h1>(ICCE2018)</h1>
          <h1>Call for Papers---EI Indexed</h1>
          <h1>Dear Prof. \{UserName\},</h1>
          <h4>
          It is our great pleasure to invite you and submit your newest research to 5th International Conference on
          Civil Engineering (ICCE2018) on Dec. 20-21, 2018, which will be held in Nanchang, one of the most famous
          national historic and cultural cities in China.
          </h4>
          </br>
          <h4>
          ICCE2018 is organized by Nanchang Institute of Technology, Civil Engineering Academy of Jiangxi Province
          and Hubei Zhongke Institute of Geology and Environment Technology, co-organized by journal of Rock and
          Soil Mechanics, Key Laboratory for Safety of Water Conservancy and Civil Engineering Infrastructure in
           Jiangxi Province, Jiangxi Provincial Engineering Research Center of Special Reinforcement and Safety
           Monitoring Technology in Hydraulic & Civil Engineering.
          </h4>
          </br>
          <ul>
          <li><h5>1. Structural Engineering;</h5></li>
          <li><h5>2. Construction Engineering;</h5></li>
          <li><h5>3. Geotechnical Engineering;</h5></li>
          </ul>
        </body>
        </html>
        """

        msg = MIMEMultipart('alternative')
        msg['Subject'] = SUBJECT
        msg['From'] = email.utils.formataddr((SENDERNAME, SENDER))  # 发件人邮箱
        # msg['To'] = RECIPIENT  # 收件人邮箱
        msg['To'] = ','.join(RECIPIENT)  # 收件人邮箱,不同收件人邮箱之间用,分割,再组合为字符串

        # 邮件内容，记得一定要用plain传入
        # part1 = MIMEText(BODY_TEXT, 'plain')
        part2 = MIMEText(BODY_HTML, 'html')
        # msg.attach(part1)
        msg.attach(part2)

        try:
            server = smtplib.SMTP(HOST, PORT)
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(USERNAME_SMTP, PASSWORD_SMTP)

            # 发送的时候需要将收件人和抄送者全部添加到函数第二个参数里
            server.sendmail(SENDER, RECIPIENT, msg.as_string())
            server.close()
        except Exception as e:
            print("Error: ", e)
        else:
            print("Email sent ok !")


if __name__ == "__main__":
    pass
    # sender = SendEmail()
    # sender.send_email("发送邮件测试内容99999999999999999999999")
