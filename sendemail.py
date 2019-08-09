#回复邮件
from email.mime.text import MIMEText
import smtplib
from email.header import Header
from email.utils import parseaddr,formataddr#设置编码格式
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
import openpyxl
import pandas as pd
import time
time.strftime("%Y/%m/%d", time.localtime())

data_path = r"C:\Users\Administrator\Desktop\0807zongbiao.xlsx"
data = pd.read_excel(data_path)

reply_root_path = r"C:\Users\Administrator\Desktop\ttt"

# for email_id,value in data.groupby("发件人邮箱"):
#     current_time = time.strftime("%Y%m%d", time.localtime())
#     reply_path = reply_root_path + "\回复附件#" +  current_time+ "#" + email_id + '.xlsx'
#     value.reset_index(drop=True).to_excel(reply_path)

#打开excel文件,读取邮箱地址和审核意见列
# wb=openpyxl.load_workbook(r'C:\Users\Administrator\Desktop\0801zongbiao.xlsx')
# sheet=wb.get_active_sheet()
# to_addrs=list(sheet.columns)[9]
# body=list(sheet.columns)[9]
#for item in list(sheet.columns)[10]:
    #print(item.value)

#将用户名编码设置成UTF-8
def _format_addr(s):
    name,addr=parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(),addr))


#定义email的地址，口令和SMTP服务器地址
from_addr='zgcindex_zh@163.com'
password=input('请输入发送邮箱的密码：')#qwertyuiop1
smtp_server='smtp.163.com'
# to_addr='sx_wy123@163.com'



#定义邮件本身内容
# msg=MIMEMultipart()
# msg['From']=_format_addr('发送者的ReedSun<%s>'%from_addr)
# msg['To']=_format_addr('接收者的ReedSun<%s>'%to_addr)
# msg['Subject']=Header('回复','utf-8').encode()

#定义邮件正文
# msg.attach(MIMEText('使用python发来的邮件3','plain','utf-8'))

#加附件
for email_id,value in data.groupby("发件人邮箱"):
    current_time = time.strftime("%Y%m%d", time.localtime())
    reply_path = reply_root_path + "\回复附件#" +  current_time+ "#" + email_id + '.xlsx'#文件名
    value.reset_index(drop=True).to_excel(reply_path)
    to_addr=email_id
    msg=MIMEMultipart()
    msg['From']=_format_addr('发送者的ReedSun<%s>'%from_addr)
    msg['To']=_format_addr('接收者的ReedSun<%s>'%to_addr)
    msg['Subject']=Header('总后审核意见','utf-8').encode()
    msg.attach(MIMEText('总后新品申请审核意见，详情请见附件！注：有多封申请邮件的公司，不再分别回复，统一一并回复到附件中！请大家一定按要求格式申请，否则可能会无法正常申请上架，谢谢配合~','plain','utf-8'))
    
    with open(reply_path,'rb') as f:
         # 设置附件的MIME和文件名，这里是jpg类型,可以换png或其他类型:
        mime=MIMEBase("\回复附件#" + current_time+ "#" + email_id, 'xlsx', filename="\回复附件#" +  current_time+ "#" + email_id + '.xlsx')
        mime.add_header('Content-Disposition','attchment',filename=reply_path)
        mime.add_header('Content-ID','<0>')
        mime.add_header('X-Attachment-ID','0')
        mime.set_payload(f.read())
        encoders.encode_base64(mime)
        msg.attach(mime)

    #定义发送文件
    server=smtplib.SMTP_SSL(smtp_server,465)
    server.set_debuglevel(1)
    server.login(from_addr,password)
    server.sendmail(from_addr,to_addr,msg.as_string())

server.quit()