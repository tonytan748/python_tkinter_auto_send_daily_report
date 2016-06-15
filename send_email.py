# -*- coding: utf-8 -*-
import os
from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.utils import parseaddr, formataddr

import smtplib

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr(( \
        Header(name, 'utf-8').encode(), \
        addr.encode('utf-8') if isinstance(addr, unicode) else addr))

class SendEmail:
    def __init__(self,msg):
        self.msg_from = msg['from']
        self.msg_to = msg['to']
        # if isinstance(msg['to'],list) else [msg['to']]
        self.msg_cc = msg['cc']
        self.msg_password = msg['password']
        self.msg_server = msg['server']
        self.msg_filepath = msg['file_path'] if msg['file_path'] else None

        # self.msg_subject = msg['subject']
        # self.msg_context = msg['context']
        self.msg_subject = u'{}-{}-{}-个人绩效日报'.format(msg['department'], msg['name'], msg['date'])
        self.msg_context = u'''
            你好，

            日报如附件，请审阅。

            {}

            {}
        '''.format(msg['name'], msg['date'])

    def define_msg(self):
        msg = MIMEMultipart()
        msg['From'] = self.msg_from
        msg['To'] = self.msg_to
        msg['Subject'] = Header(self.msg_subject, 'utf-8').encode()
        msg['Cc'] = self.msg_cc

        # add MIMEText:
        msg.attach(MIMEText(self.msg_context, 'plain', 'utf-8'))

        # add file:
        if self.msg_filepath:
            try:
                file_name = os.path.abspath(self.msg_filepath)
                this_file_name = os.path.basename(self.msg_filepath)
                this_file_name = this_file_name.encode("gbk")
                with open(file_name, 'rb') as f:
                    mime = MIMEBase('application', 'msword')
                    mime.add_header('Content-Disposition', 'attachment', filename=this_file_name)
                    mime.set_payload(f.read())
                    encoders.encode_base64(mime)
                    msg.attach(mime)
            except Exception as e:
                raise e
        return msg

    def send_email(self):
        try:
            msg = self.define_msg()

            server = smtplib.SMTP(self.msg_server, 25)
            server.set_debuglevel(1)
            server.login(self.msg_from, self.msg_password)
            server.sendmail(self.msg_from, [self.msg_to] + self.msg_cc.split(","), msg.as_string())
            server.quit()
            return True
        except Exception as e:
            print e
            return False

if __name__=="__main__":
    file_path = os.path.join(os.getcwd(),"Daily_Report", u"大数据实验室-谭勇-20160614-个人绩效日报.docx")
    msg = {
    'from':"tany@hzhz.co",
    'password':"Ty131514",
    'server':"smtp.hzhz.co",
    "to":"tany@hzhz.co",
    "cc":"violin748@hotmail.com,tonytan748@gmail.com",
    "department":u"大数据实验室",
    "name":u"谭咏麟",
    'date':"2016-08-08",
    'file_path':file_path
    }

    se = SendEmail(msg)
    res = se.send_email()
    if res:
        print "success"
    else:
        print "fail"
