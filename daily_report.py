# -*-coding: utf-8 -*-
import sqlite3

import poplib
import email
import email.utils
import datetime
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from email.utils import parsedate_tz
from email.utils import mktime_tz
from email import message_from_string
from bs4 import BeautifulSoup
import chardet

from email.Iterators import typed_subpart_iterator

from email.header import decode_header
def getheader(header_text, default="gb2312"):
    """Decode the specified header"""

    headers = decode_header(header_text)
    header_sections = [unicode(text, charset or default)
                       for text, charset in headers]
    return u"".join(header_sections)

def write_msg(msg):
    target = open('output', 'a')
    target.write(msg.encode('utf8'))
    target.close()
    
def get_charset(message, default="gb2312"):
    """Get the message charset"""

    if message.get_content_charset():
        return message.get_content_charset()

    if message.get_charset():
        return message.get_charset()

    return default

def get_body(message):
    """Get the body of the email message"""

    if message.is_multipart():
        #get the plain text version only
        text_parts = [part
                      for part in typed_subpart_iterator(message,
                                                         'text',
                                                         'plain')]
        body = []
        for part in text_parts:

            charset = get_charset(part, get_charset(message))
            body.append(unicode(part.get_payload(decode=True),
                                charset,
                                "replace"))

        return u"\n".join(body).strip()

    else: # if it is not multipart, the payload will be a string
          # representing the message body
        body = unicode(message.get_payload(decode=True),
                       get_charset(message),
                       "replace")
        return body.strip()
    

email = 'xxx@xxx.com'
password = 'xxx'
pop3_server = 'xxx'

server = poplib.POP3(pop3_server)

server.user(email)
server.pass_(password)

#print('Messages: %s. Size: %s' % server.stat())
# list()返回所有邮件的编号:
resp, mails, octets = server.list()
# 可以查看返回的列表类似['1 82923', '2 2184', ...]
#print(mails)
# 获取最新一封邮件, 注意索引号从1开始:
index = len(mails)-1

dirName = 'c:/users/admin/onedrive/daily_report/'

conn = sqlite3.connect(dirName + 'data.db')
c = conn.cursor()
conn.text_factory = str

for index in range(1, len(mails)):
    write_msg('\nindex----->' + str(index) + '\n')
    resp, lines, octets = server.retr(index)
    # lines存储了邮件的原始文本的每一行,
    # 可以获得整个邮件的原始文本:
    msg_content = '\r\n'.join(lines)

    message = message_from_string(msg_content)

    print 'from===>', message['from']
    print getheader(message['from'])
    date_str = message['date']
    if date_str:
        date_tuple = parsedate_tz(date_str)
        if date_tuple:
            date=datetime.datetime.fromtimestamp(mktime_tz(date_tuple))
            print 'Date:', date

    
    write_msg('\nFrom:\n' + getheader(message['from']))
    
    msg_txt = get_body(message)
    write_msg(msg_txt);
    
#    c.execute('insert into DailyReport values(?,?,?,?)', (from_e, from_e, e_content, date_s))
#    conn.commit()

#c.execute('select * from DailyReport')
print '----------------'
#print c.fetchone()
conn.close()


# 关闭连接:
server.quit()
