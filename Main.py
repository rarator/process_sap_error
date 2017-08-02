# encoding=utf-8
import poplib
import email
import sys
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import time
import os


def process_mail(msg):
    sender = msg.get('From', '')
    if sender ==  'c-joho@suruga-g.co.jp':
        hdr, addr = parseaddr(sender)
        name = decode_str(hdr)
        value = u'%s <%s>' % (name, addr)

        content = msg.get_payload()
        count = len(content)
        for i in range(count):
            if content[i].get_content_type() == 'application/octet-stream':
                header = content[i].get('content-type')
                if '0999' in header:
                    date = header[-18:-10]
                    date_now = time.strftime("%Y%m%d", time.localtime())
                    if date == date_now:
                        startIndex = header.find('name=')
                        filename = header[startIndex + 5:]
                        saveFilename = filename[0:-19] + '.txt'
                        text = content[i].get_payload(decode=True)
                        text = text.decode('utf-8')

                        # 读历史文件,如果存在就不再处理
                        with open('F:\\MailDoneList.txt', 'r') as f:
                            txt = f.read()
                            if saveFilename in txt:
                                return

                        # os.system('net use Z: \\172.27.254.45\\301\\JDAY_FILE\\T')
                        with open('Z:\\' + saveFilename, 'w',encoding='utf-8') as f:
                            # text = '﻿000002325779\t0999\t1\t0\n'
                            # 暂时先这样处理一下，不然结果会多一个0D导致文件不被处理并报错
                            text = text.replace("\r", "")
                            f.write(text)

                        with open('F:\\MailDoneList.txt', 'a') as f:
                            f.write(saveFilename + '\r\n')

                        print(saveFilename + '  done!')



def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def guess_charset(msg):
    # 先从msg对象获取编码:
    charset = msg.get_charset()
    if charset is None:
        # 如果获取不到，再从Content-Type字段获取:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset

if __name__ == '__main__':
    print('--------------------------------------------------------------------------')
    print('Process Begin')
    M = poplib.POP3('mail-sh.suruga-sh.com')
    M.user('rrf')
    M.pass_('rrf2016')

    # 打印有多少封信
    numMessages = len(M.list()[1])
    print('num of messages', numMessages)

    print('Messages: %s. Size: %s' % M.stat())
    # list()返回所有邮件的编号:
    #resp, mails, octets = M.list()
    # 可以查看返回的列表类似['1 82923', '2 2184', ...]
    #print(mails)
    # 获取最新一封邮件, 注意索引号从1开始:
    #index = len(mails)
    #resp, lines, octets = M.retr(index)
    # lines存储了邮件的原始文本的每一行,
    # 可以获得整个邮件的原始文本:
    #msg_content = b'\r\n'.join(lines)

    #msg_content = str(msg_content, encoding = "utf-8")
    # 稍后解析出邮件:
    #msg = Parser().parsestr(msg_content)
    #print_info(msg)

    # 从最老的邮件开始遍历
    for i in range(numMessages):
        resp, lines, octets = M.retr(i + 1)
        # lines存储了邮件的原始文本的每一行,
        # 可以获得整个邮件的原始文本:
        msg_content = b'\r\n'.join(lines)

        msg_content = str(msg_content, encoding="utf-8")
        # 稍后解析出邮件:
        msg = Parser().parsestr(msg_content)
        process_mail(msg)

    print('Process End')
    print('--------------------------------------------------------------------------')
    time.sleep(5)
