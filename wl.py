import urllib.request
import re
import os
import pdfkit
import xlrd
from lxml.html import etree

# 获取费用总表网址
def geturl0(number):
    url0="http://as.yaxingbus.com/SiteManager/RepairCost?repairno=" + str(number)
    return url0

# 获取url的HTML的txt
def gethtml(url):
    requset = urllib.request.Request(url,headers=headers)
    response = urllib.request.urlopen(requset)
    html = response.read().decode("utf-8")
    # print(html1)
    return html

# 获得查看网址
def geturlchakan(html):
    html1 = etree.HTML(html)
    html2=html1.xpath('/html/body/div[3]/div[2]/div/div/div/div/div[2]/table/tbody/tr/td[16]/a[1]/@href')
    return html2

# 获得责任单位
def getname(html):
    html1 = etree.HTML(html)
    name=html1.xpath('//*[@id="tab_1_1"]/div/div[5]/div[1]/span[2]/text()')
    return name


if __name__ == '__main__':
    baseurl="http://as.yaxingbus.com"
    headers = {
      "User-Agent":"aaa---",
      "Cookie":"aaa---"
    }
            options = {
        'page-size': 'Letter',
        'margin-top': '0.75in',
        'margin-right': '0.75in',
        'margin-bottom': '0.75in',
        'margin-left': '0.75in',
        'encoding': "UTF-8",
        'custom-header': [
            ('Accept-Encoding', 'gzip')
        ],
        'cookie': [
            ('__RequestVerificationToken',
             'aaa---'),
            ('YXASS.AUTH',
             'aaa---'),
            ('YxAss.user', 'aaa---'),
        ],
        'no-outline': None,
        'outline-depth': 10,
    }
    data = xlrd.open_workbook(r"D:\Asiastar\2020.xls")
    table = data.sheet_by_index(0)
    numberx = table.col_values(0, 0)
    for number in numberx:
        # print(number)
        url0=geturl0(number)
        html0=gethtml(url0)
        urlchakan=baseurl+geturlchakan(html0)[0]
        htmlcahkan=gethtml(urlchakan)
        name=getname(htmlcahkan)[0]
        urlrar=baseurl+re.search(r'/SiteManager/Common/DownLoadFiles[?]entityId=[0-9]{5}',htmlcahkan).group()+"&cateId=2"
        urlpdf=baseurl+re.search(r"/SiteManager/RepairCost/Print/[0-9]{5}",html0).group()
        if os.path.exists(r"D:\\Asiastar" + "\\" + "text" + "\\" +name)==False:
            os.mkdir("D:\\Asiastar" + "\\" + "text" + "\\" +name)
        urllib.request.urlretrieve(urlrar, "D:\\Asiastar" + "\\" + "text" + "\\" +name+"\\"+ str(number) + ".rar")
        pdfkit.from_url(urlpdf, "D:\\Asiastar" + "\\" + "text" + "\\"+name+"\\" + str(number) + ".pdf", options=options)

