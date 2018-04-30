#coding:utf-8
import re
import urllib
import urllib2
import string
import numpy as np
import pandas as pd
import time

#读取整个网页
def getHtml(url):
    user_agent = 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'
    values = {'name' : 'WHY',
              'location' : 'SDU',
              'language' : 'Python' }
    headers = { 'User-Agent' : user_agent }
    data = urllib.urlencode(values)
    req = urllib2.Request(url, data, headers)
    response = urllib2.urlopen(req)
    the_page = response.read()
    return the_page

#获取影片名
def getPianming(html):
    reg = r' alt="(.+?)" /></div>'
    imgre = re.compile(reg)
    imglist = re.findall(imgre,html)
    return imglist

def getPiaofang(html):
    reg = r'<span class="m-span">累计票房<br />(.+?)</span>'
    imgre = re.compile(reg)
    imglist = re.findall(imgre,html)
    return imglist

#获取类型数据
def getLeixing(html):
    reg = r'<p>类型：(.+?)</p>'
    imgre = re.compile(reg)
    imglist = re.findall(imgre,html)
    return imglist

#获取片长
def getPianchang(html):
    reg = r'：(.+?)min'
    imgre = re.compile(reg)
    imglist = re.findall(imgre,html)
    return imglist

#获取上映时间
def getShijian(html):
    reg = r'上映时间：(.+?)\n'
    imgre = re.compile(reg)
    imglist = re.findall(imgre,html,)
    return imglist

#获取制式
def getZhishi(html):
    reg = r'<p>制式：(.+?)</p>'
    imgre = re.compile(reg)
    imglist = re.findall(imgre,html)
    return imglist

#获取地区
def getDiqu(html):
    reg = r'<p>国家及地区：(.+?)</p>'
    imgre = re.compile(reg)
    imglist = re.findall(imgre,html)
    return imglist

#获取发行公司
def getFXgongsi(html):
    reg = r'发行公司：<a target="_blank" href="http://www.cbooo.cn/c/6" title="(.+?)">'
    imgre = re.compile(reg)
    imglist = re.findall(imgre,html)
    return imglist

#获取导演
def getDaoyan(html):
    reg = r'title="(.+?)</a><span></span></p>'
    imgre = re.compile(reg)
    imglist = re.findall(imgre,html)
    return imglist

#获取主演
def getZhuyan(html):
    reg = r'\n(.+?)</a><span></span>'
    imgre = re.compile(reg)
    imglist = re.findall(imgre,html)
    return imglist


url = 'http://www.cbooo.cn/m/627896'
html = getHtml(url)
#print getHtml(url)

zhuyan = getZhuyan(html)
if '<p><a target=' in zhuyan[0]:
    del zhuyan[0]
zhuyan = '/'.join(zhuyan)

from xlwt import Workbook,Formula
book = Workbook()
sheet1 = book.add_sheet('Movie')
sheet1.write(0,1,u'片名')
sheet1.write(0,2,u'票房')
sheet1.write(0,3,u'电影类型')
sheet1.write(0,4,u'片长')
sheet1.write(0,5,u'上映时间')
sheet1.write(0,6,u'制式')
sheet1.write(0,7,u'地区')
sheet1.write(0,8,u'发行公司')
sheet1.write(0,9,u'导演')
sheet1.write(0,10,u'主演')

for i in range(396673,400000):
    print i
    try:
        url = url = 'http://www.cbooo.cn/m/' + str(i)
        html = getHtml(url)
        pianming = getPianming(html)#片名
        pianming = ''.join(pianming).decode('utf-8')

        piaofang = getPiaofang(html)#票房
        piaofang = ''.join(piaofang).decode('utf-8')

        leixing = getLeixing(html)#电影类型
        leixing = ''.join(leixing).decode('utf-8')

        pianchang = getPianchang(html)#片长
        pianchang = ''.join(pianchang).decode('utf-8')

        shijian = getShijian(html)#上映时间
        shijian = ''.join(shijian).decode('utf-8')

        zhishi = getZhishi(html)#制式
        zhishi = ''.join(zhishi).decode('utf-8')

        diqu = getDiqu(html)#地区
        diqu = ''.join(diqu).decode('utf-8')

        gongsi = getFXgongsi(html)#发行公司
        gongsi = ''.join(gongsi).decode('utf-8')

        daoyan = getDaoyan(html)#导演
        daoyan = ''.join(daoyan).decode('utf-8')

        #print ''.join(zhuyan).decode('utf-8')
        zhuyan = getZhuyan(html)#主演
        zhuyan = '/'.join(zhuyan).decode('utf-8')

        sheet1.write(i-360000,0,i)
        sheet1.write(i-360000,1,pianming)
        sheet1.write(i-360000,2,piaofang)
        sheet1.write(i-360000,3,leixing)
        sheet1.write(i-360000,4,pianchang)
        sheet1.write(i-360000,5,shijian)
        sheet1.write(i-360000,6,zhishi)
        sheet1.write(i-360000,7,diqu)
        sheet1.write(i-360000,8,gongsi)
        sheet1.write(i-360000,9,daoyan)
        sheet1.write(i-360000,10,zhuyan)

        book.save('Movie_Data.xls')
    
    except urllib2.URLError,e:
        print(e)

