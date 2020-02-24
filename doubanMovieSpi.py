#_*_ coding: utf-8 _*_

import re
import urllib

from bs4 import BeautifulSoup
import xlwt

def getdata(baseurl):
    '''根据网址获取所有数据，10页'''
    dataList = []
    for i in range(0, 10):
        url = baseurl+str(i*25)
        html = askurl(url)
        parseHtml(html, dataList)
    return dataList


def parseHtml(html, dataList = []):

    '''解析网页内容'''
    # 定义链接正则表达式
    findLink = re.compile(r'<a href="(.*?)">')
    # 定义影片图片正则表达式
    findImage = re.compile(r'<img *src=(.*?)"')
    # 定义影片片名正则表达式
    findTitle = re.compile(r'<span class="title">(.*)</span>')

    # 定义影片评分正则表达式
    findGrade = re.compile(r'<span class="rating_num" property="v:average">(.*)</span>')
    # 定义影片评论数正则表达式
    findEstimateNum = re.compile(r'<span>(\d*)人评价</span>')
    # 定义影片概况正则表达式
    findMovieGeneral = re.compile(r'<span class="inq">(.*)</span>')

    soup = BeautifulSoup(html, "html.parser")
    # 找到所有class是item的div，发现都是具体电影项
    for item in soup.find_all('div', class_='item'):
        data = []
        # 数据转成字符串
        item = str(item)
        link = re.findall(findLink, item)[0]
        data.append(link)

        grade = re.findall(findGrade, item)[0]
        data.append(grade)

        estimateNum = re.findall(findEstimateNum, item)[0]
        data.append(estimateNum)

        if(len(re.findall(findMovieGeneral, item)) > 0):
            movieGeneral = re.findall(findMovieGeneral, item)[0]
            if len(movieGeneral) != 0:
                data.append(movieGeneral.replace("。", ""))
            else:
                data.append(' ')
        else:
            data.append(' ')

        titles = re.findall(findTitle, item)
        # 如果电影有两个名称，一个中文名、一个外国名，都保存下来
        if(len(titles) == 2):
            ctitle = titles[0]
            data.append(ctitle)
            otitle = titles[1].replace("/", "")
            data.append(otitle)
        else:
            data.append(titles[0])
            data.append(' ')
        dataList.append(data)
    return dataList


def askurl(url):
    """
        得到网页全部内容
        直接爬取内容会报错418——反爬取，通过模拟浏览器访问解决这个问题
    """
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36'}
    request = urllib.request.Request(url, headers=headers)

    html = ""
    try:
        # 获取响应内容
        response = urllib.request.urlopen(request)
        html = response.read()
        #print(html)
    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print("解析网页数据报错", e.code)
        if hasattr(e, "reason"):
            print("解析网页数据报错", e.reason)
    #print(html)
    return html


def saveData(dataList, savePath):
    """
        保存数据到excel
    """
    print("保存数据到：" + savePath)
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('豆瓣电影top250', cell_overwrite_ok=True)
    col = ('电影链接', '电影得分', '电影评价数', '电影简介', '电影中文名', '电影英文名')
    #   写EXCEL列标题
    for i in range(0, 6):
        sheet.write(0, i, col[i])
    #   写EXCEL数据，左闭右开
    for i in range(0, len(dataList)):
        data = dataList[i]
        print("行数据： " + str(data))
        for j in range(0, 6):
            #   从第 2 行开始写，第一行是列名
            sheet.write(i+1, j, str(data).split(",")[j].replace('\'', '').replace('[', '').replace(']', ''))
            # 这种写法报错
            # sheet.write(i+1, data)
    book.save(savePath)


print("开始爬取数据。。。")

baseurl = 'https://movie.douban.com/top250?start='
dataList = getdata(baseurl)
print("数据集大小：" + str(len(dataList)))
print(">>>>> data: " + str(dataList))
# 保存文件路径
savepath = u'E:/prepared/workspace/doubanMovieTop250/doubanMovieTop250.xlsx'
saveData(dataList, savepath)
