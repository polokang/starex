__author__ = 'Hunter'
import re
from html.parser import HTMLParser
from urllib import request, parse
import win32com.client
import test

class MyHTMLParser(HTMLParser):
    def error(self, message):
        pass

    flag = 0
    res = []
    is_get_data = 0
    times = []
    places = []
    events = []

    def handle_starttag(self, tag, attrs):
        if tag == 'td':
            for attr in attrs:
                if re.match(r'trackListEven', attr[1]) or re.match(r'trackListOdd', attr[1]):
                    self.flag += 1

                    # def handle_endtag(self, tag):
                    # print("Encountered an end tag :", tag)

    def handle_data(self, data):
        if self.flag == 1:
            self.times.append(data)
            # print(self.flag, '.times:',  data)
        elif self.flag == 2:
            self.places.append(data)
        elif self.flag == 3:
            self.events.append(data)
            self.flag = 0



excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = -1
myBook = excel.Workbooks.Open(r'C:\Users\Hunter\PycharmProjects\starex\2018.xls')
mySheet = myBook.Worksheets("3")
row = mySheet.usedrange.rows.count
col = mySheet.usedrange.columns.count
i = 2
line = ""
codeList = []
nameList = []
while i <= row:
    wpsCode = str(mySheet.Cells(i, 4).Value)
    if wpsCode.startswith('S'):
        codeList.append(wpsCode)
        nameList.append(str(mySheet.Cells(i, 2).Value))
    i += 1

index = 0
for code in codeList:
    print("----------", index, ".", code, "-----------")
    login_data = parse.urlencode([('w', 'starex'), ('cno', code)])
    req = request.Request('http://www.starex.com.au/cgi-bin/GInfo.dll?EmmisTrack')

    with request.urlopen(req, data=login_data.encode('gbk')) as f:
        if f.status == 200:
            # print("----------", index, code, nameList[index], "-----------")
            index += 1
            parser = MyHTMLParser()
            parser.feed(f.read().decode('gbk'))

            i = 0
            while i < len(parser.times):
                print(parser.times[i], ':', parser.events[i])
                i += 1
            parser.times.clear()
            parser.places.clear()
            parser.events.clear()
