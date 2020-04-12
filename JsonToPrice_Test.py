#  -*- coding:utf-8 -*-
import clr
import sys
import os
import json
import logging
import time
import xml.etree.ElementTree as ET
import shutil
from ctypes import *
base_dir = os.path.abspath(os.path.join(os.getcwd()))
sys.path.append(base_dir +'\\Python3\\Server\\Dll\\')
clr.FindAssembly('TexttoXls.dll')
from TexttoXls import *
import tornado.ioloop
from tornado.options import define,options, parse_command_line
from tornado.web import RequestHandler
import uuid
define("port", default=8005, help="run on the given port", type=int)
dll = ConvertXls()
openxls = dll.openxls
GetSheetNums = dll.GetSheetNums
GetCell = dll.GetCell
InsertCell = dll.InsertCell
RemoveOneCol = dll.RemoveOneCol
RemoveOneRow = dll.RemoveOneRow
InsertRow = dll.InsertRow
closexls = dll.closexls
InsertNumCell = dll.InsertNumCell
InsertPicture = dll.InsertPicture
ChangeSheetName = dll.ChangeSheetName#void changesheetname(int sheetnum, string sheetname)
HideCol = dll.HideCol  #增加隐藏某一列功能， 不隐藏某一列。

SetCellRangeAddress = dll.SetCellRangeAddress
GetPathByXlsToHTML = dll.GetPathByXlsToHTML
# 20200227增加单元格行高和列宽
# 设置单元格颜色
#
#
SetColor = dll.SetColor    #设置单元格颜色
SetCellRowHeight = dll.SetCellRowHeight    #设置单元格行高
SetCellColumnWidth = dll.SetCellColumnWidth     #设置单元格列宽
CreateSheet = dll.CreateSheet

SetCellStyle = dll.SetCellStyle
ExcelToHtml = dll.ExcelToHtml
# 20200409
XlsToJson = dll.XlsToJson
Getbase64Picture = dll.Getbase64Picture
Insertbase64Picture = dll.Insertbase64Picture
Getbase64PictureTest = dll.Getbase64PictureTest
mMaxRow = 4096
mMaxCol = 100
mRow = 4096
mCol = 100

cellcolor = {
    'Black': 8,
    'Brown': 60,
    'Olive_Green': 59,
    'Dark_Green': 58,
    'Dark_Teal': 56,
    'Dark_Blue': 18,
    'Indigo': 62,
    'Grey_80_PERCENT': 63,
    'Dark_Red': 16,
    'Orange': 53,
    'DARK_YELLOW': 19,
    'Green': 17,
    'Teal':	21,
    'Blue':	12,
    'Blue_Grey': 54,
    'Grey_50_PERCENT': 23,
    'Red': 10,
    'LIGHT_ORANGE':	52,
    'LIME':	50,
    'SEA_GREEN': 57,
    'AQUA':	49,
    'LIGHT_BLUE': 48,
    'VIOLET': 20,
    'GREY_40_PERCENT': 55,
    'Pink':	14,
    'Gold':	51,
    'Yellow': 13,
    'BRIGHT_GREEN':	11,
    'TURQUOISE': 15,
    'SKY_BLUE':	40,
    'Plum':	61,
    'GREY_25_PERCENT': 22,
    'Rose':	45,
    'Tan': 47,
    'LIGHT_YELLOW': 43,
    'LIGHT_GREEN': 42,
    'LIGHT_TURQUOISE': 41,
    'PALE_BLUE': 44,
    'LAVENDER': 46,
    'White': 9,
    'CORNFLOWER_BLUE': 24,
    'LEMON_CHIFFON': 26,
    'MAROON': 25,
    'ORCHID': 28,
    'CORAL': 29,
    'ROYAL_BLUE': 30,
    'LIGHT_CORNFLOWER_BLUE': 31,
    'AUTOMATIC': 64

}

def CopyFile(srcfile, dstfile):
    if not os.path.isfile(srcfile):

        return False
        # print "%s not exist!"%(srcfile)
    else:
        fpath, fname = os.path.split(dstfile)  # 分离文件名和路径
        if not os.path.exists(fpath):
            os.makedirs(fpath)  # 创建路径
        if os.path.isfile(dstfile):
            os.remove(dstfile)
        shutil.copyfile(srcfile, dstfile)  # 复制文件
        return True


#日志的配置文件
def outputlog():
    # main_log_handler = logging.FileHandler(log_path +
    #                                        "/JsonToPrice_%s.log" % time.strftime("%Y-%m-%d_%H-%M-%S",
    #                                                                     time.localtime(time.time())), mode="w+",
    #                                        encoding="utf-8")
    # main_log_handler.setLevel(logging.DEBUG)
    # formatter = logging.Formatter("%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s")
    # main_log_handler.setFormatter(formatter)
    # logger.addHandler(main_log_handler)

    # 控制台打印输出日志
    console = logging.StreamHandler()  # 定义一个StreamHandler，将INFO级别或更高的日志信息打印到标准错误，并将其添加到当前的日志处理对象
    console.setLevel(logging.DEBUG)  # 设置要打印日志的等级，低于这一等级，不会打印
    formatter = logging.Formatter("%(asctime)s - %(levelname)s: %(message)s")
    console.setFormatter(formatter)
    logging.getLogger('').addHandler(console)
    # 进程号
    pid = os.getpid()
    # 本文件名（不含后缀.py）
    myfilename = os.path.split(__file__)[-1].split(".")[0]
    # 生成关闭进程的脚本文件
    #produce_stop_bat(pid, myfilename)

# 退出程序操作
class ExitSystem(tornado.web.RequestHandler):
    def get(self, *args, **kwargs):
        try:
            pass
        except:
            pass
        finally:
            print("***服务器退出***")
            os._exit(0)  # 退出程序

class ProduceExcel(tornado.web.RequestHandler):
    def post(self, *args, **kwargs):
        datacontent = self.get_argument('pricedata')
        guid = str(uuid.uuid1())  # 唯一标识符guid
        guid = ''.join(guid.split('-'))
        ToPrice = JsonToPrice(datacontent, guid)
        state = ToPrice.ToSrc()
        if state:
            self.set_header('Content-Type', 'application/x-xls')
            self.set_header('Content-Disposition', 'attachment; filename=NetPrice.xls')
            with open(state, 'rb') as f:
                while True:
                    data = f.read(1024)
                    if not data:
                        break
                    self.write(data)
                # # 记得有finish哦
                # self.write('1')
            self.finish()

        else:
            self.write(json.dumps({'result':'0'}).encode('utf8'))

class JsonToPrice(object):
    def __init__(self, datacontent, guid):
        self.datacontent = datacontent
        self.guid = guid
        self.excelname = ''
    def ToSrc(self):
        sheetsobj = json.loads(self.datacontent)
        self.excelname = sheetsobj['fileName']
        filename = self.producedatasheet(sheetsobj)
        if filename and os.path.exists(filename):
            return filename,self.excelname
        return False, ''

    def ToRGBColor(self, value):
        digit = list(map(str, range(10))) + list("ABCDEF")
        print(digit)
        if isinstance(value, tuple):
            string = '#'
            for i in value:
                a1 = i // 16
                a2 = i % 16
                string += digit[a1] + digit[a2]
            return string
        elif isinstance(value, str):
            a1 = digit.index(value[1]) * 16 + digit.index(value[2])
            a2 = digit.index(value[3]) * 16 + digit.index(value[4])
            a3 = digit.index(value[5]) * 16 + digit.index(value[6])
            return (a1, a2, a3)

    def producedatasheet(self, sheetsobj):
        if sheetsobj['type'] !='xls':
            sheetsobj['type'] = 'xls'

        excel = excelPath + sheetsobj['fileName'] + '_' + self.guid + '.' + sheetsobj['type']

        if not os.path.exists(excel):
            CopyFile(srcexcelfile, excel)
        else:
            os.remove(excel)
            CopyFile(srcexcelfile, excel)
        print(excel)
        openxls(excel)
        sheetnum = len(sheetsobj['sheets'])
        if sheetnum - 1 != 0: CreateSheet(sheetnum - 1)
        for k in range(len(sheetsobj['sheets'])):
            sheet = sheetsobj['sheets'][k]
            sheetname = sheet['sheetName']
            if sheetname !='':
                ChangeSheetName(k+1, sheetname)


            # if 'RowHeight' in sheet:
            #     rowobj = sheet['RowHeight']
            #     for row, height in rowobj.items():
            #         rownum = int(row[1:]) + 1
            #         logger.info(str(rownum) + '行高:' + str(height))
            #         SetCellRowHeight(k+1, rownum, height)
            #
            # if 'ColumnWidth' in sheet:
            #     print(sheet['ColumnWidth'])
            #     colobj = sheet['ColumnWidth']
            #     for col, width in colobj.items():
            #         colnum = int(col[1:]) + 1
            #         logger.info(str(colnum) + '列宽:' + str(width))
            #         SetCellColumnWidth(k+1, colnum, width)


            dataobj = sheet['data']  #data数据

            for key in sorted(dataobj.keys()):

                colobj = dataobj[key] # key 为 row number, colobj 为col对象
                i = int(key[1:])+1  # i 代表第几行
                collist = list(colobj.keys())
                print(collist)
                for col in range(0, len(collist)):
                    onecolobj = collist[col]    #列元素
                    j = int(onecolobj[1:]) +1 # j 代表第几列
                    cellobj = colobj[onecolobj] #单元格元素
                    print(k+1, i, j, str(cellobj['Text']))
                    #InsertCell(k+1, i, j, str(cellobj['Text']))
                    if '_mergeCount' in cellobj:
                        addcol = cellobj['_mergeCount']
                        #SetCellRangeAddress(k+1, i, i, j, j+addcol)   #int k, int rowstart, int rowend, int colstart, int colend

                    #必须存在单元格，改变颜色才有效
                    if 'style' in cellobj:
                        # style = cellobj['style']
                        # if 'Color' in style:
                        #     color = style['Color']
                        #     R, G, B = self.ToRGBColor(color.upper())
                        #     SetColor(k+1, i, j, R, G, B)
                        #     logger.info(str(i) + '行,' + str(j)+'列'+',Color='+color+','+str(cellcolor[color]))

                        style = cellobj['style']
                        # thin 代表框的上横线一条细线， medium 中等线 dashed 虚线 hair 点线
                        # thick 厚 double 俩条线 medium_dashed 中厚虚线，
                        if style:
                            print(k+1, i, j, style)
                            SetCellStyle(k+1, i, j, style)
        #print(GetPathByXlsToHTML(excel))
            string = ExcelToHtml(k)
            htmlpath = ServerPath+'\\ExcelHtml\\'+str(k)+'_'+sheetsobj['fileName'] + '_' + self.guid+'.html'
            if not os.path.exists(os.path.dirname(htmlpath)):
                os.makedirs(os.path.dirname(htmlpath))
            with open(htmlpath, 'w+', encoding='utf8') as f:
                f.write(string)
        closexls()
        openxls(excel)
        # with open('s.txt', 'r', encoding='utf8') as f:
        #     picturelist = json.loads(f.read())
        #     for picture in picturelist:
        #         print(picture)
        #         print(1, picture['startrow'], picture['startcol'], picture['endrow'], picture['endcol'],
        #               picture['picturedata'])
        #         print(type(picture['picturedata']))
        #
        #         Insertbase64Picture(1, picture['startrow'], picture['startcol'], picture['endrow'], picture['endcol'] + 1,
        #                             picture['picturedata'])
        for k in range(len(sheetsobj['sheets'])):
            sheet = sheetsobj['sheets'][k]
            picturelistobj = sheet['pictures']
            picturelist = json.loads(picturelistobj)
            for picture in picturelist:
                print(picture)
                print(1, picture['startrow'], picture['startcol'], picture['endrow'], picture['endcol'],
                      picture['picturedata'])
                print(type(picture['picturedata']))

                Insertbase64Picture(1, picture['startrow'], picture['startcol'], picture['endrow']+1, picture['endcol'] + 1,
                                    picture['picturedata'])
        # SetCellStyle(1, 3, 1, "font-weight:normal;font-name:宋体;font-size:12;text-align:CENTER;border-type:THIN None THIN MEDIUM;WrapText:True")





        closexls()
        return excel

application = tornado.web.Application([
    (r"/exit_localserver", ExitSystem),  # 退出程序
    (r"/Excel/ConAndDown/", ProduceExcel),
],autoreload=True)
template_path='templates'
static_path='static'


def jsontoexcel():
    srcexcelfile = base_dir + '\\reports\\pricesource.xls'
    excel = base_dir + '\\reports\\pricesourceTest.xls'
    if not os.path.exists(excel):
        CopyFile(srcexcelfile, excel)
    else:
        os.remove(excel)
        CopyFile(srcexcelfile, excel)

    # SetCellRowHeight(1, 1, 40)
    # InsertCell(1, 1, 1, '')
    # SetColor(1, 1, 1, 0, 0, 0)
    # #CreateSheet(2)
    # InsertCell(1, 3, 1, '李涛')
    # SetCellStyle(1, 3, 1, "color:red;font-weight:bold;font-size:11;font-name:宋体;border-type:thin;")
    # SetCellRangeAddress(1, 1, 1, 1, 3)


    # print (cellcolor.get('Black1'))
    excelPath = base_dir + '\\Python3\\Server\\Dll\\exceltojson\\json.txt'
    with open(excelPath, 'r', encoding='utf8') as f:
        datacontent = f.read()
    print(type(datacontent))
    sheetsobj = json.loads(datacontent)

    guid = str(uuid.uuid1())  # 唯一标识符guid
    guid = ''.join(guid.split('-'))
    ToPrice = JsonToPrice(datacontent, guid)
    state, excelname = ToPrice.ToSrc()
    # print(state, excelname)
    print('excel=', state)

def ExtractFileExt(filename):
    '''
        :param filename: '123.txt'
        :return: .txt
        '''
    if '.' not in filename:
        return ''
    for i in range(len(filename)-1, -1, -1):
        if filename[i] == '.' and (i != 0):
            return filename[:i]
    return ''
def exceltojson():
    srcexcelfile = base_dir + '\\reports\\经销商订货表.xls'
    excel = base_dir + '\\Python3\\Server\\Dll\\exceltojson\\经销商订货表.xls'
    #excel = base_dir + '\\reports\\pricesourceTest.xls'
    if not os.path.exists(excel):
        print(123)
        state = CopyFile(srcexcelfile, excel)
        print(state)

    print(123)
    jsontext = XlsToJson(excel)
    jsonobj = json.loads(jsontext)

    jsonobj['fileName'] = ExtractFileExt(os.path.basename(excel))
    with open(os.path.join(os.path.dirname(excel),'json.txt'), 'w+', encoding='utf8') as f:
        f.write(json.dumps(jsonobj, ensure_ascii=False))
    # srcexcelfile = base_dir + '\\reports\\pricesource.xls'
    # excel = base_dir + '\\reports\\pricesourceTest.xls'
    # if not os.path.exists(excel):
    #     CopyFile(srcexcelfile, excel)
    # else:
    #     os.remove(excel)
    #     CopyFile(srcexcelfile, excel)
    # openxls(excel)
    # for k in range(len(jsonobj['sheets'])):
    #     sheet = jsonobj['sheets'][k]
    #     picturelistobj = sheet['pictures']
    #     picturelist = json.loads(picturelistobj)
    #     for picture in picturelist:
    #         print(picture)
    #         print(1, picture['startrow'], picture['startcol'], picture['endrow'], picture['endcol'],
    #               picture['picturedata'])
    #         print(type(picture['picturedata']))
    #
    #         Insertbase64Picture(1, picture['startrow'], picture['startcol'], picture['endrow']+1, picture['endcol'] + 1,
    #                             picture['picturedata'])
    # closexls()
#测试插入图片接口
def TestInterface():
    srcexcelfile = base_dir + '\\reports\\零售订货表.xls'
    excel = base_dir + '\\reports\\pricesourceTest.xls'
    if not os.path.exists(excel):
        CopyFile(srcexcelfile, excel)
    else:
        os.remove(excel)
        CopyFile(srcexcelfile, excel)
    openxls(excel)
    pngfile = base_dir + '\\reports\\图片1.png'

    # 测试插入图片文件
    #InsertPicture(1,1,1,1,12,pngfile)startrow, int startcol, int lastrow, int lastcol
    #测试获取图片信息
    pictureliststring = Getbase64PictureTest(1)
    closexls()
    picturelist = json.loads(pictureliststring)
    excel = base_dir + '\\reports\\零售订货表诗尼曼.xls'
    openxls(excel)
    print(picturelist)
    with open('s.txt', 'w+', encoding='utf8') as f:
        f.write(json.dumps(picturelist, ensure_ascii=False))
    with open('s.txt', 'r', encoding='utf8') as f:
        picturelist = json.loads(f.read())

    for picture in picturelist:
        print(1, picture['startrow'], picture['startcol'], picture['endrow'], picture['endcol'],  picture['picturedata'])
        print(type(picture['picturedata']))

        Insertbase64Picture(1, picture['startrow'], picture['startcol'], picture['endrow'], picture['endcol']+1, picture['picturedata'])
    #SetCellStyle(1, 3, 1, "font-weight:normal;font-name:宋体;font-size:12;text-align:CENTER;border-type:THIN None THIN MEDIUM;WrapText:True")
    closexls()

def exceltohtml():
    jsontoexcel()
    excelfile = base_dir + '\\Python3\\Server\\Dll\\excel\\网络报价.xls'
    if not os.path.exists(excelfile):
        print(' No excel!!!')
    import pandas as pd
    import codecs
    xd = pd.ExcelFile(excelfile)
    df = xd.parse()
    print(os.path.dirname(excelfile)+'1.html')
    with codecs.open(os.path.dirname(excelfile)+'1.html', 'w', 'utf-8') as html_file:
        html_file.write(df.to_html(header=True, index=False))


if __name__=='__main__':
    # 如果日志文件夹不存在，则创建
    log_dir = "PriceLog"  # 日志存放文件夹名称
    log_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), log_dir)
    if not os.path.isdir(log_path):
        os.makedirs(log_path)
    print(log_path)
    # 设置logging
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    outputlog()
    excelPath = base_dir + '\\Python3\\Server\\Dll\\excel\\'
    srcexcelfile = base_dir + '\\reports\\pricesource.xls'
    excel = base_dir + '\\reports\\pricesourceTest.xls'
    ServerPath = os.path.abspath(os.path.join(os.getcwd(), '..')) + '\\nginx-1.0.11\\nginx-1.0.11\\html\\'
    print(ServerPath)
    #TestInterface()
    #jsontoexcel()
    #TestInterface()
    exceltojson()
    jsontoexcel()
    # excelPath = base_dir + '\\Python3\\Server\\Dll\\excel\\'
    # if not os.path.exists(excelPath):
    #     os.makedirs(excelPath)

    # # with open(file, 'r', encoding='utf8') as f:
    # #     datacontent = f.read()
    # # ToPrice = JsonToPrice(datacontent, excel)
    # # ToPrice.ToSrc()
    # tornado.options.parse_command_line()
    # http_server = tornado.httpserver.HTTPServer(application)
    # http_server.listen(options.port)
    # try:
    #     tornado.ioloop.IOLoop.instance().start()
    # except:
    #     pass