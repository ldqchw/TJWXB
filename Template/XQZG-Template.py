# -*- coding: utf-8-*-
import os
import sys
import getopt
import win32com
import xlwings as xw
from win32com.client import Dispatch, constants
#from envelopes import Envelope

def banner():
    print r'''
     _     _         _____                    _       _       
    | | __| | __ _  |_   _|__ _ __ ___  _ __ | | __ _| |_ ___ 
    | |/ _` |/ _` |   | |/ _ \ '_ ` _ \| '_ \| |/ _` | __/ _ \
    | | (_| | (_| |   | |  __/ | | | | | |_) | | (_| | ||  __/
    |_|\__,_|\__, |   |_|\___|_| |_| |_| .__/|_|\__,_|\__\___|
                |_|                    |_|
'''
def usage():
    print u'''\
py price.py [option][value]...

-h or --help     --help for detail
-o or --ofile    输出生成文件路径，默认为当前目录
-i or --ifile    输入配置文件路径，默认是当前目录:XQZG-Template-Parameters.xlsx
-t or --template 输入模板文件路径，默认为当前目录:XQZG-Template-V8.docx
'''
def get(argv):
   inputfile = pwd() + u'\\XQZG-Template-Parameters.xlsx'
   inputfile_T = pwd() + u"\\XQZG-Template-V8.docx"
   outputpath = pwd()
   try:
      opts, args = getopt.getopt(argv,"hi:t:o:",["help","ifile=","template=","ofile="])
      if args:
          banner()
          print '-h or --help for detail'
          usage()
          sys.exit(2)
   except getopt.GetoptError:
      banner()
      usage()
      sys.exit(2)
   for opt, arg in opts:
      if opt in ('-h',"--help"):
         banner()
         usage()
         sys.exit()
      elif opt in ("-i", "--ifile"):
         inputfile = arg
      elif opt in ("-t","--template"):
         inputfile_T = arg
      elif opt in ("-o", "--ofile"):
         outputpath = arg
         try:
             open(outputpath+"/rtest", 'w')
         except Exception, e:
             # ul_log.write(ul_log.fatal,'%s,%s,@line=%d,@file=%s' \
             # %(type(e),str(e),sys._getframe().f_lineno,sys._getframe().f_code.co_filename))
             print u"输出文件不能写入，请用有写权限的运行..."
             sys.exit(1)
   banner()
   return inputfile,inputfile_T,outputpath

def pwd():
    pwd = os.path.split(os.path.realpath(__file__))[0]
    return pwd;

def conf(pathxlsx):
    app_xlsx = xw.App(visible=True, add_book=False)
    wb = app_xlsx.books.open(pathxlsx)
    print u'正在从%s文件中获取配置...' % pathxlsx
    lists = wb.sheets[0].range('B1:B7').value
    wb.close()
    app_xlsx.quit()
    for x in lists:
        print u'成功获取: "%s"...' % x
    return lists

def replace(pathword,outputpath):
    lists = conf(pathxlsx)
    OldCompany = u'天津市XXXXXX公司'
    NewCompany = lists[0]
    OldArea = u'XX区'
    NewArea = lists[1]
    OldUrl = u'www.xxxxxx.com'
    NewUrl = lists[2]
    OldUrlname = u'XXXXXX网站'
    NewUrlname = lists[3]
    OldData1 = u'2018年X月X日'
    NewData1 = lists[4]
    OldData2 = u'监测时间：2018-XX-XX'
    NewData2 = lists[5]
    OldData3 = u'2018-xx-xx xx:xx:xx'
    NewData3 = lists[6]

    app = win32com.client.Dispatch('Word.Application')
    doc = app.Documents.Open(pathword)
    print u'正在生成模板中...'
    app.Visible = 0
    app.ScreenUpdating = 0
    app.Selection.Find.ClearFormatting()
    app.Selection.Find.Replacement.ClearFormatting()
    app.Selection.Find.Execute(OldCompany, False, False, False, False, False, True, 1, False, NewCompany, 2)
    app.Selection.Find.Execute(OldArea, False, False, False, False, False, True, 1, False, NewArea, 2)
    app.Selection.Find.Execute(OldUrl, False, False, False, False, False, True, 1, False, NewUrl, 2)
    app.Selection.Find.Execute(OldUrlname, False, False, False, False, False, True, 1, False, NewUrlname, 2)
    app.Selection.Find.Execute(OldData1, False, False, False, False, False, True, 1, False, NewData1, 2)
    app.Selection.Find.Execute(OldData2, False, False, False, False, False, True, 1, False, NewData2, 2)
    app.Selection.Find.Execute(OldData3, False, False, False, False, False, True, 1, False, NewData3, 2)
    print u"正在保存模板中..."
    print outputpath + "\\" + NewCompany + '.docx'
    doc.SaveAs(outputpath + "\\" + NewCompany + '.docx')
    doc.Close()
    # app.Documents.Close()
    app.Quit()
    print u'文件生成成功...'
def check_outputpath(outputpath):
    try:
        print u"检查输出路径是否有写权限..."
        open(outputpath + "/rtest", 'w')
    except Exception, e:
        # ul_log.write(ul_log.fatal,'%s,%s,@line=%d,@file=%s' \
        # %(type(e),str(e),sys._getframe().f_lineno,sys._getframe().f_code.co_filename))
        print u"输出文件不能写入，请用有写权限的运行..."
        sys.exit(1)
    var = u"检查输出路径成功..."
    return var



if __name__ == "__main__":
    inputfile,inputfile_T, outputpath = get(sys.argv[1:])
    print u"获取当前运行路径为：", pwd()
    print u'获取输入的配置文件为：', inputfile
    print u'获取输入的模板文件为：', inputfile_T
    print u'获取输出的文件路径：', outputpath
    print check_outputpath(outputpath)
    pathxlsx = inputfile
    pathword = inputfile_T
    replace(pathword,outputpath)





