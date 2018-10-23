#-*- coding:utf-8 -*-

# import sys
# reload(sys)
# sys.setdefaultencoding('utf-8')
import time
import json
import paramiko
from xlrd import open_workbook
from xlutils.copy import copy


def connect(host):
    'this is use the paramiko connect the host,return conn'
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        # ssh.connect(host,username='root',allow_agent=True,look_for_keys=True)
        ssh.connect(host,username='root',password='passwd',allow_agent=True)
        return  ssh
    except:
        return  None

def command(args,outpath):
    'this is get the command the args to return the command'
    cmd = '%s %s' %(outpath,args)
    return cmd

def exec_commands(conn,cmd):
    'this is use the conn to excute the cmd and return the results of excute the command'
    stdin,stdout,stderr = conn.exec_command(cmd)
    results=stdout.read()
    return results
def excutor(host,outpath,args):
    conn = connect(host)
    if not conn:
        return [host,None]
    cmd = command(args,outpath)
    result = exec_commands(conn,cmd)
    # result = json.dumps(result)
    # return [host,result]
    print host
    return result

def Mem_Free(host):
    # print excutor(host,'who ','')
    # 测试
    Mem_Total = exec_commands(connect(host),"free -g  | sed -n '2,1p' | awk '{print $2}'")
    Mem_Used = exec_commands(connect(host),"free -g  | sed -n '2,1p' | awk '{print $3}'")
    Mem_Free = int(Mem_Total) - int(Mem_Used)
    print('\033[1;31;40m')
    print u"%s已连接成功." % host
    print('\033[0m')
    print u"可用内存获取成功，值为:" ,Mem_Free,'..'
    return Mem_Free


def Disk_Waring(host):
    Disk_list = exec_commands(connect(host), "df -h | awk 'NR>1' | awk '{print $5}'")
    # print Disk_list
    temp = Disk_list.split("\n")
    temp.pop()
    # print temp
    for x in temp:
        # print x
        # print int(x.strip("%"))
        i = int(x.strip("%"))
        if (i > 80):
            Disk_Waring = u"是"
            break
        else:
            Disk_Waring = u"否"
    print u"磁盘利用率获取成功...\n\t是否有警告：" , Disk_Waring
    return Disk_Waring

def CPU_Waring(host):
    # vmstat | awk 'NR>2' | awk '{print $14}'
    CPU_SY = exec_commands(connect(host),"vmstat | awk 'NR>2' | awk '{print $14}'")
    if (int(CPU_SY) > 80):
        print CPU_SY
        CPU_Waring = u"是"
    else:
        CPU_Waring = u"否"
    print u"CPU系统利用率获取成功....\n\t是否有警告：" ,CPU_Waring
    return CPU_Waring

def Excel_Exec(host):
    rexcel = open_workbook("D:\TJWXB_PY\ssh.xls")
    rows = rexcel.sheets()[0].nrows
    excel = copy(rexcel)
    table = excel.get_sheet(0)
    row = rows
    # print row
    data = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    table.write(row,0,data)
    table.write(row,1,host)
    table.write(row,2, Mem_Free(host))
    table.write(row,3, Disk_Waring(host))
    table.write(row,4, CPU_Waring(host))
    row += 1
    excel.save("D:\TJWXB_PY\ssh.xls")
    print u'%s已成功写入execl.....' % host



if __name__=='__main__':
    # Mem_Free = Mem_Free('x.x.x.11')
    # print Mem_Free
    # Disk_Waring = Disk_Waring('x.x.x.11')
    # print Disk_Waring
    # CPU_Waring = CPU_Waring('x.x.x.11')
    # print CPU_Waring
    #
    # data = time.strftime('%Y-%m-%d',time.localtime(time.time()))
    # print data
    Excel_Exec('x.x.x.11')
    Excel_Exec('x.x.x.16')
    Excel_Exec('x.x.x.86')