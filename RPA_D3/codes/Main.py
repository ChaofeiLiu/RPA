# coding=utf-8
# 编译日期：2020-09-17 16:47:52
# 版权所有：www.i-search.com.cn
import ubpa.init_input as iinput
from ubpa.base_util import StdOutHook, ExceptionHandler
import ubpa.iexcel as iexcel
import ubpa.ibox as ibox
import ubpa.itools.rpa_str as rpa_str
import pandas as pd
import ubpa.icsv as icsv
import time
import pdb
from ubpa.ilog import ILog
from ubpa.base_img import set_img_res_path
import getopt
from sys import argv
import sys
import os

class RPA_D3:
     
    def __init__(self,**kwargs):
        self.__logger = ILog(__file__)
        self.path = set_img_res_path(__file__)
        self.robot_no = ''
        self.proc_no = ''
        self.job_no = ''
        self.input_arg = ''
        if('robot_no' in kwargs.keys()):
            self.robot_no = kwargs['robot_no']
        if('proc_no' in kwargs.keys()):
            self.proc_no = kwargs['proc_no']
        if('job_no' in kwargs.keys()):
            self.job_no = kwargs['job_no']
        ILog.JOB_NO, ILog.OLD_STDOUT = self.job_no, sys.stdout
        sys.stdout = StdOutHook(self.job_no, sys.stdout)
        ExceptionHandler.JOB_NO, ExceptionHandler.OLD_STDERR = self.job_no, sys.stderr
        sys.excepthook = ExceptionHandler.handle_exception
        if('input_arg' in kwargs.keys()):
            self.input_arg = kwargs['input_arg']
            if(len(self.input_arg) <= 0):
                self.input_arg = iinput.load_init(__file__)
            if self.input_arg is None:
                sys.exit(0)
      
    def Dataframe_list(self):
        value_isearch03=None
        #读取Excel
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Dataframe_list,StepNodeTag:2020091715025272589,Title:读取Excel,Note:')
        tvar_2020091715025272690=pd.read_excel(io='D:/YiSaiQi/RPA_D3/mydata.xls')
        #表格过滤
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Dataframe_list,StepNodeTag:2020091714592597678,Title:表格过滤,Note:')
        value_isearch03=tvar_2020091715025272690[(tvar_2020091715025272690['产品代码'].str.startswith('i-Search-03'))]
        #设置变量
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Dataframe_list,StepNodeTag:20200917150802293115,Title:设置变量,Note:')
        value_isearch03=value_isearch03.values.tolist()
        # For循环
        self.__logger.dlogs(job_no=self.job_no, logmsg='Flow:Dataframe_list,StepNodeTag:20200917151128158119,Title:For循环,Note:')
        for tvar_20200917151128161120 in value_isearch03:
            #输出
            self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Dataframe_list,StepNodeTag:2020091715004175083,Title:输出,Note:')
            rpa_str.iprints(tvar_20200917151128161120)
      
    def Excel(self):
        sheet2_val=None
        sheet1_val=None
        #读取Excel
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Excel,StepNodeTag:20200917160132552143,Title:读取Excel,Note:')
        sheet1_val=pd.read_excel(io='D:/YiSaiQi/RPA_D3/mydata.xls')
        #读取Excel
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Excel,StepNodeTag:20200917160727776167,Title:读取Excel,Note:')
        sheet2_val=pd.read_excel(io='D:/YiSaiQi/RPA_D3/mydata.xls',sheet_name=1)
        #代码块
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Excel,StepNodeTag:20200917161108141183,Title:代码块,Note:')
        sheet1_val = sheet1_val.append(sheet2_val)
        #输出
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Excel,StepNodeTag:20200917161214523186,Title:输出,Note:')
        rpa_str.iprints(sheet1_val)
        # Try异常
        self.__logger.dlogs(job_no=self.job_no, logmsg='Flow:Excel,StepNodeTag:20200917162950407191,Title:Try异常,Note:')
        try:
            #创建excel
            self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Excel,StepNodeTag:20200917154223848138,Title:创建excel,Note:')
            iexcel.create_excel(path='D:/YiSaiQi/RPA_D3',file_name='Create_excel.xls')
        except Exception as e:
            #导出excel
            self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Excel,StepNodeTag:20200917163240519199,Title:导出excel,Note:')
            icsv.write_excel(path='D:/YiSaiQi/RPA_D3/Create_excel.xls',df=sheet1_val)
        finally:
            #导出excel
            self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Excel,StepNodeTag:20200917160944573180,Title:导出excel,Note:')
            icsv.write_excel(path='D:/YiSaiQi/RPA_D3/Create_excel.xls',df=sheet1_val)
      
    def flow1(self):
        #单元格读取
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:flow1,StepNodeTag:202009171430210059,Title:单元格读取,Note:')
        tvar_2020091714302100810=iexcel.read_cell(path='D:/YiSaiQi/RPA_D3/D3_demo.xls',cell_type=None)
        # Return返回
        self.__logger.dlogs(job_no=self.job_no, logmsg='Flow:flow1,StepNodeTag:2020091714434624064,Title:Return返回,Note:')
        return tvar_2020091714302100810
      
    def flow2(self,excel=None):
        #消息框
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:flow2,StepNodeTag:2020091714340135526,Title:消息框,Note:')
        ibox.msgs_box('成功' + excel,timeout=0)
      
    def Main(self):
        #子流程
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Main,StepNodeTag:2020091714405559957,Title:子流程,Note:')
        tvar20200917144055599571=self.flow1()
        #子流程
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Main,StepNodeTag:2020091714441293867,Title:子流程,Note:')
        tvar20200917144412938671=self.flow2(excel=tvar20200917144055599571)
 
if __name__ == '__main__':
    ILog.begin_init()
    robot_no = ''
    proc_no = ''
    job_no = ''
    input_arg = ''
    try:
        argv = sys.argv[1:]
        opts, args = getopt.getopt(argv,"hr:p:j:i:",["robot = ","proc = ","job = ","input = "])
    except getopt.GetoptError:
        print ('robot.py -r <robot> -p <proc> -j <job>')
    for opt, arg in opts:
        if opt == '-h':
            print ('robot.py -r <robot> -p <proc> -j <job>')
        elif opt in ("-r", "--robot"):
            robot_no = arg
        elif opt in ("-p", "--proc"):
            proc_no = arg
        elif opt in ("-j", "--job"):
            job_no = arg
        elif opt in ("-i", "--input"):
            input_arg = arg
    pro = RPA_D3(robot_no=robot_no,proc_no=proc_no,job_no=job_no,input_arg=input_arg)
    pro.Main()
