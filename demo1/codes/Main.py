# coding=utf-8
# 编译日期：2020-09-17 03:57:47
# 版权所有：www.i-search.com.cn
import ubpa.init_input as iinput
from ubpa.base_util import StdOutHook, ExceptionHandler
import ubpa.iie as iie
import ubpa.ikeyboard as ikeyboard
import ubpa.iexcel as iexcel
import ubpa.itools.rpa_str as rpa_str
import time
import pdb
from ubpa.ilog import ILog
from ubpa.base_img import set_img_res_path
import getopt
from sys import argv
import sys
import os

class demo1:
     
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
      
    def Baidu_news(self):
        '''打开百度新闻\n'''
        lv_2=20
        lv_1=10
        #网站
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Baidu_news,StepNodeTag:2020091515062370615,Title:网站,Note:')
        iie.open_url(ie_path='C:/Program Files (x86)/Internet Explorer/iexplore.exe',url='www.baidu.com')
        #鼠标点击
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Baidu_news,StepNodeTag:2020091515075726216,Title:鼠标点击,Note:点击新闻')
        time.sleep(1)
        iie.do_click_pos(waitfor=10.000,run_mode='unctrl',button='left',curson='center',continue_on_error='break',win_title=r'百度一下，你就知道 - Internet Explorer',selector=r'#s-top-left > A:nth-of-type(1)',url=r'https://www.baidu.com/')
        time.sleep(1)
      
    def Baidu_search(self,pv_1=None):
        input='RPA'
        #网站
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Baidu_search,StepNodeTag:2020091702300252617,Title:网站,Note:')
        iie.open_url(ie_path='C:/Program Files (x86)/Internet Explorer/iexplore.exe',url='www.baidu.com')
        #鼠标点击
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Baidu_search,StepNodeTag:2020091702303042024,Title:鼠标点击,Note:')
        iie.do_click_pos(waitfor=10.000,run_mode='unctrl',button='left',curson='center',continue_on_error='break',win_title=r'百度一下，你就知道 - Internet Explorer',selector=r'#kw',url=r'https://www.baidu.com/')
        #模拟按键
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Baidu_search,StepNodeTag:2020091702310358326,Title:模拟按键,Note:')
        ikeyboard.key_send_cs(waitfor=10.000,text='{LShift}')
        #模拟按键
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Baidu_search,StepNodeTag:2020091702315568728,Title:模拟按键,Note:')
        ikeyboard.key_send_cs(waitfor=10.000,text=input)
        #鼠标点击
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Baidu_search,StepNodeTag:2020091702332197730,Title:鼠标点击,Note:')
        iie.do_click_pos(waitfor=10.000,run_mode='unctrl',button='left',curson='center',continue_on_error='break',win_title=r'百度一下，你就知道 - Internet Explorer',selector=r'#su',url=r'https://www.baidu.com/')
      
    def CSM_demo(self):
        product_code='i-search-20200917-0351'
        #网站
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:2020091703094021399,Title:网站,Note:')
        iie.open_url(ie_path='C:/Program Files (x86)/Internet Explorer/iexplore.exe',url='http://122.112.200.222:9080/login.action')
        #设置文本
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:20200917031229648101,Title:设置文本,Note:')
        iie.set_text(waitfor=10.000,text=r'ceshi001',selector=r'#loginWrap > UL:nth-of-type(1) > LI:nth-of-type(1) > INPUT:nth-of-type(1)',url=r'http://122.112.200.222:9080/login.action')
        #设置文本
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:20200917031301891104,Title:设置文本,Note:')
        iie.set_text(waitfor=10.000,text=r'typSVU',selector=r'#loginWrap > UL:nth-of-type(1) > LI:nth-of-type(2) > INPUT:nth-of-type(1)',url=r'http://122.112.200.222:9080/login.action')
        #鼠标点击
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:20200917031353836106,Title:鼠标点击,Note:')
        iie.do_click_pos(waitfor=10.000,run_mode='unctrl',button='left',curson='center',continue_on_error='break',win_title=r'双录系统-录音、录像、录屏 - Internet Explorer',selector=r'#loginWrap > UL:nth-of-type(1) > LI:nth-of-type(2) > INPUT:nth-of-type(2)',url=r'http://122.112.200.222:9080/login.action')
        #鼠标点击
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:20200917031452058108,Title:鼠标点击,Note:')
        iie.do_click_pos(waitfor=10.000,run_mode='unctrl',button='left',curson='center',continue_on_error='break',win_title=r'双录系统-录音、录像、录屏 - Internet Explorer',selector=r'#popup_ok',url=r'http://122.112.200.222:9080/login.action')
        #鼠标点击
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:20200917031802702110,Title:鼠标点击,Note:')
        iie.do_click_pos(waitfor=10.000,run_mode='unctrl',button='left',curson='center',continue_on_error='break',win_title=r'双录系统-录音、录像、录屏 - Internet Explorer',selector=r'#frame-nav > UL:nth-of-type(1) > LI:nth-of-type(1) > A:nth-of-type(1)',url=r'http://122.112.200.222:9080/login.action')
        #鼠标点击
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:20200917031836401112,Title:鼠标点击,Note:')
        iie.do_click_pos(waitfor=10.000,run_mode='unctrl',button='left',curson='center',continue_on_error='break',win_title=r'双录系统-录音、录像、录屏 - Internet Explorer',selector=r'#MenuContext > TABLE:nth-of-type(1) > TBODY:nth-of-type(1) > TR:nth-of-type(1) > TD:nth-of-type(2) > LI:nth-of-type(5) > A:nth-of-type(1)',url=r'http://122.112.200.222:9080/login.action')
        #鼠标点击
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:20200917031929670114,Title:鼠标点击,Note:')
        iie.do_click_pos(waitfor=10.000,run_mode='unctrl',button='left',curson='center',continue_on_error='break',win_title=r'双录系统-录音、录像、录屏 - Internet Explorer',selector=r'#ListForm > DIV:nth-of-type(2) > DIV:nth-of-type(1) > DIV:nth-of-type(1) > A:nth-of-type(2)',title=r'理财管理')
        #设置文本
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:20200917032133267116,Title:设置文本,Note:')
        iie.set_text(waitfor=10.000,text=product_code,selector=r'body > DIV:nth-of-type(1) > FORM:nth-of-type(1) > DIV:nth-of-type(1) > DIV:nth-of-type(2) > UL:nth-of-type(1) > LI:nth-of-type(1) > INPUT:nth-of-type(1)',title=r'理财管理')
        #设置文本
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:20200917032229059118,Title:设置文本,Note:')
        iie.set_text(waitfor=10.000,text=r'LL-demo',selector=r'body > DIV:nth-of-type(1) > FORM:nth-of-type(1) > DIV:nth-of-type(1) > DIV:nth-of-type(2) > UL:nth-of-type(1) > LI:nth-of-type(2) > INPUT:nth-of-type(1)',title=r'理财管理')
        #设置文本
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:20200917032304015120,Title:设置文本,Note:')
        iie.set_text(waitfor=10.000,text=r'S8',selector=r'#proType',title=r'理财管理')
        #设置文本
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:20200917032344848122,Title:设置文本,Note:')
        iie.set_text(waitfor=10.000,text=r'5',selector=r'#proRiskLevel',title=r'理财管理')
        #设置文本
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:20200917032410932124,Title:设置文本,Note:')
        iie.set_text(waitfor=10.000,text=r'3',selector=r'#videoDtaTime',title=r'理财管理')
        #设置文本
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:20200917032513188130,Title:设置文本,Note:')
        iie.set_text(waitfor=10.000,text=r'刘浪测试',selector=r'#proDesc',title=r'理财管理')
        #鼠标点击
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:CSM_demo,StepNodeTag:20200917032903457132,Title:鼠标点击,Note:')
        iie.do_click_pos(waitfor=5.000,run_mode='unctrl',button='left',curson='center',continue_on_error='break',win_title=r'双录系统-录音、录像、录屏 - Internet Explorer',selector=r'#DetailForm > DIV:nth-of-type(1) > DIV:nth-of-type(3) > INPUT:nth-of-type(1)',title=r'理财管理')
      
    def Excel(self):
        text=None
        #获取文本
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Excel,StepNodeTag:2020091703541846134,Title:获取文本,Note:')
        text=iie.get_text(waitfor=10.000,selector=r'#frame-nav > UL:nth-of-type(1) > LI:nth-of-type(1) > A:nth-of-type(1)',url=r'http://122.112.200.222:9080/login.action')
        print('[Excel] [获取文本] [2020091703541846134]  返回值：[' + str(type(text)) + ']' + str(text))
        #输出
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Excel,StepNodeTag:2020091703550173641,Title:输出,Note:')
        rpa_str.iprints(text)
        #单元格写入
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Excel,StepNodeTag:2020091703553501645,Title:单元格写入,Note:')
        tvar_2020091703553501646=iexcel.write_cell(path='C:/Users/帅气可爱的小飞/Desktop/RPA_test1.xls',text=text)
        print('[Excel] [单元格写入] [2020091703553501645]  返回值：[' + str(type(tvar_2020091703553501646)) + ']' + str(tvar_2020091703553501646))
      
    def text_input(self):
        password=None
        #网站
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:text_input,StepNodeTag:2020091702470364538,Title:网站,Note:')
        iie.open_url(ie_path='C:/Program Files (x86)/Internet Explorer/iexplore.exe',url='www.baidu.com')
        #设置文本
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:text_input,StepNodeTag:2020091702471776644,Title:设置文本,Note:')
        iie.set_text(waitfor=10.000,text=r'RPA',selector=r'.s_ipt',url=r'https://www.baidu.com/')
        #鼠标点击
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:text_input,StepNodeTag:2020091703023316894,Title:鼠标点击,Note:')
        iie.do_click_pos(waitfor=10.000,run_mode='unctrl',button='left',curson='center',continue_on_error='break',win_title=r'百度一下，你就知道 - Internet Explorer',selector=r'#su',url=r'https://www.baidu.com/')
      
    def Main(self):
        pass
 
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
    pro = demo1(robot_no=robot_no,proc_no=proc_no,job_no=job_no,input_arg=input_arg)
    pro.Excel()
