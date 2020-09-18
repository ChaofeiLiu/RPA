# coding=utf-8
# 编译日期：2020-09-18 08:45:23
# 版权所有：www.i-search.com.cn
import ubpa.init_input as iinput
from ubpa.base_util import StdOutHook, ExceptionHandler
import ubpa.iexcel as iexcel
import ubpa.iie as iie
import ubpa.itools.rpa_str as rpa_str
import ubpa.itools.rpa_fun as rpa_fun
import ubpa.iocr as iocr
import ubpa.iimg as iimg
import time
import pdb
from ubpa.ilog import ILog
from ubpa.base_img import set_img_res_path
import getopt
from sys import argv
import sys
import os

class DEMO2:
     
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
      
    def Excel_write(self):
        value1=None
        #获取文本
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Excel_write,StepNodeTag:202009180744505107,Title:获取文本,Note:')
        value1=iie.get_text(waitfor=10.000,selector=r'#page > NAV:nth-of-type(1) > DIV:nth-of-type(1) > DIV:nth-of-type(1) > DIV:nth-of-type(1) > DIV:nth-of-type(3) > UL:nth-of-type(1) > LI:nth-of-type(8) > A:nth-of-type(1)',url=r'https://www.i-search.com.cn/course.html')
        print('[Excel_write] [获取文本] [202009180744505107]  返回值：[' + str(type(value1)) + ']' + str(value1))
        #单元格写入
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Excel_write,StepNodeTag:2020091807473501913,Title:单元格写入,Note:')
        tvar_2020091807473502914=iexcel.write_cell(path='D:/YiSaiQi/DEMO2/Day2.xls',text=value1)
        print('[Excel_write] [单元格写入] [2020091807473501913]  返回值：[' + str(type(tvar_2020091807473502914)) + ']' + str(tvar_2020091807473502914))
      
    def OCR(self):
        YZM=None
        #截图
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:OCR,StepNodeTag:20200918084212118111,Title:截图,Note:')
        tvar_20200918084212118112=iimg.capture_image(waitfor=30.000,win_title=r'中国工商银行企业网上银行 - Internet Explorer',left_indent=935,top_indent=458,width=90,height=34)
        print('[OCR] [截图] [20200918084212118111]  返回值：[' + str(type(tvar_20200918084212118112)) + ']' + str(tvar_20200918084212118112))
        #获取OCR文本
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:OCR,StepNodeTag:2020091808312854183,Title:获取OCR文本,Note:')
        YZM=iocr.general_recognize(image_path=tvar_20200918084212118112,apiKey='c578a472e212448fb9f78b17a79fffc9',secretKey='5588921ccc2047fd87a2e54fabee90ed')
        print('[OCR] [获取OCR文本] [2020091808312854183]  返回值：[' + str(type(YZM)) + ']' + str(YZM))
        #replace
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:OCR,StepNodeTag:2020091808351721592,Title:replace,Note:')
        tvar_2020091808351722593=rpa_str.replace(string=YZM,old=' ',new='')
        print('[OCR] [replace] [2020091808351721592]  返回值：[' + str(type(tvar_2020091808351722593)) + ']' + str(tvar_2020091808351722593))
        #输出
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:OCR,StepNodeTag:20200918083727282102,Title:输出,Note:')
        rpa_str.iprints(tvar_2020091808351722593)
      
    def for_demo(self):
        # For循环
        self.__logger.dlogs(job_no=self.job_no, logmsg='Flow:for_demo,StepNodeTag:2020091808042353439,Title:For循环,Note:')
        for i in [1,2,3,4,5]:
            #输出
            self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:for_demo,StepNodeTag:2020091808050105143,Title:输出,Note:')
            rpa_str.iprints(i)
      
    def pic_JC(self):
        #元素截图
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:pic_JC,StepNodeTag:2020091808241867472,Title:元素截图,Note:')
        tvar_2020091808243847274=iie.capture_element_img(waitfor=10.000,win_title=r'金融RPA_银行RPA机器人_RPA金融银行解决方案_艺赛旗 - Internet Explorer',selector=r'#fh5co-header > DIV:nth-of-type(1) > DIV:nth-of-type(1) > DIV:nth-of-type(1) > DIV:nth-of-type(1) > DIV:nth-of-type(1) > DIV:nth-of-type(2) > IMG:nth-of-type(1)',in_img_path=r'C:/Users/帅气可爱的小飞/Desktop/图片素材',url=r'https://www.i-search.com.cn/finance.html')
        print('[pic_JC] [元素截图] [2020091808241867472]  返回值：[' + str(type(tvar_2020091808243847274)) + ']' + str(tvar_2020091808243847274))
        #输出
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:pic_JC,StepNodeTag:2020091808252203975,Title:输出,Note:')
        rpa_str.iprints(tvar_2020091808243847274)
      
    def while_dmeo(self):
        carr=0
        # While循环
        self.__logger.dlogs(job_no=self.job_no, logmsg='Flow:while_dmeo,StepNodeTag:2020091807503206120,Title:While循环,Note:')
        while carr  <= 100:
            #相加
            self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:while_dmeo,StepNodeTag:2020091807551612223,Title:相加,Note:')
            carr=rpa_fun.add(a=carr,b=1)
            print('[while_dmeo] [相加] [2020091807551612223]  返回值：[' + str(type(carr)) + ']' + str(carr))
            # IF分支
            self.__logger.dlogs(job_no=self.job_no, logmsg='Flow:while_dmeo,StepNodeTag:2020091807565996529,Title:IF分支,Note:')
            if carr % 2 == 0:
                #输出
                self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:while_dmeo,StepNodeTag:2020091807574119732,Title:输出,Note:')
                rpa_str.iprints(carr)
            else:
                # Continue继续
                self.__logger.dlogs(job_no=self.job_no, logmsg='Flow:while_dmeo,StepNodeTag:2020091807582317635,Title:Continue继续,Note:')
                continue
      
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
    pro = DEMO2(robot_no=robot_no,proc_no=proc_no,job_no=job_no,input_arg=input_arg)
    pro.OCR()
