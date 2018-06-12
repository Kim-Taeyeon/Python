# -*- coding: utf-8 -*-
import comtypes.client
import os
import time
import json
import sys
import docx
import win32com
from win32com.client import Dispatch, constants

def init_powerpoint():
    powerpoint=comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible=1
    return powerpoint

def ppt_to_pdf(powerpoint,inputFileName,outputFileName,formatType=32):
    if outputFileName[-3:] != 'pdf':
        if outputFileName.endswith(".ppt"):
            outputFileName = outputFileName.rstrip(".ppt")+'.pdf'
        else:
            outputFileName = outputFileName.rstrip(".pptx")+'.pdf'
    deck=powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName,formatType)
    deck.Close()

def convert_files_folder(powerpoint):
    if os.walk("."):
        fullpath_list = []
        for root, dirs, files in os.walk("."):   # "."表示文件所在当前目录
            for file in files: 
                if file.endswith(".ppt") | file.endswith(".pptx"):  
                    fullpath = os.path.join(root, file)
                    full = os.path.abspath(fullpath)
                    ppt_to_pdf(powerpoint, full, full)

if __name__ == '__main__':
    print u"******* 欢迎使用PPT to PDF转换小工具 *******"
    print u"*                                          *"
    print u"*                                          *"
    print u"*                                          *"
    print u"*                                          *"
    print u"*                                          *"
    print u"*                                          *"
    print u"*                                          *"
    print u"*                                          *"
    print u"*                                          *"
    print u"*                                          *"
    print u"转换中..."

    powerpoint=init_powerpoint()
    cwd=os.getcwd()
    convert_files_folder(powerpoint)  
    powerpoint.Quit()
