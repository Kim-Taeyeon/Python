# -*- coding: utf-8 -*-
import time
import json
import sys
import os
import docx
import win32com
from win32com.client import Dispatch, constants

# 在当前目录下查找出所有word文档并存为列表
def Findword():
	if os.walk("."):
	    Word_path_list = []
	    for root, dirs, files in os.walk("."):   # "."表示文件所在当前目录
	        for file in files: 
	            if file.endswith(".doc") | file.endswith(".docx"):	
	            # 这边判断是否是word文档，endswith() 方法用于判断字符串是否以指定后缀结尾,如果以指定后缀结尾返回True
	                Word_path = os.path.join(root, file)
	            	if "All Users" not in Word_path and "Windows" not in Word_path and "ProgramData" not in Word_path and "Program Files" not in Word_path and "$" not in Word_path :	# 这边要剔除临时文件等
	            		# print Word_path
	            		# write_file(Word_path)
	            		Word_path_list.append(os.path.abspath(Word_path))	# os.path.abspath(Word_path)  # 返回文件的绝对路径
	return Word_path_list

def write_file(par):
	# 创建存放路径的文件
	Log_dir = os.path.abspath(os.curdir) + "\\Check_result.txt"
	f = open(Log_dir, "a") # "w+"会覆盖原来的文件，所以用"a"
	a = u"未设置密级:"
	f.write(a.encode("gbk") + "\n")
	f.write(par + "\n")
	f.close()

def Public_interior(par):
	# 创建存放路径的文件
	Log_dir = os.path.abspath(os.curdir) + "\\Check_result.txt"
	f = open(Log_dir, "a") 
	a = u"内部公开:"
	f.write(a.encode("gbk") + "\n")
	f.write(par + "\n")
	f.close()

def main():
	print u"******* 欢迎使用'文档密级'自检小工具 *******"
	print u"*                                          *"
	print u"*                                          *"
	print u"*                                          *"
	print u"*                                          *"
	print u"*         未设置'密级'的文档路径           *"
	print u"* 存放在同级目录下的Check_result.txt文件中 *"
	print u"*                                          *"
	print u"*                                          *"
	print u"*                                          *"
	print u"*                                          *"
	print u"检查中...                                   "

	# 删除之前存放“文档路径”的txt文件
	Log_dir = os.path.abspath(os.curdir) + "\\Check_result.txt"
	if os.path.exists(Log_dir) : os.remove(Log_dir)
	
	Word_path_list = Findword()
	Unencrypted_file_list = []
	Public_interior_file_list = []

	# 将doc转换成docx
	word = win32com.client.Dispatch("Word.Application")
	
	# 查找页眉内容，Sections.Headers比正则表达式方便一些
	for word_path in Word_path_list:
		# print type(word_path)
		# print word_path
		time.sleep(1)
		doc = word.Documents.Open(word_path)
		time.sleep(1)
		# 中文字符前面加上u表示以utf-8形式显示
		Check_secret_000 = word.ActiveDocument.Sections[0].Headers[0].Range.Find.Execute(u"秘密")	# 包含关键字"秘密"则返回True，否则False
		Check_secret_001 = word.ActiveDocument.Sections[0].Headers[0].Range.Find.Execute(u"机密")	# 这边需要再添加
		Check_secret_002 = word.ActiveDocument.Sections[0].Headers[0].Range.Find.Execute(u"绝密")
		Check_secret_003 = word.ActiveDocument.Sections[0].Headers[0].Range.Find.Execute(u"内部公开")
		# 判断的结果，如果没有设置等级，就将该文件地址存到列表。
		if Check_secret_000 == False and Check_secret_001 == False and Check_secret_002 == False:
			if Check_secret_003 == True:
				Public_interior_file_list.append(word_path)
			else:
				Unencrypted_file_list.append(word_path)
		doc.Close() #关闭文档
		time.sleep(1)
	word.Quit() #关闭进程

	for i in Public_interior_file_list:
		Public_interior(i)

	for i in Unencrypted_file_list:
		write_file(i)
		

if __name__ == "__main__":
    main()
