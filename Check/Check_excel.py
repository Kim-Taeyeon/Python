# -*- coding: utf-8 -*-
import time
import json
import sys
import os
import docx
import win32com
import openpyxl
from win32com.client import Dispatch, constants
from openpyxl import Workbook

def Findexcel():
	if os.walk("."):
	    Excel_path_list = []
	    for root, dirs, files in os.walk("."):   # "."表示文件所在当前目录
	        for file in files: 
	            if file.endswith(".xlsx"):	 # file.endswith(".xls") |
	            # 这边判断是否是excel文档，endswith() 方法用于判断字符串是否以指定后缀结尾,如果以指定后缀结尾返回True
	                Excel_path = os.path.join(root, file)
	            	if "All Users" not in Excel_path and "Windows" not in Excel_path and "ProgramData" not in Excel_path and "Program Files" not in Excel_path and "$" not in Excel_path :	# 这边要剔除临时文件等
	            		# print Excel_path
	            		# write_file(Excel_path)
	            		Excel_path_list.append(os.path.abspath(Excel_path))	# os.path.abspath(Excel_path)  # 返回文件的绝对路径
	return Excel_path_list

def Write_file(par):
	# 创建存放路径的文件
	Log_dir = os.path.abspath(os.curdir) + "\\Check_excel_result.txt"
	f = open(Log_dir, "a") # "w+"会覆盖原来的文件，所以用"a"
	a = u"未设置密级:"
	f.write(a.encode("gbk") + "\n")
	f.write(par + "\n")
	f.close()

def Public_interior(par):
	# 创建存放路径的文件
	Log_dir = os.path.abspath(os.curdir) + "\\Check_excel_result.txt"
	f = open(Log_dir, "a") 
	a = u"内部公开:"
	f.write(a.encode("gbk") + "\n")
	f.write(par + "\n")
	f.close()

def Unreadable(par):
	Log_dir = os.path.abspath(os.curdir) + "\\Check_excel_result.txt"
	f = open(Log_dir, "a") 
	a = u"不可读文件:"
	f.write(a.encode("gbk") + "\n")
	f.write(par + "\n")
	f.close()



def main():
	print u"******* 欢迎使用'文档密级'自检小工具 *******"
	print u"*                                          *"
	print u"*                                          *"
	print u"*                                          *"
	print u"*                                          *"
	print u"*      未设置'密级'的文档路径存放在        *"
	print u"* 同级目录下的Check_excel_result.txt文件中 *"
	print u"*                                          *"
	print u"*                                          *"
	print u"*                                          *"
	print u"*                                          *"
	print u"检查中...                                   "

	# 删除之前存放“文档路径”的txt文件
	Log_dir = os.path.abspath(os.curdir) + "\\Check_excel_result.txt"
	if os.path.exists(Log_dir) : os.remove(Log_dir)
	
	Excel_path_list = Findexcel()
	Unencrypted_file_list = []
	Public_interior_file_list = []
	Unreadable_file_list = []
	
	# 查找页眉内容，Sections.Headers比正则表达式方便一些
	for excel_path in Excel_path_list:
		
		load_workbook = win32com.client.Dispatch("Excel.Application")

		wb = openpyxl.load_workbook(excel_path)
		for ws in wb.worksheets:
 			# 设置首页与其他页不同
 			ws.HeaderFooter.differentFirst = True
 			print "1111111"
 			print  ws.HeaderFooter.oddHeader
 			a = []
 			b = ws.HeaderFooter.oddHeader
 			
 			print b.
 			print "2222222"
 			# 设置奇偶页不同
 			ws.HeaderFooter.differentOddEven = True
 			# 设置首页页眉页脚
 			ws.firstHeader.left = _HeaderFooterPart('第一页左页眉', size=24, color='FF0000')
 			ws.firstFooter.center = _HeaderFooterPart('第一页中页脚', size=24, color='00FF00')
 			# 设置奇偶页页眉页脚
 			ws.oddHeader.right = _HeaderFooterPart('奇数页右页眉')
 			ws.oddFooter.center = _HeaderFooterPart('奇数页中页脚')
 			ws.evenHeader.left = _HeaderFooterPart('偶数页左页眉')
 			ws.evenFooter.center = _HeaderFooterPart('偶数页中页脚')
 			wb.save('new_'+excel_path)


		# if os.access(excel_path, os.R_OK):	# 加入了是否可读的判断，防止文件损坏
		# 	excel = win32com.client.Dispatch("Excel.Application")
		# 	doc = excel.Workbooks.Open(excel_path)
		# 	# time.sleep(0.5)
		# 	# 中文字符前面加上u表示以utf-8形式显示
		# 	print excel.ActiveDocument.Tables[0].Rows[0].Cells[0]
		# 	Check_secret_000 = excel.ActiveDocument.Tables[0].Headers[0].Range.Find.Execute(u"秘密")	# 包含关键字"秘密"则返回True，否则False
		# 	Check_secret_001 = excel.ActiveDocument.Tables[0].Headers[0].Range.Find.Execute(u"机密")	# 这边需要再添加
		# 	Check_secret_002 = excel.ActiveDocument.Tables[0].Headers[0].Range.Find.Execute(u"绝密")
		# 	Check_secret_003 = excel.ActiveDocument.Tables[0].Headers[0].Range.Find.Execute(u"内部公开")
		# 	# 判断的结果，如果没有设置等级，就将该文件地址存到列表。
		# 	if Check_secret_000 == False and Check_secret_001 == False and Check_secret_002 == False:
		# 		if Check_secret_003 == True:
		# 			Public_interior_file_list.append(excel_path)
		# 		else:
		# 			Unencrypted_file_list.append(excel_path)
		# 	doc.Close() #关闭文档

		# 	excel.Workbooks.Close()
		# 	excel.Quit()
		# 	# time.sleep(0.5)
		# else:
		# 	Unreadable_file_list.append(excel_path)
	# excel.Quit() #关闭进程

	for i in Public_interior_file_list:
		Public_interior(i)

	for i in Unencrypted_file_list:
		Write_file(i)
		
	for i in Unreadable_file_list:
		Unreadable(i)

if __name__ == "__main__":
    main()
