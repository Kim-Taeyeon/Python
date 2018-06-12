# -*- coding: utf-8 -*-
# 根据txt文档中的文件路径删除文件
import os

Log_dir = os.path.abspath(os.curdir) + "\\Check_result.txt"
with open(Log_dir,"r") as Delete_file:
	for each_line in Delete_file:
		new_line = each_line.replace("\n", "")
		if os.path.exists(new_line) :
			os.remove(new_line)
