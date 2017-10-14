# -*- coding:utf-8 -*-
# Python版本的telnet工具

import sys
import socket

host = sys.argv[1]
port = int(sys.argv[2])

def IsOpen(ip, port):
	s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
	s.settimeout(2) # 设置超时时间
	try:
		s.connect((ip, port))
		s.shutdown(1)
		print("You can connect to the IP %s, port %d" %(ip, port))
		return True
	except:
		print("You can't connect to the IP %s, port %d" %(ip, port))
		return False

a = IsOpen(host, port)
