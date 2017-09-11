# -*- coding:utf-8 -*-
# 多线程并发：客户端
import socket

host, port = "192.168.2.144", 9000
c = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
c = connect((host, port))

while True:
	user_input = raw_input("msg to send:").strip()
	c.send(user_input)
	return_data = c.recv(1024)
	print "Receved", return_data

c.close()
