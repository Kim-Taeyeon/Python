# -*- coding:utf-8 -*-
# 多线程并发：服务器
import socket

class Mysocket(SocketServer.BaseRequestHandler):
	def handle(self):
		print "Got a connect from", self.client_address
		while True:
			data = self.request.recv(1024)
			print "recv", data

self.request.send(data.upper())

if __name__ == "__main__" :
	host = "0.0.0.0"
	port = 9001
	s = SocketServer.ThreadingTCPServer((host, port), Mysocket)  # 多线程并发
