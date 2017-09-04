# -*- coding:utf-8 -*-

import time
import os
import json
import codecs
import gzip
import StringIO
import urllib
import urllib2
import re
import sys
from os.path import exists

def getComment(Pgnum_max):
	for page in xrange(Pgnum_max):
		Comments = []
		page = str(page)
		print "第"+page+"页"
		url = "https://remark.vmall.com/remark/queryEvaluate.json?pid=738677717&pageNumber="+page+"&t=1503906858436&callback=jsonp1503906385929"
		headers = {
		'User-Agent':'',
		'Accept':'*/*',
		'Accept-Language':'',
		'Accept-Encoding':'',
		'Connection':'',
		'Referer':'',
		'Host':'remark.vmall.com'
		}

		req = urllib2.Request(url, None, headers)
		res = urllib2.urlopen(req)
		data = res.read()
		res.close()

		data = StringIO(data)
		gz = gzip.GzipFile(fileobj = data)
		ungz = gz.read()

		# 正则提取
		pattern_Comment = re.compile(r'"content":"\W*"')
		pattern_ReplyComment = re.compile(r'"replyContent":"\W*"')
		pattern_Username = re.compile(r'"custName":"\W*"')
		pattern_CommentTime = re.compile(r'"createtime":"....-..-..\s..:..:.."')
		pattern_RemarkLevel = re.compile(r'"remarkLevel":"\W*"')
		pattern_Score = re.compile(r'"score":"\d"')

		for Username in re.findall(pattern_Username, ungz):
			Comments.append(Username)

		for CommentTime in re.findall(pattern_CommentTime, ungz):
			Comments.append(CommentTime)

		for Comment in re.findall(pattern_Comment, ungz):
			Comments.append(Comment)

		for RemarkLevel in re.findall(pattern_RemarkLevel, ungz):
			Comments.append(RemarkLevel)

		for Score in re.findall(pattern_Score, ungz):
			Comments.append(Score)

		# 保存成HTML
		Comment = ",".join(Comments) # list转str
		cdir = u'./Comments/'
		if not exists(cdir):
			os.makedirs(cdir)
      
		f = codecs.open(cdir + u'page' + page + u'.txt', 'wb+', 'utf-8')
		f.write(Comment)

Pgnum_max = 1
getComment(Pgnum_max)
