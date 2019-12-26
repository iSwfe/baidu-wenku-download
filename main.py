#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
from urllib import parse
import urllib.parse as urllibParse
import urllib.request as urllibReq
import json
import sys
from importlib import reload
import config
import log

reload(sys)
# sys.setdefaultencoding('utf-8')

# 全局变量定义
logger = log.Log(config.log_dir, config.log_name)


# 获取属性
def getAttribute(info, key):
	key = '\'' + key + '\''
	tmp_pos = info.find(key)
	tmp_str = info[tmp_pos + len(key):]
	tmp_str = tmp_str[tmp_str.find('\'') + 1:]
	return tmp_str[0: tmp_str.find('\'')]


# 读页面获取文档名称、类型、页数等信息
def httpGet(url):
	response = urllibReq.urlopen(url)
	res = response.read()
	# 1、分析http用的编码格式，再编解码
	head_begin = res.find(b'<head>')
	head_end = res.find(b'</head>')
	head_str = res[head_begin:head_end]
	charset = head_str.find(b'charset=')
	if charset != -1:
		charset_str = head_str[charset + len('charset='):]
		pos = 0
		while True:
			if (ord('a') <= charset_str[pos] <= ord('z')) or \
					(ord('0') <= charset_str[pos] <= ord('9')) or \
					charset_str[pos] == '-':
				pos += 1
			else:
				break
		charset = charset_str[:pos]
		if charset != b'utf-8':
			res = res.decode(charset.decode())
	# 2、获取文档类型、文档名称、每页URL信息
	WkInfo = {}
	wkinfo_begin = res.find('WkInfo.DocInfo')
	wkinfo_end = res.find('Data.set(\'WkInfo\', WkInfo)')
	wkinfo_str = res[wkinfo_begin: wkinfo_end]
	WkInfo['title'] = getAttribute(wkinfo_str, 'title').replace(' ', '-')
	WkInfo['docTypeNum'] = getAttribute(wkinfo_str, 'docTypeNum')
	# WkInfo['docType'] = getAttribute(wkinfo_str, 'docType')
	WkInfo['docType'] = config.docTypeBDWK[WkInfo['docTypeNum']]
	WkInfo['totalPage'] = int(getAttribute(wkinfo_str, 'totalPageNum'))
	# 获取下载链接
	tmp_pos = wkinfo_str.find('WkInfo.htmlUrls')
	if tmp_pos != -1:  # txt 类型没有htmlUrls
		tmp_str = wkinfo_str[tmp_pos:]
		tmp_str = tmp_str[tmp_str.find('\'') + 1:]
		tmp_str = tmp_str[:tmp_str.find('\'')]
		tmp_str = urllibParse.unquote(tmp_str)
		WkInfo['htmlUrls'] = json.loads(tmp_str.replace('\\x22', '"'))
	logger.info('获取文库URL[' + url + ']，信息如下：' + str(WkInfo))
	return WkInfo


def usage():
	print('Usage:' + os.path.basename(sys.argv[0]) + ' <文库url> [<pageStart>]')
	sys.exit(-1)


def main():
	if len(sys.argv) not in [2, 3]:
		usage()
	url = sys.argv[1]
	if not url.startswith('http') and not url.startswith('https'):
		usage()
	if not url.find('wenku.baidu.com'):
		print('only support baidu wenku currently!')
		return
	pageStart = 1
	try:
		if len(sys.argv) is 3:
			pageStart = int(sys.argv[2])
	except ValueError:
		usage()

	WkInfo = httpGet(url)
	result = False
	if WkInfo['docType'] == 'txt':
		import downTXT
		d = downTXT.DownTXT(config.file_dir, url, WkInfo)
		result = d.down()
	elif WkInfo['docType'] == 'pdf':
		import downPDF
		d = downPDF.DownPDF(config.file_dir, url, WkInfo)
		result = d.down()
	elif WkInfo['docType'] in ['docx', 'doc']:
		import downDocx
		d = downDocx.DownDocx(config.file_dir, url, WkInfo)
		result = d.down(pageStart, WkInfo['totalPage'])
	else:
		logger.info('暂时不支持该类型[' + WkInfo['docType'] + ']文档下载，敬请期待！')
	if result:
		print('download success! file is saved in dir: ' + config.file_dir)
	else:
		print('download failed!')
	return


if __name__ == '__main__':
	main()
