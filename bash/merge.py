#coding=utf-8
import re
import os
import time
import sys
import cgi
import datetime 

reload(sys) 
sys.setdefaultencoding('utf8')

import codecs


## 导出路径
root_path = u'D:/iphone/微信消息记录-李雄峰的 iPhone/'

path1 = u'D:/iphone/微信消息记录-李雄峰的 iPhone/201803242129-凤凰牌老熊/表格格式/'
path2 = u'D:/iphone/微信消息记录-李雄峰的 iPhone/201803270901-李雄峰/表格格式/'
path3 = u'D:/'

def merge(source_path_1, source_path_2, target_path) : 
	

	## unpath = unicode(inpath, "utf8")

	source1=codecs.open(source_path_1, "r",encoding='gbk', errors='ignore')
	source2=codecs.open(source_path_2, "r",encoding='gbk', errors='ignore')
	lines = {}
	for line in source1:
		if len(line)>690 and line[63:67] == 'xl24':
			datetime=line[233:252]
			date = line[233:243]
			time = line[244:252]
			
			pos_end = line.find('</td>', 335)
			name = line[335: pos_end].decode('gbk', errors='ignore').encode('utf-8')
			

			pos_start = line.find('x:str>', pos_end+5) + 6
			pos_end = line.find('</td>', pos_start)
			wechat_no = line[pos_start:pos_end]
			
			pos_start = line.find('x:str>', pos_end+5) + 6
			pos_end = line.find('</td>', pos_start)
			action = line[pos_start:pos_end].decode('gbk', errors='ignore').encode('utf-8')

			pos_start = line.find('x:str>', pos_end+5) + 6
			pos_end = line.find('</td>', pos_start)
			msgtype = line[pos_start:pos_end].decode('gbk', errors='ignore').encode('utf-8')

			pos_start = line.find('x:str>', pos_end+5) + 6
			pos_end = line.find('</td>', pos_start)
			msg = cgi.escape(line[pos_start:pos_end]).decode('gbk', errors='ignore').encode('utf-8')

			lines[datetime + name] = line
				
	source1.close()
	
	for line in source2:
		if len(line)>690 and line[63:67] == 'xl24':
			datetime=line[233:252]
			date = line[233:243]
			time = line[244:252]
			
			pos_end = line.find('</td>', 335)
			name = line[335: pos_end].decode('gbk', errors='ignore').encode('utf-8')
			

			pos_start = line.find('x:str>', pos_end+5) + 6
			pos_end = line.find('</td>', pos_start)
			wechat_no = line[pos_start:pos_end]
			
			pos_start = line.find('x:str>', pos_end+5) + 6
			pos_end = line.find('</td>', pos_start)
			action = line[pos_start:pos_end].decode('gbk', errors='ignore').encode('utf-8')

			pos_start = line.find('x:str>', pos_end+5) + 6
			pos_end = line.find('</td>', pos_start)
			msgtype = line[pos_start:pos_end].decode('gbk', errors='ignore').encode('utf-8')

			pos_start = line.find('x:str>', pos_end+5) + 6
			pos_end = line.find('</td>', pos_start)
			msg = cgi.escape(line[pos_start:pos_end]).decode('gbk', errors='ignore').encode('utf-8')

			lines[datetime + name] = line
				
	source2.close()

	
	target = codecs.open(target_path, 'w', encoding = 'gbk', errors='ignore')

	for key in sorted(lines.keys()) :
		target.write(lines[key])
	target.flush();
	target.close();
	
	print('finished merge ' + target_path)
	
	return
## 文件相对路径的位置
## file_name = u'支付产品技术交流群.xls'
##file_name = u'支付技术架构交流群.xls'
file_name = u'支付产品架构交流群.xls'
merge(path1+file_name, path2+ file_name, path3+ file_name)

