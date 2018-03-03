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

## 最新的文件路径
last_folder = u' '

## 文件相对路径的位置
sub_path = u'/表格格式/支付产品技术交流群.xls'

## 寻找最近日期的xls文件

for cur_folder in os.listdir(root_path):
	if cmp(cur_folder, last_folder)>0 and  os.path.exists(root_path + cur_folder + sub_path):
		last_folder = cur_folder
	
## unpath = root_path + last_folder + sub_path

unpath = u'D:/iphone/微信消息记录-李雄峰的 iPhone/201709301031-李雄峰/表格格式/支付产品技术交流群.xls'
print('使用最新文件: '+ unpath +'\n')

## unpath = unicode(inpath, "utf8")

source=open(unpath)
targetpath = r'D:/github/payment-wechat/stat-'+ datetime.datetime.now().strftime("%Y-%m-%d")+'-range.csv'
target = codecs.open(targetpath, 'w', encoding = 'utf-8', errors='ignore')
target.write(u'姓名, 频率, 最后发言时间 \n')              
freq ={}
last = {}
alias = {}

try:
	for line in source:
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

			if freq.has_key(wechat_no) :
				freq[wechat_no] += 1 
			else :
				freq[wechat_no] = 1
			
			alias[wechat_no] = name 
			last[wechat_no] = datetime 

finally:
     source.close()



unpath = u'D:/iphone/微信消息记录-李雄峰的 iPhone/201712161438-凤凰牌老熊/表格格式/支付产品技术交流群.xls'
print('使用最新文件: '+ unpath +'\n')

## unpath = unicode(inpath, "utf8")

source=open(unpath)

try:
	for line in source:
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

			if freq.has_key(wechat_no) :
				freq[wechat_no] += 1 
			else :
				freq[wechat_no] = 1

			last[wechat_no] = datetime 
			alias[wechat_no] = name 

finally:
     source.close()

for wechat_no in freq:
	target.write('\"' + wechat_no + "\"\t" + alias[wechat_no] + "\"\t" + str(freq[wechat_no]) +"\t"+ last[wechat_no]+'\n')

target.flush()
target.close()

print('完成导出：'+ targetpath +'\n')