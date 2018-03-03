#coding=utf-8
import re
import os
import time
import sys
import cgi

reload(sys) 
sys.setdefaultencoding('utf8')

import codecs

## 导出原始excel文件所在的路径
root_path = u'D:/iphone/微信消息记录-李雄峰的 iPhone/'

## 最新的文件路径
last_folder = u' '

## 文件相对路径的位置
sub_path = u'/表格格式/支付产品技术交流群.xls'

## 寻找最近日期的xls文件

for cur_folder in os.listdir(root_path):
	if cmp(cur_folder, last_folder)>0 and  os.path.exists(root_path + cur_folder + sub_path):
		last_folder = cur_folder
	
unpath = root_path + last_folder + sub_path

print('使用最新文件: '+ unpath +'\n')

## 解析最近时间
last_date = time.strptime(last_folder[0:12], '%Y%m%d%H%M')


## 目标路径
dest_path = 'D:/cocolian/cocolian-docs/source/merged/_posts/' + time.strftime("%Y-%m-%d", last_date) + '-merged.markdown'

print('目标文件: '+ dest_path +'\n')

source=open(unpath)
target = codecs.open(dest_path, 'w', encoding = 'utf-8', errors='ignore')
target.write(u'---\n')              
target.write(u'layout:     source \n')                        
target.write(u'title:      "'+ time.strftime("%Y-%m-%d", last_date) +'-WeChat"\n')
target.write(u'date:       '+  time.strftime("%Y-%m-%d %H:%M", last_date) +':00\n')
target.write(u'author:     "PaymentGroup"\n')   
target.write(u'tag:		  [merged]\n')                           
target.write(u'---\n')              

## 当前日期
cur_date = ''

try:
	for line in source:
		if len(line)>690 and line[63:67] == 'xl24':
			datetime=line[233:252]
			date = line[233:243]
			time = line[244:252]
			
			if date!=cur_date:
				cur_date = date 
				target.write(u'## ' + cur_date + '   \n')
				

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

			if msgtype == u'文本':
				target.write(u'> ')
				target.write(datetime)
				target.write(u'  ')
				target.write(name)
				# target.write(u'   ')
				# target.write(wechat_no.decode('gbk', errors='ignore').encode('utf-8'))
				# target.write('  ')				
				# target.write(action.decode('gbk', errors='ignore').encode('utf-8'))
				# target.write('   ')
				# target.write(msgtype.decode('gbk', errors='ignore').encode('utf-8'))
				target.write(u'  \n   \n')				
				target.write(msg + u'  \n   \n')
			if msgtype == u'网页':
				target.write(u'> ')
				target.write(datetime)
				target.write(u'  ')
				target.write(name)
				# target.write(u'   ')
				# target.write(wechat_no.decode('gbk', errors='ignore').encode('utf-8'))
				# target.write('  ')				
				# target.write(action.decode('gbk', errors='ignore').encode('utf-8'))
				# target.write('   ')
				# target.write(msgtype.decode('gbk', errors='ignore').encode('utf-8'))
				target.write(u'  \n   \n')				
				target.write(msg + u'  \n   \n')
			if msgtype == unicode('照片壁纸','utf8'):
				target.write(u'> ')
				target.write(datetime)
				target.write(u'  ')
				target.write(name)
				# target.write(u'   ')
				# target.write(wechat_no.decode('gbk', errors='ignore').encode('utf-8'))
				# target.write('  ')				
				# target.write(action.decode('gbk', errors='ignore').encode('utf-8'))
				# target.write('   ')
				# target.write(msgtype.decode('gbk', errors='ignore').encode('utf-8'))
				target.write(u'  \n   \n')
				
				target.write('!['+ datetime + '](http://static.cocolian.org/img/'+ date.replace('-', '')+'_'+ time.replace(':','') + '.png' + ') \n   \n')

finally:
     source.close()

target.flush()
target.close()

print('完成导出：'+ unpath +' 到 ' + dest_path + '\n')