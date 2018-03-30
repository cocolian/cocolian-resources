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

def gen_date_file(source_path, sub_path, target_folder, thedate) : 
	

	## unpath = unicode(inpath, "utf8")

	source=open(source_path)
	## 聊天内容总数；
	line_count = 0; 
	## 发言频率, 连续发言5次以上，就当做嘉宾了。 
	## 计算发言最多的人，作为嘉宾
	pres_name = 'PaymentGroup'
	cur_name = 'PaymentGroup'
	cur_count = 0 
	pres_count = 5
	## 计算聊天内容总数；
	
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

			if date == thedate and msgtype == u'文本':
				line_count += 1
				if name == cur_name :
					cur_count += 1 
					if cur_count > pres_count :
						pres_name = cur_name
						pres_count = cur_count
				else :
					cur_name = name 
					cur_count = 0 
				
	source.close()
	
	if line_count<2 :
		return
	
	
	target_path = r'D:/cocolian/cocolian-docs/source/'+target_folder+'/_posts/'+thedate+'-chat.markdown'
	target = codecs.open(target_path, 'w', encoding = 'utf-8', errors='ignore')
	target.write(u'---\n')              
	target.write(u'layout:     source \n')                        
	target.write(u'title:      "'+thedate+'-WeChat"\n')
	target.write(u'date:       '+ thedate+' 12:00:00\n')
	target.write(u'author:     '+ pres_name +'\n')
	target.write(u'lines:      '+ str(line_count) +' \n') 
	target.write(u'tag:		  [chat]\n')                        
	target.write(u'---\n')              

	source=open(source_path)

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

				if date == thedate and msgtype == u'文本':
					target.write(u'> ')
					target.write(time)
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
				if date == thedate and msgtype == u'网页':
					target.write(u'> ')
					target.write(time)
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
				if date == thedate and msgtype == unicode('照片壁纸','utf8'):
					target.write(u'> ')
					target.write(time)
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

	print('完成导出：'+ target_path +'\n')
	return
	
def gen_files(sub_path, target_folder) :
	## 默认导出最近10天的聊天记录
	exp_days = 10 
	if len(sys.argv)==2 :
		exp_days = int(sys.argv[1])
	range = datetime.timedelta(days=exp_days)

## 寻找最近日期的xls文件
	## 最新的文件路径
	last_folder = u' '

	for cur_folder in os.listdir(root_path):
		if cmp(cur_folder, last_folder)>0 and  os.path.exists(root_path + cur_folder + sub_path):
			last_folder = cur_folder
		
	unpath = root_path + last_folder + sub_path

	print('使用最新文件: '+ unpath +'\n')


	thedate = datetime.datetime.now().date() - range

	while thedate < datetime.datetime.now().date() : 
		gen_date_file(unpath, sub_path, target_folder, thedate.strftime("%Y-%m-%d"))
		thedate = thedate + datetime.timedelta(days = 1)
	return
## 文件相对路径的位置
gen_files(u'/表格格式/支付产品技术交流群.xls', 'proddev')
gen_files(u'/表格格式/支付技术架构交流群.xls', 'devarch')
gen_files(u'/表格格式/支付产品架构交流群.xls', 'prodarch')