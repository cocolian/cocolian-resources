#coding=utf-8
import re
import os
import time
import sys
import cgi
import datetime 
import shutil

reload(sys) 
sys.setdefaultencoding('utf8')

import codecs


## 导出路径
root_path = u'D:/iphone/微信消息记录-李雄峰的 iPhone/'
target_folder = u'D:/cocolian/cocolian-static/img/'


## 寻找最近日期的xls文件
## 最新的文件路径
last_folder = u' '
sub_path = u'/媒体文件'

for cur_folder in os.listdir(root_path):
	if cmp(cur_folder, last_folder)>0 and  os.path.exists(root_path + cur_folder + sub_path):
		last_folder = cur_folder
	
unpath = root_path + last_folder + sub_path

print('使用最新导出目录: '+ unpath +'\n')

for cur_foldername in os.listdir(unpath):
	cur_folder = os.path.join(unpath, cur_foldername)
	for cur_filename in os.listdir(cur_folder): 
		if cur_filename.endswith(".png"): 
			cur_path = os.path.join(cur_folder, cur_filename)
			target_path = os.path.join(target_folder, cur_filename)
			if not os.path.exists(target_path) or os.path.getsize(cur_path)>os.path.getsize(target_path) :
				shutil.copy(cur_path, target_folder) 
				print(cur_path +' --> ' + target_path)
		
	print('完成导出：'+ cur_folder +'\n')
	
