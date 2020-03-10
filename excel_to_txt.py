"""
从excel中读取指定位置数据并保存到txt文件夹中
"""

import openpyxl

def process(file_path):
	data = openpyxl.load_workbook(file_path)
	table = data.active
	nrows = table.max_row
	texts = []
	for i in range(2,nrows+1):
		# 拼接文本位置
		text_location = 'C'+str(i)
		# 提取文本内容
		text = table[text_location].value
		if text:
			texts.append(text)
	# 追加方式写入txt，防止覆盖
	with open('1.txt','a') as f:
		for line in texts:
			f.write(line+'\n')
	data.save(file_path)


def main():
	file_path = r''
	process(file_path)
	

if __name__ == '__main__':
	main()