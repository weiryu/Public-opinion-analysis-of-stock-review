"""
从excel中读取数据绘制蜡烛图
"""

from pandas import DataFrame, Series
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import dates as mdates
from matplotlib import ticker as mticker
from matplotlib.finance import candlestick_ohlc
from matplotlib.dates import DateFormatter, WeekdayLocator, DayLocator, MONDAY, YEARLY
from matplotlib.dates import MonthLocator, MONTHLY
import datetime
import pylab
import time
import os

def draw_candlestick_ohlc(file_path):
	MA1 = 10 #10日均线
	MA2 = 50 #50日均线
	# 数据预处理
	data = pd.DataFrame(pd.read_excel(file_path,sheet_name=1,index_col='日期'))
	data = data.sort_index()
	stdata = pd.DataFrame({'DateTime':data.index,'Open':data.开盘价,'High':data.最高价,
		'Close':data.收盘价,'Low':data.最低价})
	stdata['DateTime'] = mdates.date2num(stdata['DateTime'].astype(datetime.date)) #日期转化为天数
	daysreshape = stdata.reset_index()
	daysreshape = daysreshape.reindex(columns=['DateTime','Open','High','Low','Close'])
	print(daysreshape)
	Av1 = pd.rolling_mean(daysreshape.Close.values,MA1)
	Av2 = pd.rolling_mean(daysreshape.Close.values,MA2)
	SP = len(daysreshape.DateTime.values[MA2-1:])
	fig = plt.figure(facecolor='#07000d',figsize=(15,10))
	ax1 = plt.subplot2grid((6,4),(1,0),rowspan=4,colspan=4,axisbg='#07000d')
	candlestick_ohlc(ax1,daysreshape.values[-SP:],width=.6,colorup="#ff1717",colordown='#53c156')
	# 画出均线
	label1 = str(MA1)+' SMA'
	label2 = str(MA2)+' SMA'
	ax1.plot(daysreshape.DateTime.values[-SP:],Av1[-SP:],'#e1edf9',label=label1,linewidth=1.5)
	ax1.plot(daysreshape.DateTime.values[-SP:],Av2[-SP:],'#4ee6fd',label=label2,linewidth=1.5)
	ax1.grid(True,color='w')
	ax1.xaxis.set_major_locator(mticker.MaxNLocator(10))
	ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d"))
	ax1.yaxis.label.set_color('w')
	ax1.spines['bottom'].set_color('#5998ff')
	ax1.spines['top'].set_color('#5998ff')
	ax1.spines['left'].set_color('#5998ff')
	ax1.spines['right'].set_color('#5998ff')
	ax1.tick_params(axis='y',colors='w')
	plt.gca().yaxis.set_major_locator(mticker.MaxNLocator(prune='upper'))
	ax1.tick_params(axis='x',colors='w')
	plt.ylabel('Stock price and Volume')
	plt.show()



def main():
	file_folder = r''
	file = r'000001_2019.xlsx'
	file_path = os.path.join(file_folder,file)
	draw_candlestick_ohlc(file_path)

if __name__ == '__main__':
	main()