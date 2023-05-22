# -*- coding = utf-8 -*-
# @time : 2023/4/27 19:40
# @Author : X²
# @File : main.py
# @Software : PyCharm
#先引入后面分析、可视化等可能用到的库
import datetime
from PIL import Image
import json
import sqlite3
from tkinter import ttk
import requests
import xlwt,xlrd,openpyxl
import pandas as pd
import mplfinance as mpl
import tkinter as tk
class spider:
    def open_new_window(self):
        # 创建新窗口
        new_window = tk.Toplevel(self.window)
        new_window.geometry("700x700+800+400")  # 将新窗口调整到相对于主窗口的右下角
        text = tk.Text(new_window)
        text.pack()
        df = pd.read_excel('data.xls')
        text.insert('end', df.to_string())
        # 在新窗口中添加标签和按钮
        # label = tk.Label(new_window, text="This is a new window!")
        # label.pack()
        #
        # button = tk.Button(new_window, text="Close", command=new_window.destroy)
        # button.pack()
    def OpenPic(self):
        # 打开图片
        img = Image.open('KLine.png')
        # 显示图片
        img.show()

    def startSearch(self):
        self.Num = self.entry.get()
        self.startDate = self.com_start.get()
        self.endDate = self.com_end.get()
        self.url="https://q.stock.sohu.com/hisHq?code=cn_"+self.Num+"&start="+self.startDate+"&end="+self.endDate
        # print(self.url)
        self.requestUrl()
        self.SaveInXls()
        self.DrawKLine()
        self.runsqlite()
    def creatDate(self):
        start_date = datetime.date(2023, 1, 1)
        end_date = datetime.date(2023, 5, 31)
        delta = end_date - start_date
        dates = [str(start_date + datetime.timedelta(days=i)).replace("-", "") for i in range(delta.days + 1)]
        dates = tuple(dates)
        return dates
    def drawGui(self):
        self.window = tk.Tk()
        self.window.title('股票信息查询')
        self.window.geometry('600x200')
        # 主界面绘制
        self.label = tk.Label(self.window, text='请输入股票代码', fg="black", font=("华文行楷", 30))
        self.label.grid(row=1, column=0)
        #显示提示
        self.entry = tk.Entry(self.window, font=("宋体", 25), fg="black")
        self.entry.grid(row=1, column=1)
        #给予输入框提示
        self.label_date_start = tk.Label(self.window, text="请输入开始日期", font=("宋体", 25))
        self.label_date_end = tk.Label(self.window, text="请输入结束日期", font=("宋体", 25))
        self.label_date_end.grid(row=2, column=0)
        self.label_date_start.grid(row=2, column=1)
        #输入输出框绘制
        self.com_start = ttk.Combobox(self.window)
        self.com_start["value"] = self.creatDate()
        self.com_end = ttk.Combobox(self.window)
        self.com_end["value"] = self.creatDate()
        self.com_end.grid(row=3, column=0)
        self.com_start.grid(row=3, column=1)
        #日期选择，且仅选择，不可自主输入
        self.button_check = ttk.Button(self.window, text="选择参数完成后点击开始查询", width=80,command=lambda: self.startSearch())
        self.button_check.grid(row=6, columnspan=2)
        self.button_picLine = ttk.Button(self.window, text="查看K线图",command=lambda :self.OpenPic())
        self.button_picLine.grid(row=9, column=0)
        self.button_Openxls = ttk.Button(self.window, text="打开xls",command=lambda :self.open_new_window())
        self.button_Openxls.grid(row=9, column=1)
        #绘制打开K线与xls的按钮
        self.window.mainloop()
        #允许窗口循环运行
    def __init__(self,url):
        self.url = url
     # "https://q.stock.sohu.com/hisHq?code=cn_600035&start=20230401&end=20230501"
    def requestUrl(self):
        str = requests.get(url=self.url)
        str = str.text
        print(str)
        str = str[1:-2]
        self.txt = json.loads(str)
        #切割处理，转化成标准json格式进行操作
    # print(txt)

    # url = 'https://q.stock.sohu.com/hisHq?code=cn_688366&stat=1&order=D&period=d&callback=historySearchHandler&rt=jsonp&0.9782923048414047'
    # str = requests.get(url=url)
    # str = str.text
    # str = str[22:-3]
    # txt = json.loads(str)
    #
    #
    #
    def SaveInXls(self):
        for item in self.txt["hq"]:
            print(item)
        #提取hq部分数据
        header = ['Date','Open','Close','涨跌额','涨跌幅','Low','High','成交量(手)','Volume','换手率']
        #自行命名数据类型
        book = xlwt.Workbook(encoding="utf-8",style_compression=0)
        #开辟新的文件与工作扇面sheet
        sheet = book.add_sheet("股票涨幅数据表",cell_overwrite_ok=True)
        s = 0
        for i in header:
            sheet.write(0,s,i)
            s+=1
        list = 1
        for item in self.txt["hq"]:
            for num in range(0,10):
                sheet.write(list,num,item[num])
            list+=1
        #利用循环操作写入
        book.save("data.xls")

    def DrawKLine(self):
        my_color = mpl.make_marketcolors(up='r',
                                         down='g',
                                         edge='inherit',
                                         wick='inherit',
                                         volume='inherit')
        my_style = mpl.make_mpf_style(marketcolors=my_color,
                                      figcolor='(0.82, 0.83, 0.85)',
                                      gridcolor='(0.82, 0.83, 0.85)')
        df = pd.read_excel("data.xls")
        #读取excel中表格的数据，从而规避js提出转化的过程
        df = df.filter(items=['Date','Open','High','Low','Close','Volume'])
        #为他们给予相应的正确的表名类型
        df = df[::-1]
        #由于获取的数据是倒置格式，应当翻转成顺序
        df.set_index(["Date"], inplace=True)
        df.index = pd.to_datetime(df.index)
        # 设置正确的时间索引格式
        print(df)
        # plot_kline_volume()
        mpl.plot(df, type='candle',mav=(3,6,9), volume=True,style = my_style,savefig = "Kline.png")
        mpl.plot(df, type='line',mav=(3,6,9), volume=True,style = my_style,savefig = "line.png")
        # ,savefig = 'KLine.png'
        # print(df)
    def SaveSql(self):
        conn = sqlite3.connect('mydatabase.db')
        cursor = conn.cursor()
    def create_table(self):
        conn = sqlite3.connect('data.db')
        #创建与链接
        cursor = conn.cursor()
        #创建游标
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS stock_data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                Date TEXT,
                Open REAL,
                Close REAL,
                Change REAL,
                ChangeRate REAL,
                Low REAL,
                High REAL,
                Volume REAL,
                Amount REAL,
                TurnoverRate REAL
            )
        """)
        #建表语句
        conn.commit()
        cursor.close()
        conn.close()
        #关闭操作

    def runsqlite(self):
        # 创建表格
        self.create_table()

        conn = sqlite3.connect('data.db')
        cursor = conn.cursor()
        cursor.executemany("""
             INSERT INTO stock_data (
                 Date, Open, Close, Change, ChangeRate, Low, High, Volume, Amount, TurnoverRate
             )
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
         """, self.txt["hq"])
        conn.commit()
        cursor.close()
        conn.close()
if __name__ == "__main__":
    # test = spider(url="https://q.stock.sohu.com/hisHq?code=cn_600035&start=20230401&end=20230501")
    test = spider(url="https://q.stock.sohu.com/hisHq?code=cn_600035&start=20230401&end=20230501")
    test.drawGui()