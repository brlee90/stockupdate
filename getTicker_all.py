#import packages
import yfinance as yf
from datetime import datetime
from xlrd import open_workbook
from xlutils.copy import copy
import tkinter as tk

class doc():
    #Open spreadsheet
    def open(filepath):
        rb = open_workbook(filepath)
        wb = copy(rb)
        s = wb.get_sheet(0)
        return wb, s

    #Write values to spreadsheet   
    def write(s, tList, tData):
        timestamp = datetime.today()
        s.write(0, 0, 'Timestamp:')
        s.write(0, 1, timestamp)
        for i in range(len(tList)):
            s.write(i+1, 0, tList[i])
            s.write(i+1, 1, tData[i])

    #Save spreadsheet
    def save(wb, filepath):
        wb.save(filepath)

class popup():
    def __init__(self, titleText, message):
        self.root = tk.Tk()
        self.root.title(titleText)
        self.root.eval('tk::PlaceWindow . center')

        label = tk.Label(text=message)
        label.pack(padx=25, pady=10)

        button = tk.Button(text='Close', command=self.quit)
        button.pack(pady=(0, 10))

        self.root.mainloop()
        
    def quit(self):
        self.root.destroy()

#Declare lists and spreadsheet filepath
tickerList = ['VTSAX', 'VTIAX', 'VGSLX', 'ARKK', 'ARKF', 'MSOS']
filepath = r'C:\Users\HP\Documents\Python\StockData.ods'
tickerData = []

#Get current market price for each fund
for ticker in tickerList:
    stock = yf.Ticker(ticker)
    tickerData.append(stock.info['regularMarketPrice'])

#Update spreadsheet and save
workbook, sheet = doc.open(filepath)
doc.write(sheet, tickerList, tickerData)
doc.save(workbook, filepath)

#Notify user of completion
title = 'Stonks^^'
message = 'Stock prices have been updated.'
popup(title, message)
