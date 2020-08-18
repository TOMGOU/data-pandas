import win_unicode_console
win_unicode_console.enable()
import sys,os
# from PyQt5.QtCore import *
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (QWidget, QPushButton, QLineEdit, QLabel, QApplication, QFileDialog)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

class crawler(QWidget):
  def __init__(self):
    super(crawler, self).__init__()
    self.switch = True
    self.initUI()

  def initUI(self):
     #关键词输入框和提示
    self.keyLabel = QLabel(self)
    self.keyLabel.move(30, 30)
    self.keyLabel.resize(100,30)
    self.keyLabel.setText("关键词：")
    self.key_le = QLineEdit(self)
    self.key_le.move(120,30)
    self.key_le.resize(250,30)

    # 存储文件选择按钮和选择编辑框
    self.target_btn = QPushButton('保存路径', self)
    self.target_btn.move(30, 90)
    self.target_btn.resize(80, 30)
    self.target_btn.clicked.connect(self.select_target)
    self.target_le = QLineEdit(self)
    self.target_le.move(120, 90)
    self.target_le.resize(250, 30)

    #保存按钮，调取数爬取函数等
    self.save_btn = QPushButton('开始',self)
    self.save_btn.move(200, 200)
    self.save_btn.resize(140, 30)
    self.save_btn.clicked.connect(self.kick)

    #用户提示区
    self.result_le = QLabel('请输入关键词和保存路径（包含excel文件名后缀）', self)
    self.result_le.move(30, 270)
    self.result_le.resize(340, 30)
    self.result_le.setStyleSheet('color: blue;')

    #整体界面设置
    # self.setGeometry(400, 400, 400, 400)
    self.resize(400, 400)
    self.center()
    self.setWindowTitle('小猪找货')#设置界面标题名
    self.show()
  
  def center(self):
    screen = QtWidgets.QDesktopWidget().screenGeometry()#获取屏幕分辨率
    size = self.geometry()#获取窗口尺寸
    self.move(int((screen.width() - size.width()) / 2), int((screen.height() - size.height()) / 2))#利用move函数窗口居中

  #保存的excel文件名称，要写上后缀名
  def select_target(self):
    target,fileType = QFileDialog.getSaveFileName(self, "选择保存路径", "C:/")
    self.target_le.setText(str(target))

  def set_label_func(self, text):
    self.result_le.setText(text)

  def switch_func(self, bools):
    self.switch = bools

  def kick(self):
    key_words = self.key_le.text().strip()#关键词
    target = self.target_le.text().strip()#保存路径
    if self.switch and key_words != '' and target != '':
      self.switch = False
      self.set_label_func('请耐心等待，数据疯狂抓取中...')
      self.my_thread = MyThread(key_words, target, self.set_label_func)#实例化线程对象
      self.my_thread.start()#启动线程
      self.my_thread.my_signal.connect(self.switch_func)

class MyThread(QThread):#线程类
  my_signal = pyqtSignal(bool)  #自定义信号对象。参数str就代表这个信号可以传一个字符串
  def __init__(self, key_words, target, set_label_func):
    super(MyThread, self).__init__()
    self.key_words = key_words
    self.target = target
    self.set_label_func = set_label_func

  def run(self): #线程执行函数
    length = self.fetchData(self.key_words, self.target, self.set_label_func)
    if length != 0:
      string = '获取了' + str(length) + '条数据，请直接查看excel表格！'
    else: 
      string = '没有抓取到数据，请尝试其他关键词！'
    self.set_label_func(string)
    self.my_signal.emit(True)  #释放自定义的信号

  def fetchData(self, key_words, target, set_label_func):
    option = webdriver.ChromeOptions()
    option.add_argument('headless')
    browser = webdriver.Chrome(options=option)
    # browser = webdriver.Chrome(executable_path='/Users/tangyong/Application/chromedriver')
    # browser = webdriver.Chrome(executable_path='/Users/tangyong/Application/chromedriver', options=option)
    browser.get('https://z.vanmmall.com/?from=singlemessage&isappinstalled=0')
    browser.find_element_by_id('kw').send_keys(key_words)
    search = browser.find_element_by_id('btn_search')
    search.click()
    self.sleep(2)
    while not self.isElementExist(browser, '.dropload-noData'):
      browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
      self.sleep(1)
    date = browser.find_elements_by_css_selector('td.date')
    shop = browser.find_elements_by_css_selector('td.shop div.name')
    msg = browser.find_elements_by_css_selector('td.msg')
    wechat = browser.find_elements_by_css_selector('td.call div button.wx')
    mobile = browser.find_elements_by_css_selector('td.call div button.phone')
    data = []
    cal_str = '已经获取了' + str(len(data)) + '条数据，正在合成excel，请耐心等待！'
    self.set_label_func(cal_str)
    data_length = len(date)
    for index in range(data_length):
      date_text = date[index].get_attribute('textContent')
      shop_text = shop[index].get_attribute('textContent')
      msg_text = msg[index].get_attribute('textContent')
      wechat_text = wechat[index].get_attribute("data-msg")
      mobile_text = mobile[index].get_attribute("data-msg")
      data.append([date_text, shop_text, msg_text, wechat_text, mobile_text])
      process_str = '当前数据合成进度： ' + str(index + 1) + ' / ' + str(data_length)
      self.set_label_func(process_str)
    browser.quit()
    self.set_label_func('excel表格保存中...')
    self.saveDataToExcel(data, target, key_words)
    return data_length

  def isElementExist(self, browser, element):
    flag=True
    try:
      browser.find_element_by_css_selector(element)
      return flag
    except:
      flag=False
      return flag

  def saveDataToExcel(self, data, target, key_words):
    if len(data) == 0:
      return
    data_df = pd.DataFrame(data)
    data_df.columns = ['日期','商家','产品信息+报价','微信号','手机号']
    writer = pd.ExcelWriter(target, engine='xlsxwriter')
    data_df.to_excel(writer, float_format='%.5f', index=True, index_label='索引', na_rep='--', sheet_name=key_words)
    worksheets = writer.sheets
    worksheet = worksheets[key_words]
    worksheet.set_column("B:B", 15)
    worksheet.set_column("C:C", 30)
    worksheet.set_column("D:D", 30)
    worksheet.set_column("E:E", 20)
    worksheet.set_column("F:F", 20)
    writer.save()

if __name__=="__main__":
  app = QApplication(sys.argv)
  ex = crawler()
  ex.show()
  sys.exit(app.exec_())