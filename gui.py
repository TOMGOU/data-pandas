import win_unicode_console
win_unicode_console.enable()
import sys,os
from PyQt5.QtCore import *
from PyQt5.QtWidgets import (QWidget, QPushButton, QLineEdit, QLabel, QApplication, QFileDialog)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

class upload(QWidget):
  def __init__(self):
    super(upload, self).__init__()
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
    self.save_btn.move(120, 200)
    self.save_btn.resize(140, 30)
    self.save_btn.clicked.connect(self.kick)

    #执行成功返回值显示位置设置
    self.result_le = QLabel(self)
    self.result_le.move(30, 270)
    self.result_le.resize(340, 30)

    #整体界面设置
    self.setGeometry(400, 400, 400, 400)
    self.setWindowTitle('小猪找货数据爬取')#设置界面标题名
    self.show()
  #保存的excel文件名称，要写上后缀名
  def select_target(self):
    target,fileType = QFileDialog.getSaveFileName(self, "选择保存路径", "C:/")
    self.target_le.setText(str(target))
  def kick(self):
    key_words = self.key_le.text().strip()#关键词
    target = self.target_le.text().strip()#保存路径
    self.fetchData(key_words, target)
  def fetchData(self, key_words, target):
    print('关键词:', key_words)
    print('保存路径:', target)
    browser = webdriver.Chrome()
    browser.get('https://z.vanmmall.com/?from=singlemessage&isappinstalled=0')
    browser.find_element_by_id('kw').send_keys(key_words)
    search = browser.find_element_by_id('btn_search')
    search.click()
    time.sleep(2)
    while not self.isElementExist(browser, '.dropload-noData'):
      browser.execute_script("window.scrollTo(0,document.body.scrollHeight);")
      time.sleep(1)
    time.sleep(2)
    date = browser.find_elements_by_css_selector('td.date')
    shop = browser.find_elements_by_css_selector('td.shop div.name')
    msg = browser.find_elements_by_css_selector('td.msg')
    wechat = browser.find_elements_by_css_selector('td.call div button.wx')
    mobile = browser.find_elements_by_css_selector('td.call div button.phone')
    data = []
    for index in range(len(date)):
      date_text = date[index].get_attribute('textContent')
      shop_text = shop[index].get_attribute('textContent')
      msg_text = msg[index].get_attribute('textContent')
      wechat_text = wechat[index].get_attribute("data-msg")
      mobile_text = mobile[index].get_attribute("data-msg")
      data.append([date_text, shop_text, msg_text, wechat_text, mobile_text])
    print(data)
    self.saveDataToExcel(data, target)
  def isElementExist(self, browser, element):
    flag=True
    try:
      browser.find_element_by_css_selector(element)
      return flag
    except:
      flag=False
      return flag
  def saveDataToExcel(self, data, target):
    data_df = pd.DataFrame(data)
    data_df.columns = ['日期','商家','产品信息+报价','微信号','手机号']
    writer = pd.ExcelWriter(target)
    data_df.to_excel(writer, float_format='%.5f')
    writer.save()

if __name__=="__main__":
  app = QApplication(sys.argv)
  ex = upload()
  sys.exit(app.exec_())