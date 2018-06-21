# -*- coding: utf-8 -*-
import xlrd
import os
import xlsxwriter
import math
import time
from selenium import webdriver
from PIL import Image
import glob
from ftplib import FTP

## ExcelからURLを取得
book = xlrd.open_workbook('url.xlsx')
sheet = book.sheet_by_index(0)

## 取得した情報を保存しておくための入れ物
IDLIST = []
TITLELIST = []
URLLIST = []

## IDを取得
for row in range(sheet.nrows):
    ## セルの値（URL）を取得
    ID = sheet.cell(row, 0).value
    IDLIST.append(ID)

## タイトルを取得
for row in range(sheet.nrows):
    ## セルの値（URL）を取得
    TITLE = sheet.cell(row, 1).value
    TITLELIST.append(TITLE)

## PCキャプチャ
## ブラウザを起動（ページ全体をキャプチャしてくれるためSafariを使用）
driver = webdriver.Chrome()
driver.maximize_window()

##  URLを取得（キャプチャ保存）
for row in range(sheet.nrows):
    ## セルの値（URL）を取得
    URL = sheet.cell(row, 2).value
    URLLIST.append(URL)

    ## 画面遷移
    driver.get(URL)

    ## 遷移直後だと崩れた状態でスクショされる可能性があるため、1秒待機
    time.sleep(5)

    ## 画面キャプチャを保存
    FILENAME = os.path.join(os.path.dirname(os.path.abspath(__file__)),'png', IDLIST[row] + '.png')
    driver.save_screenshot(FILENAME)

## ブラウザを閉じる
driver.quit()

## 幅500pxにリサイズ

## フォルダ内のファイル取得

IMGLIST = glob.glob(os.path.dirname(os.path.abspath(__file__)) +'/png/*')
for row in range(len(IMGLIST)):
    img = Image.open(IMGLIST[row])
    img_resize = img.resize((500,302),Image.BICUBIC)
    img_resize.save(IMGLIST[row])


## FTPでサーバーにアップ
DUP = False;
ftp = FTP(
    "ftp-tds-bdt.heteml.net",
    "tds-bdt",
    passwd="tds-bdt52257775"
)
ftp.cwd("/web/research/")
ITEMS = ftp.nlst(".")
txt = open('directory.txt')
DIRECTORY = txt.readline()

for i in range(len(ITEMS)):
    if DIRECTORY == ITEMS[i]:
        DUP = True;

if DUP == False:
    ftp.mkd(DIRECTORY)
    for i in range(len(IMGLIST)):
        fp = open(IMGLIST[i], 'rb')
        ftp.storbinary('STOR /web/research/' + DIRECTORY + "/" + IDLIST[i] + '.png',fp)
else:
    print "ディレクトリ名が重複しています。"

print 'filepath--> https://www.tds-g.biz/research/' + DIRECTORY + "/" + IDLIST[0] + '.png'
txt.close()
ftp.quit
