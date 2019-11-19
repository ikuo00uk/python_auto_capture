import xlrd
import os
import xlsxwriter
import math
import time
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# 実行スクリプト

options = Options()
# Chrome v75からW3CモードがデフォルトONになったので指定する
options.add_experimental_option('w3c', False)
options.add_argument('--headless')
options.add_argument('--dns-prefetch-disable')


## -------------------------------------
## ExcelのURLからスクショを保存
## -------------------------------------

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

## UA判定が必要な場合は上記を削除し、以下を使用
# USER_AGENT = "Mozilla/5.0 (iPhone; CPU iPhone OS 11_0 like Mac OS X) AppleWebKit/604.1.38 (KHTML, like Gecko) Version/11.0 Mobile/15A356 Safari/604.1"
# driver = webdriver.PhantomJS(desired_capabilities={'phantomjs.page.settings.userAgent': USER_AGENT})

def capture(device):
  config = getDeviceConfig(device)
  # Chrome Driver 起動
  driver = webdriver.Chrome(options=options)
  print('START CAPTURE FILE')
  ##  URLを取得（キャプチャ保存）
  for row in range(sheet.nrows):
    try:
      URL = sheet.cell(row, 2).value
      _idlist = str(IDLIST[row])
      print(URL)
      URLLIST.append(URL)
      ## 画面遷移
      driver.get(URL)

      # スクリーンサイズ設定
      page_height = driver.execute_script('return document.body.scrollHeight')
      driver.set_window_size(config['windowWidth'], page_height)

      time.sleep(2)
      driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
      ## 画面キャプチャを保存
      FILENAME = os.path.join(saveDir, _idlist + config['imgSuffix'] + '.png')
      driver.save_screenshot(FILENAME)
    except:
      print('error')
      pass
  # Chrome Driver 終了
  driver.quit()
  print('DONE CAPTURE')

## -------------------------------------
## 保存したスクショをExcelに添付
## -------------------------------------

def createExcelFile(condition):
    ## 添付した数をカウントするための変数
    count = 0

    for i in range(math.ceil((row+1)/10)):
        ## 画像添付用のExcelを作成
        workbook = xlsxwriter.Workbook(saveDir + '/' + 'capture' + str(i+1) + '.xlsx')

        ## 残り添付数を計算
        if (row+1)-count >= 10:
            num = 10
        else:
            num = (row+1)%10
        for j in range(num):
            _idlist = str(IDLIST[count])
            _title = _idlist + '_' + TITLELIST[count]
            _title_truncated = _title[:31]
            
            ## シートを追加
            worksheet = workbook.add_worksheet(_title_truncated)

            ## 対象ページの情報を記載
            worksheet.write('A1', _idlist)
            worksheet.write('A2', _title_truncated)
            worksheet.write('A3', URLLIST[count])

            ## 画像を添付
            if(condition == 1 or condition == 2):
              IMAGE_PC = saveDir + '/' + _idlist + '_pc.png'
              worksheet.insert_image('H4', IMAGE_PC, {'x_scale': 0.3, 'y_scale': 0.3})

            if(condition == 1 or condition == 3):
              IMAGE_SP = saveDir + '/' + _idlist + '_sp.png'
              worksheet.insert_image('B4', IMAGE_SP, {'x_scale': 0.45, 'y_scale': 0.45})
            ## 添付するごとにカウントを増やす
            count = count+1

    ## 10シートごとにExcelを閉じる
    workbook.close()

print('DONE CREATE EXCEL FILE')

#
# キャプチャ対象のデバイス確認
#
def confirmCaptureDevice():
  yes = ['y', 'ye', 'yes']
  no = ['n', 'no']
  choice = input("PCとSPそれぞれキャプチャを取得しますか？ [y/N]: ").lower()
  if choice in yes:
      return 1
  elif choice in no:
    choice2 = input("PCのみキャプチャを取得しますか？ NOの場合SPのみ保存します [y/N]: ").lower()
    if choice2 in yes:
      return 2
    else:
      return 3
#
# 保存用ディレクトリの作成
#
def getCaptureDir():
  path = 'capture'
  isDir = os.path.isdir(path)
  if(not isDir):
    os.mkdir(path)
  else:
    shutil.rmtree(path)
    os.mkdir(path)
  print('INITIALIZE CAPTURE SAVE DIRECTORY')
  return os.path.dirname(os.path.abspath(__file__))+ '/' + path

#
# キャプチャデバイス用変数管理
# windowWidth *2 でレンダリングされる
#
def getDeviceConfig(device = 'pc'):
  config = {
    'pc': {
      'windowWidth' : 600, 
      'imgSuffix' : '_pc'
    },
    'sp': {
      'windowWidth' : 375,
      'imgSuffix' : "_sp"
    }
  }
  if(device in config):
    return config[device]
  else:
    raise ValueError("該当するキーが存在しません")


# 格納するディレクトリを生成
saveDir = getCaptureDir()

# キャプチャ実行条件
captureCond = confirmCaptureDevice()

def execute():
  if(captureCond == 1 or captureCond == 2):
    capture('pc')
  if(captureCond == 1 or captureCond == 3):
    capture('sp')
  createExcelFile(captureCond)
  print('FINISH')

execute()
