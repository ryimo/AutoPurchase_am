# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import threading
import os
import xlwings as xw
import datetime
import time

# グローバル変数
UserName_amazon = ""
Password_amazon = ""
IsHideChrome = False

# 実行フォルダパスを取得
ExecDir = os.path.dirname(__file__)

# ログファイルの準備
d = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
logFileName = os.path.join(ExecDir, R'log\log_' + d + '.txt')

##############################
#   エクセル書き込み用関数    #
##############################
def write_Excel(productName, price):

    wb = xw.Book.caller()
    ws = wb.sheets('購入履歴')

    # 空白行番号を取得
    RowNum_rireki = 5
    while True:
        if ws.range((RowNum_rireki, 3)).value is None:
            break
        else:
            RowNum_rireki += 1

    # 情報を記載
    ws.range((RowNum_rireki, 3)).value = productName
    ws.range((RowNum_rireki, 4)).value = price
    ws.range((RowNum_rireki, 5)).value = datetime.datetime.now().strftime('%Y年%m月%d日')

    wb = None
    ws = None

##############################
#     スレッド作成用クラス    #
##############################
class AmazonThreading(threading.Thread):

    # コンストラクタ
    def __init__(self, thread_name, productName, url, minPrice, maxPrice):
        self.thread_name = str(thread_name)
        self.productName = productName
        self.url = url
        self.minPrice = int(minPrice)
        self.maxPrice = int(maxPrice)

        threading.Thread.__init__(self)

    def __str__(self):
        return self.thread_name

    # 実処理
    def run(self):
        try:
            #ChromeDriver準備
            chromDirverPath = os.path.join(ExecDir, R"chromedriver\chromedriver.exe")
            if os.path.isfile(chromDirverPath) == False:
                f = open(logFileName, 'a')
                f.write("chromedriver.exeが見つかりません\n")
                f.write("想定パス： "+ chromDirverPath + "\n")
                f.close()
                return
            
            options = Options()
            if IsHideChrome:
                options.add_argument('--headless')
            driver = webdriver.Chrome(executable_path=chromDirverPath, chrome_options=options)
            wait = WebDriverWait(driver, 10)
            
            # [S] Amazonへのログイン
            driver.get("https://www.amazon.co.jp")
            # ユーザー名を入力
            driver.get(driver.find_element_by_xpath("//*[@id='nav-flyout-ya-signin']/a").get_attribute("href"))
            elm = driver.find_element_by_xpath("//*[@id='ap_email']")
            elm.clear()
            elm.send_keys(UserName_amazon)
            elm.send_keys(Keys.ENTER)
            
            # パスワードを入力
            elm = driver.find_element_by_xpath("//*[@id='ap_password']")
            elm.clear()
            elm.send_keys(Password_amazon)
            elm.send_keys(Keys.ENTER)
            # [E] Amazonへのログイン

            # ログイン時に電話番号連携について聞かれることがあるので、そのページが出た場合はスキップする
            if len(driver.find_elements_by_id('ap-account-fixup-phone-skip-link')) != 0:
                driver.find_element_by_id('ap-account-fixup-phone-skip-link').Click

            # ホーム画面がロードするまで待機（待機しないとURLの移動ができなかったため）
            wait.until(EC.visibility_of_element_located((By.ID, 'navBackToTop')))
            
            # 指定されたURLに移動
            driver.get(self.url)
            wait.until(EC.presence_of_all_elements_located)
            
            # 最新商品順に並び替え（並び替え選択肢に'最新商品'がなければスキップ）
            selector = Select(driver.find_element_by_id('s-result-sort-select'))
            for op in selector.options:
                if op.text == "最新商品":
                    selector.select_by_visible_text("最新商品")
        
            # ループ処理
            quit_flg = False
            while quit_flg == False:

                # ページ内の商品情報divタグを取得
                wait.until(EC.visibility_of_all_elements_located((By.TAG_NAME, 'h2')))
                # 商品情報divタグを判定
                divTag_css_selector=''
                if len(driver.find_elements_by_css_selector('div.sg-col-4-of-12.s-result-item.s-asin.sg-col-4-of-16.sg-col.sg-col-4-of-20')) != 0:
                    divTag_css_selector = 'div.sg-col-4-of-12.s-result-item.s-asin.sg-col-4-of-16.sg-col.sg-col-4-of-20'
                elif len(driver.find_elements_by_css_selector('div.s-result-item.s-asin.sg-col-0-of-12.sg-col-16-of-20.sg-col.sg-col-12-of-16')) != 0:
                    divTag_css_selector = 'div.s-result-item.s-asin.sg-col-0-of-12.sg-col-16-of-20.sg-col.sg-col-12-of-16'
                
                divTagElements = driver.find_elements_by_css_selector(divTag_css_selector)
                
                for divTagEle in divTagElements:
                    
                    # 取得した商品情報divタグに含まれる商品名を取得
                    h2Tag_css_selector=''
                    if len(divTagEle.find_elements_by_css_selector('span.a-size-base-plus.a-color-base.a-text-normal')) != 0:
                        h2Tag_css_selector = 'span.a-size-base-plus.a-color-base.a-text-normal'
                    elif len(divTagEle.find_elements_by_css_selector('span.a-size-medium.a-color-base.a-text-normal')) != 0:
                        h2Tag_css_selector = 'span.a-size-medium.a-color-base.a-text-normal'

                    if h2Tag_css_selector != '':
                        productNameElement = divTagEle.find_element_by_css_selector(h2Tag_css_selector)
                        productNameElementText = str(productNameElement.text)

                        # 指定された商品名が含まれているか確認
                        if self.productName in productNameElementText:
                            
                            if len(divTagEle.find_elements_by_class_name('a-price-whole')) != 0:
                                priceText = divTagEle.find_element_by_class_name('a-price-whole').text
                                price_webPage = int(priceText.replace('￥', '').replace(',', ''))
                            
                                # 価格判定
                                if self.minPrice <= price_webPage and price_webPage <= self.maxPrice:
                                    
                                    # 商品ページに移動
                                    if len(divTagEle.find_elements_by_tag_name("a")) != 0:
                                        divTagEle.find_element_by_tag_name("a").click()
                                        driver.switch_to.window(driver.window_handles[1])
                                        wait.until(EC.visibility_of_element_located((By.ID, 'buy-now-button')))
                                                        
                                        # 購入処理
                                        if len(driver.find_elements_by_id("buy-now-button")) != 0:
                                            driver.find_element_by_id("buy-now-button").click()

                                            try:
                                                # 商品ページに小窓が出てくるパターン（おそらく購入だとこっち）
                                                
                                                wait.until(EC.visibility_of_element_located((By.ID, 'turbo-checkout-iframe')))
                                                if len(driver.find_elements_by_id('turbo-checkout-iframe')) != 0:
                                                    # フレーム移動
                                                    driver.switch_to.frame(driver.find_element_by_id('turbo-checkout-iframe'))

                                                    if len(driver.find_elements_by_id('turbo-checkout-pyo-button')) != 0:
                                                        # 購入
                                                        driver.find_element_by_id('turbo-checkout-pyo-button').click()

                                                        # 購入履歴を記載
                                                        write_Excel(productNameElementText, price_webPage)
                                                    
                                                    #　親フレームに戻す
                                                    driver.switch_to.default_content()
                                            except:
                                                # 商品ページから移動して購入を確定するパターン（おそらく予約だとこっち）

                                                # プライム会員の誘いをスキップ  
                                                if len(driver.find_elements_by_xpath('//*[@id="primeAutomaticPopoverAdContent"]/div/div[1]/div[1]/a')) != 0:
                                                    driver.find_element_by_xpath('//*[@id="primeAutomaticPopoverAdContent"]/div/div[1]/div[1]/a').click()

                                                if len(driver.find_elements_by_css_selector('input.a-button-text.place-your-order-button')) != 0:
                                                    # 購入
                                                    driver.find_element_by_css_selector('input.a-button-text.place-your-order-button').click()

                                                    # 購入履歴を記載
                                                    write_Excel(productNameElementText, price_webPage)                                        

                                            quit_flg = True
                                    
                                    # 商品ページを閉じる
                                    driver.close()
                                    driver.switch_to.window(driver.window_handles[0]) 

                                    if quit_flg:
                                        break

                # ループ終了判定
                if quit_flg == False:

                    LOOP_NUM = 12
                    i = 1
                    # あまり連続して更新をかけると怒られそうなので、1分(5秒×12回)待機
                    while i <= LOOP_NUM:

                        # 10秒待機ごとに終了フラグを確認
                        # エクセルが終了通知用のファイルを作成することで、終了指示を検知する
                        if(os.path.isfile(os.path.join(ExecDir,'loop_end.txt'))):
                            quit_flg = True
                            break
                        else:
                            time.sleep(5)
                            i += 1

                    if i > LOOP_NUM:
                        # 指定されたURLを再ロードして一覧更新
                            driver.refresh()
                            

        except Exception as e:
            # 例外発生時に内容をファイルに出力する
            f = open(logFileName, 'a')
            f.write('=== エラー内容 ===\n')
            f.write('type:' + str(type(e)) + '\n')
            f.write('args:' + str(e.args) + '\n')
            f.write('message:' + e.message + '\n')
            f.write('e自身:' + str(e) + '\n')
            f.close()
        
        finally:
            # 後片付け
            driver.quit()

################################
#  エクセルから呼び出される関数  #
################################
def callFromExcel():
    
    # [S] Excelファイルから情報読み取り
    wb = xw.Book.caller()
    
    # [S] ツール操作シートの情報取得
    # エクセル上の実行ステータス変更
    ws = wb.sheets('ツール操作')
    ws.range((3, 6)).value = '実行中'

    # Amazonへのログイン情報を取得
    global UserName_amazon
    global Password_amazon
    UserName_amazon = ws.range((9, 3)).value
    Password_amazon = ws.range((10, 3)).value
    
    global IsHideChrome
    if int(ws.range((9, 4)).value) == 2:
        IsHideChrome = True
    # [E] ツール操作シートの情報取得

    # [S] 購入希望商品一覧シートの情報取得
    # 商品情報を取得
    ws = wb.sheets('購入希望商品一覧')
    RowNum = 6
    excelInfoList = []
    while True:
        productName_excel = ws.range((RowNum, 4)).value
        url_excel = ws.range((RowNum, 5)).value
        minPrice_excel = ws.range((RowNum, 6)).value
        maxPrice_excel = ws.range((RowNum, 7)).value
        
        if productName_excel is None or url_excel is None or minPrice_excel is None or maxPrice_excel is None:
            break
        else:
            excelInfoList.append([productName_excel, url_excel, minPrice_excel, maxPrice_excel])
        
        RowNum += 1
    # [E] 購入希望商品一覧シートの情報取得
    
    # エクセルを手放す（これをやらないと別スレッドからの参照ができない）
    wb = None
    ws = None

    # [E] Excelファイルから情報読み取り

    # 購入希望商品分スレッドを作成し、処理実行
    thList=[]
    for i in range(len(excelInfoList)):
        thread = AmazonThreading(thread_name=i, productName=excelInfoList[i][0], url=excelInfoList[i][1], minPrice=excelInfoList[i][2], maxPrice=excelInfoList[i][3])
        thread.start()
        thList.append(thread)
    
    for thread in thList:
        thread.join()
    
    # エクセル上の実行ステータス変更
    wb = xw.Book.caller()
    ws = wb.sheets('ツール操作')
    ws.range((3, 6)).value = '停止'

    # 後片付け
    wb = None
    ws = None
    