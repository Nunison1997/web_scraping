#!/usr/bin/python
# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from selenium.common.exceptions import TimeoutException
import traceback
import pyperclip
import gspread
from oauth2client.service_account import ServiceAccountCredentials 
from bs4 import BeautifulSoup

# ChromeOptionsの設定
options = Options()
options.add_argument(r'--user-data-dir=C:\Users\User\AppData\Local\Google\Chrome\User Data')
options.add_argument('--profile-directory=Profile 4')
options.add_argument('--no-sandbox')
# options.add_argument("--headless")
options.add_argument('--start-maximized')


# WebDriverのインスタンスを作成
driver = webdriver.Chrome(options=options)
url = "https://www.google.com/?hl=ja"
driver.get(url)

url_input = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'q')))
url_input.send_keys("google search console")
url_input.send_keys(Keys.RETURN)


url_click = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Google Search Console")))
url_click.click()

url_click = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "今すぐ開始")))
url_click.click()

try: 
    result_link = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='search-console/performance/search-analytics']")))
    result_link.send_keys(Keys.RETURN)
except TimeoutException as e:
    print("TimeoutException")
except Exception as e:
    print(traceback.format_exc())

#過去3カ月
term_change = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/c-wiz[3]/c-wiz/div/div[1]/div[2]/div/div/div[2]/div/div")))
term_change.send_keys(Keys.RETURN)

#過去28日
calander = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//label[@for='c12']")))
calander.click()

#適用
apply = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[6]/div[2]/div/div[2]/div/div/div[2]/button")))
apply.send_keys(Keys.RETURN)

time.sleep(5)#これがないととってくれない
#クリック回数をget beautifulsoupを使用
current_url = driver.current_url
html = driver.page_source.encode('utf-8')
soup = BeautifulSoup(html, 'html.parser')
element = soup.findAll("div", class_= "nnLLaf vtZz6e")

#dfに要素追加
df = pd.DataFrame({'月間検索流入数':[element[4].text]})

#ユニゾンキャリアに変更
change_link = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/c-wiz[1]/div/gm-coplanar-drawer/div/div/div[2]/div[2]/c-wiz/div/div/div[2]/div[1]/div[2]/div[1]/div/div[1]/input")))
change_link.click()

#unisonに移す
change_link = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/c-wiz[1]/div/gm-coplanar-drawer/div/div/div[2]/div[2]/c-wiz/div/div/div[2]/div[2]/div/div[1]/div[1]")))
change_link.click()

#指名検索(ユニゾンキャリア)

from selenium.webdriver.common.action_chains import ActionChains
time.sleep(3)
# ウィンドウのサイズを取得
window_size = driver.get_window_size()
window_width = window_size['width']
window_height = window_size['height']

# ウィンドウの中央座標を計算
center_x = window_width // 2
center_y = window_height // 2
# 中央をクリック
action = ActionChains(driver)
action.move_by_offset(center_x+30, 150).click().perform()

key_word = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/c-wiz[4]/div/div/div/span[1]'))
    )
key_word.send_keys(Keys.RETURN)

# 入力
input_name = WebDriverWait(driver, 30).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[6]/div[2]/div/div[1]/div/div/div[2]/span/div/div[2]/label'))
)
input_name.send_keys("ユニゾンキャリア")
input_name.send_keys(Keys.RETURN)

time.sleep(5)
#クリック数と表示回数を取得
driver.get(driver.current_url)
html = driver.page_source.encode('utf-8')
soup = BeautifulSoup(html, 'html.parser')
element = soup.findAll("div", class_= "nnLLaf vtZz6e")

df['指名検索表示回数(ユニゾンキャリア)'] = [element[1].text] 
df['指名検索クリック数(ユニゾンキャリア)'] = [element[0].text]

time.sleep(3)

action.click().perform()

#指名検索(ユニゾンテクノロジー)
input_name = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[6]/div[2]/div/div[1]/div/div/div[2]/span/div/div[2]/label/span[2]/input'))
)
from selenium.webdriver.common.keys import Keys
time.sleep(5)
input_name.send_keys(Keys.CONTROL + 'a', Keys.DELETE)
time.sleep(3)
input_name.send_keys("ユニゾンテクノロジー")
input_name.send_keys(Keys.RETURN)

time.sleep(5)
#クリック数と表示回数を取得
driver.get(driver.current_url)
html = driver.page_source.encode('utf-8')
soup = BeautifulSoup(html, 'html.parser')
element = soup.findAll("div", class_= "nnLLaf vtZz6e")

df['指名検索表示回数(ユニゾンテクノロジー)'] = [element[1].text] 
df['指名検索クリック数(ユニゾンテクノロジー)'] = [element[0].text]

time.sleep(3)


#########ここからはanalytics########

#urlの取得
url = "https://analytics.google.com/analytics/web/?hl=ja&pli=1#/p370990079/reports/intelligenthome"
driver.get(url)
time.sleep(5)

#必要に応じてある場所をクリック(GA4に移行後は不要)
# action.move_by_offset(0, 0).click().perform()
# time.sleep(5)

#レポートのスナップショットを表示
report = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/ga-hybrid-app-root/ui-view-wrapper/div/app-root/div/div/div[1]/ga-report-container/div/div/div/report-view/div/ga-intelligent-home-report/div/div/div/div[1]/div/ga-card-list/div/ga-card[1]/xap-card/xap-card-footer/ga-view-link/button")))
report.send_keys(Keys.RETURN)

time.sleep(5)
#平均エンゲージメント時間を取得
driver.get(driver.current_url)
html = driver.page_source.encode('utf-8')
soup = BeautifulSoup(html, 'html.parser')
element = soup.findAll("div", class_= "value ng-star-inserted")
df['平均エンゲージメント時間'] = [element[2].text]

time.sleep(3)
#表示回数
target = driver.find_element(By.XPATH, "/html/body/ga-hybrid-app-root/ui-view-wrapper/div/app-root/div/div/ui-view-wrapper/div/ga-report-container/div/div/div/report-view/ui-view-wrapper/div/ui-view/ga-dashboard-report/div/div/div/ga-card-list/div/ga-card[9]/xap-card/xap-card-footer/ga-view-link/button")

action.move_to_element(target).perform()
time.sleep(2)
target.click()
time.sleep(5)

#beutifulsoupで表示回数を取得
driver.get(driver.current_url)
html = driver.page_source.encode('utf-8')
soup = BeautifulSoup(html, 'html.parser')
element = soup.findAll("div", class_= "inline ga-value-info-target total-value ng-star-inserted")
df["PV数"] = [element[0].text]

#集客までカーソル移動
repeat = driver.find_element(By.XPATH, "/html/body/ga-hybrid-app-root/ui-view-wrapper/div/app-root/div/xap-deferred-loader-outlet/ga-left-nav2/ga-nav2/ga-secondary-nav/mat-tree/mat-tree-node[22]/ga-secondary-nav-item/button")

action.move_to_element(repeat).perform()

#リピーター数
repeat = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/ga-hybrid-app-root/ui-view-wrapper/div/app-root/div/xap-deferred-loader-outlet/ga-left-nav2/ga-nav2/ga-secondary-nav/mat-tree/mat-tree-node[22]/ga-secondary-nav-item/button")))
repeat.click()
time.sleep(5)

#beautifulsoup
driver.get(driver.current_url)
html = driver.page_source.encode('utf-8')
soup = BeautifulSoup(html, 'html.parser')
element = soup.findAll("div", class_= "value ng-star-inserted")
df["維持率"] = [element[1].text]

#LPを中途用に切り替え
lp = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/ga-hybrid-app-root/ui-view-wrapper/div/app-root/xap-deferred-loader-outlet/ga-header/header/gmp-header/gmp-universal-picker/button")))
lp.send_keys(Keys.RETURN)

#中途用LP
lp = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "中途用LP")))
lp.click()

time.sleep(3)
#ホームに戻る
home = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/ga-hybrid-app-root/ui-view-wrapper/div/app-root/div/xap-deferred-loader-outlet/ga-left-nav2/ga-nav2/ga-primary-nav/mat-nav-list[1]/ga-primary-nav-item[1]/a")))
home.click()

# #過去7日間を変更
# past_seven = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/ga-hybrid-app-root/ui-view-wrapper/div/app-root/div/div/div[1]/ga-report-container/div/div/div/report-view/div/ga-intelligent-home-report/div/div/div/div[1]/div/ga-card-list/div/ga-card[1]/xap-card/xap-card-footer/ga-date-range-selector/ga-date-range-picker-v2/div[2]/button")))
# driver.execute_script("arguments[0].click();", past_seven)

# #過去28日
# past_28 = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[6]/div[2]/div/reach-date-range-picker-content/xap-card/div/reach-date-range-calendar/div/reach-calendar-presets-menu/div/div[8]")))
# past_28.click()

time.sleep(5)

#beautifulsoup 
driver.get(driver.current_url)
html = driver.page_source.encode('utf-8')
soup = BeautifulSoup(html, 'html.parser')
element = soup.findAll("div", class_= "value ng-star-inserted")
df["CV数(中途用)"] = [element[1].text]
df["新規ユーザー数(中途用)"] = [element[3].text]

#転職用LP
newwer = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/ga-hybrid-app-root/ui-view-wrapper/div/app-root/xap-deferred-loader-outlet/ga-header/header/gmp-header/gmp-universal-picker/button")))
newwer.click()

#切り替え
lp = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "転職用LP（経＆未経験）")))
lp.click()
time.sleep(5)

#beautifulsoup 
driver.get(driver.current_url)
html = driver.page_source.encode('utf-8')
soup = BeautifulSoup(html, 'html.parser')
element = soup.findAll("div", class_= "value ng-star-inserted")
df["CV数(転職用)"] = [element[1].text]
df["新規ユーザー数(転職用)"] = [element[3].text]

#HPに切り替え
hp = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/ga-hybrid-app-root/ui-view-wrapper/div/app-root/xap-deferred-loader-outlet/ga-header/header/gmp-header/gmp-universal-picker/button")))
hp.click()

lp = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "ユニゾンキャリア_HP - GA4")))
lp.click()
time.sleep(5)

#beautifulsoup 
driver.get(driver.current_url)
html = driver.page_source.encode('utf-8')
soup = BeautifulSoup(html, 'html.parser')
element = soup.findAll("div", class_= "value ng-star-inserted")
df["CV数(HP)"] = [element[1].text]

#レポートのスナップショットを表示
report = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/ga-hybrid-app-root/ui-view-wrapper/div/app-root/div/div/div[1]/ga-report-container/div/div/div/report-view/div/ga-intelligent-home-report/div/div/div/div[1]/div/ga-card-list/div/ga-card[1]/xap-card/xap-card-footer/ga-view-link/button")))
report.click()

time.sleep(5)
#新規ユーザー数
driver.get(driver.current_url)
html = driver.page_source.encode('utf-8')
soup = BeautifulSoup(html, 'html.parser')
element = soup.findAll("div", class_= "value ng-star-inserted")
df["新規ユーザー数(HP)"] = [element[1].text]

#HPのPV数を取得するための操作

clickElement = driver.find_element(By.XPATH, "/html/body/ga-hybrid-app-root/ui-view-wrapper/div/app-root/div/div/ui-view-wrapper/div/ga-report-container/div/div/div/report-view/ui-view-wrapper/div/ui-view/ga-dashboard-report/div/div/div/ga-card-list/div/ga-card[9]/xap-card/xap-card-footer/ga-view-link/button")
driver.execute_script("arguments[0].click();", clickElement)
time.sleep(5)

#Beautifulsoup
driver.get(driver.current_url)
html = driver.page_source.encode('utf-8')
soup = BeautifulSoup(html, 'html.parser')
element = soup.findAll("div", class_= "inline ga-value-info-target total-value ng-star-inserted")
df["HPのPV数"] = [element[0].text]

#HPの回遊率を確認
circulation_percent = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/ga-hybrid-app-root/ui-view-wrapper/div/app-root/div/xap-deferred-loader-outlet/ga-left-nav2/ga-nav2/ga-secondary-nav/mat-tree/mat-tree-node[9]/ga-secondary-nav-item/button")))
circulation_percent.click()

#オーディエンス
ordience = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/ga-hybrid-app-root/ui-view-wrapper/div/app-root/div/xap-deferred-loader-outlet/ga-left-nav2/ga-nav2/ga-secondary-nav/mat-tree/mat-tree-node[12]/ga-secondary-nav-item/button")))

ordience.click()

time.sleep(10)
#Beautfulsoup
driver.get(driver.current_url)
html = driver.page_source.encode('utf-8')
soup = BeautifulSoup(html, 'html.parser')
element = soup.find_all("div", class_= "inline ga-value-info-target total-value ng-star-inserted")
df["HPの回遊率"] = [element[3].text]

#以下、dfの形式を変換する操作
import re

# 正規表現を使って数字部分を抽出
def extract_time(time_str):
    match = re.match(r"(\d+) 分 (\d+) 秒", time_str)
    if match:
        minutes = int(match.group(1))
        seconds = int(match.group(2))
        total_seconds = minutes * 60 + seconds
        return total_seconds
    else:
        return None

# 平均エンゲージメント時間を数値に変換
df["平均エンゲージメント時間"] = df["平均エンゲージメント時間"].apply(extract_time)

#万を数字に変換する
def thousand(number):
    match = re.match(r"([+-]?(?:\d+\.?\d*|\.\d+))万", number)
    if match:
        num = float(match.group(1))
        output = int(num * 10000)
        return output
    else:
        return None

df['月間検索流入数'] = df['月間検索流入数'].apply(thousand)
df['維持率'] = df['維持率'].apply(thousand)

#文字列から数字に変換
df["指名検索表示回数(ユニゾンキャリア)"] = df["指名検索表示回数(ユニゾンキャリア)"].str.replace(',', '').astype(float)
df["指名検索クリック数(ユニゾンキャリア)"] = df["指名検索クリック数(ユニゾンキャリア)"].str.replace(',', '').astype(float)
df["指名検索表示回数(ユニゾンテクノロジー)"] = df["指名検索表示回数(ユニゾンテクノロジー)"].str.replace(',', '').astype(float)
df["指名検索クリック数(ユニゾンテクノロジー)"] = df["指名検索クリック数(ユニゾンテクノロジー)"].str.replace(',', '').astype(float)
df["CV数(中途用)"] = df["CV数(中途用)"].str.replace(',', '').astype(float)
df["CV数(転職用)"] = df["CV数(転職用)"].str.replace(',', '').astype(float)
df["新規ユーザー数(中途用)"] = df["新規ユーザー数(中途用)"].str.replace(',', '').astype(float)
df["新規ユーザー数(転職用)"] = df["新規ユーザー数(転職用)"].str.replace(',', '').astype(float)

df["HPのPV数"] = df["HPのPV数"].str.replace(',', '').astype(float)
df["PV数"] = df["PV数"].str.replace(',', '').astype(float)
df["CV数(HP)"] = df["CV数(HP)"].str.replace(',', '').astype(float)

###ここからはdfの値を用いて計算する###
#回遊率の計算
df["回遊率"] = [100 * df["PV数"] / df["月間検索流入数"]]

df["回遊率"] = df["回遊率"].iloc[0]
#HP回遊数
df["HP回遊数"] = df["HPのPV数"] - (df["指名検索クリック数(ユニゾンキャリア)"] + df["指名検索クリック数(ユニゾンテクノロジー)"])
#メディア率
df["メディア率"] = 100 * df["HP回遊数"] / df["HPのPV数"]

#dfをペーストするために作り替える
df = df.reindex(columns=['月間検索流入数','PV数','回遊率', '平均エンゲージメント時間','維持率','指名検索表示回数(ユニゾンキャリア)','指名検索クリック数(ユニゾンキャリア)','指名検索表示回数(ユニゾンテクノロジー)','指名検索クリック数(ユニゾンテクノロジー)','新規ユーザー数(中途用)','新規ユーザー数(転職用)','HPのPV数','HP回遊数','HPの回遊率','メディア率','CV数(中途用)','CV数(転職用)','CV数(HP)'])


#最大500列で表示
pd.set_option('display.max_columns',500)
print(df)
# str_output1 = df.iloc[:,0:2].to_string(index=False)
# str_output2 = df.iloc[:,2:5].to_string(index=False)
# str_output3 = df.iloc[:,5:19].to_string(index=False)

# print(str_output1)
# print(str_output2)
# print(str_output3)

#行と列を指定して要素のみ抽出

#listにdfの要素を入れる

# スプレッドシートに書き込むリストを作成
list1 = df.iloc[0, 0:2].astype(float).tolist()
list2 = df.iloc[0, 2:5].astype(float).tolist()
list3 = df.iloc[0, 5:].astype(float).tolist()

round_list1 = [round(value,2) for value in list1]
round_list2 = [round(value,2) for value in list2]
round_list3 = [round(value,2) for value in list3]


###ここからスプシに移す###
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

#認証情報設定
#ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
credentials = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\User\Desktop\web_scraping\anal-410708-38581d77a165.json", scope)

#OAuth2の資格情報を使用してGoogle APIにログインします。
gc = gspread.authorize(credentials)

#共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
SPREADSHEET_KEY = '1QmikDXr4JYplvaqt9m_1t0LFj3vQmkqN3q2z7dHxBew'

#共有設定したスプレッドシートの特定のシートを開く
worksheet = gc.open_by_key(SPREADSHEET_KEY).worksheet("KPIデータ")

#1番左にある文字を認識する
colum_length = len(worksheet.row_values(3))#3行目の最も右

#数字からアルファベットに変換するコード
#例　11→A1,
Al = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
num = colum_length

def Alphabet(num):
    # num = colum_length+1
    if ((num+1)//26) == 0:
        First = Al[(num+1)%26 -1]
        Sentence = First
        return Sentence

    elif ((num+1) >=1) and (num+1)//26**2==0:
        First = Al[(num+1)%26 -1]
        Second = Al[(num+1)//26-1]
        Sentence = Second + First
        return Sentence
    
    else:
        Third = Al[(num+1)//26**2-1]
        Second = Al[((num+1)%26**2)//26-1]
        First = Al[((num+1)%26**2)%26 -1]
        Sentence = Third + Second + First 
        return Sentence

#3列の取得
start_low1 = "3:"
end_low1 = "4"
first_colum= str(Alphabet(num)+start_low1+Alphabet(num)+end_low1)#第1sector

start_low2 = "6:"
end_low2 = "8"
second_colum = str(Alphabet(num)+start_low2+Alphabet(num)+end_low2) #第2sector

start_low2 = "10:"
end_low2 = "22"
third_colum = str(Alphabet(num)+start_low2+Alphabet(num)+end_low2) #第3sector

# str_list1 = ", ".join(map(str, list1))

worksheet.update(range_name=first_colum, values=[[value]for value in round_list1])
worksheet.update(range_name=second_colum, values=[[value]for value in round_list2])
worksheet.update(range_name=third_colum, values=[[value]for value in round_list3])

driver.quit()


