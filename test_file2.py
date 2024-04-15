from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import pyperclip
import gspread
from oauth2client.service_account import ServiceAccountCredentials 

#変数の設定
k=0 #初期条件
i = 0
j = 0
count = 0
def task():

    def initialize():
        options = Options()
        options.add_argument(r'--user-data-dir=C:\Users\User\AppData\Local\Google\Chrome\User Data')
        options.add_argument('--profile-directory=Profile 4')
        options.add_argument('--no-sandbox')
        driver_path = ChromeDriverManager().install()
        return webdriver.Chrome(options=options, service=Service(executable_path=driver_path))

    browser = initialize()

    url = "https://www.google.com/?hl=ja"

    browser.get(url)

    # URLを検索する
    url_input = browser.find_element(By.ID, "APjFqb")
    url_input.send_keys("google search console")
    url_input.send_keys(Keys.RETURN)

    # 検索候補の中からGoogle search consoleをクリック
    url_click = WebDriverWait(browser, 20).until(
        EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Google Search Console"))
    )
    url_click.send_keys(Keys.RETURN)

    # 今すぐ開始
    url_click = WebDriverWait(browser, 20).until(
        EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "今すぐ開始"))
    )
    url_click.send_keys(Keys.RETURN)

    # 検索結果
    result = WebDriverWait(browser, 20).until(
        EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "検索結果"))
    )
    result.send_keys(Keys.RETURN)


    # 新規
    newer = WebDriverWait(browser, 20).until(
        EC.element_to_be_clickable((By.ID, "hRpgHe"))
    )
    newer.send_keys(Keys.RETURN)

    # 検索キーワード
    key_word = WebDriverWait(browser, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'span[aria-label="検索キーワード…"]'))
    )
    key_word.send_keys(Keys.RETURN)

    # 入力
    input_name = WebDriverWait(browser, 30).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[6]/div[2]/div/div[1]/div/div/div[2]/span/div/div[2]/label/span[2]/input'))
    )
    input_name.send_keys("ses")
    input_name.send_keys(Keys.RETURN)
    time.sleep(3)

    # 過去28日に変更する(必要に応じて)
    input_name = WebDriverWait(browser, 20).until(
        EC.element_to_be_clickable((By.XPATH,'//*[@id="yDmH0d"]/c-wiz[4]/c-wiz/div/div[1]/div[2]/div/div/div[2]/div/div' ))
    )
    input_name.send_keys(Keys.RETURN)


    #変更する
    input_name = WebDriverWait(browser, 20).until(
        EC.element_to_be_clickable((By.XPATH,'//*[@id="tab1"]/div/div[1]/span/label[3]/div'))
    )
    input_name.click()


    #適用
    input_name = WebDriverWait(browser,20).until(
        EC.element_to_be_clickable((By.XPATH,'//*[@id="yDmH0d"]/div[6]/div/div[2]/div[3]/div[3]'))
    )
    input_name.send_keys(Keys.RETURN)
    time.sleep(5)


    #クリックして最下部までスクロール
    browser.find_element(By.TAG_NAME,"body").click()
    browser.find_element(By.TAG_NAME, "body").send_keys(Keys.PAGE_DOWN)
    for i in range(10):
        browser.find_element(By.TAG_NAME, "body").send_keys(Keys.PAGE_DOWN)
    time.sleep(3)

    # 最後のクリック前に待機
    selection = browser.find_element(By.CSS_SELECTOR, ".gb_Mc")
    browser.execute_script('arguments[0].click();', selection)
    time.sleep(3)

    #csvとして出力する
    download = browser.find_element(By.XPATH,"/html/body/div[1]/c-wiz[5]/c-wiz/div/div[1]/div[1]/div[1]/div[2]/div/div")
    browser.execute_script('arguments[0].click();', download)
    time.sleep(1)

    #csvをダウンロード
    download = browser.find_element(By.XPATH, '//*[@id="yDmH0d"]/c-wiz[5]/div/div/div/span[3]')
    download.send_keys(Keys.RETURN)
    time.sleep(20)

    ##########以下、スプシにペーストするコード##########
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

    #認証情報設定
    #ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
    credentials = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\User\Downloads\anal-410708-38581d77a165.json", scope)

    #OAuth2の資格情報を使用してGoogle APIにログインします。
    gc = gspread.authorize(credentials)

    #共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
    SPREADSHEET_KEY = '1eEQyRb_mn8_ZDwsGSF9unpDkz3862OeIweDuICiUOYE'

    #共有設定したスプレッドシートの特定のシートを開く
    worksheet = gc.open_by_key(SPREADSHEET_KEY).worksheet("【月間変化データ】")

    #指定した列すべての値を取得(c列)
    colum = worksheet.col_values(3)

    del colum[0]
    del colum[0]

    #dfに要素を追加
    output=pd.DataFrame(data=colum,columns=["スプシから引用"])


    #最新のファイルを開く
    import glob
    import os
    import zipfile
    #zipファイルを解凍するディレクトリ
    extract_dir = r"C:\Users\User\Downloads"
    # 最新のZIPファイルを取得
    latest_zip_file = max(glob.glob(os.path.join(extract_dir, '*.zip')), key=os.path.getctime)

    # ZIPファイルを解凍
    with zipfile.ZipFile(latest_zip_file, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

    # 解凍したディレクトリ内のCSVファイル一覧を取得
    csv_files = [file for file in os.listdir(extract_dir) if file.lower().endswith('.csv')]

    # 目的のCSVファイル名
    target_csv_filename = "クエリ.csv"
    # Pandasのread_csvを使用してCSVファイルをDataFrameに読み込む
    target_csv_path = os.path.join(extract_dir, target_csv_filename)
    df = pd.read_csv(target_csv_path, encoding="UTF-8")
    output2=df[['上位のクエリ','クリック数','表示回数','掲載順位']]

    #dataframe2を1に追加する
    output = pd.concat([output, output2], axis=1)


    #以下、不要なファイルを消す
    os.remove(latest_zip_file)
    link_list = [r"C:\Users\User\Downloads\クエリ.csv",r"C:\Users\User\Downloads\デバイス.csv",r"C:\Users\User\Downloads\フィルタ.csv",r"C:\Users\User\Downloads\ページ.csv",r"C:\Users\User\Downloads\検索での見え方.csv",r"C:\Users\User\Downloads\国.csv",r"C:\Users\User\Downloads\日付.csv"]
    for delate in link_list:
        os.remove(delate)

    browser.close() 
    browser.quit()

    def change_order(i,j):
    # print(output.iloc[:,[0]].isna().sum())#213
    #iがdfの0列目の値になるまで継続(784)
    #(len(output.iloc[:,[0]])-output.iloc[:,[0]].isna().sum())).any()
    # output['スプシから引用'].count()
        count = 0
        i = 0
        j = 0
        while(i<=output['スプシから引用'].count()):
            if output.iloc[i,0] == output.iloc[j%len(output.iloc[:,[1]]),1]:
                # print(output.iloc[i,0],output.iloc[j%len(output.iloc[:,[1]]),1])
                content_list=output.iloc[i,1:5]
                output.iloc[i,1:5]=output.iloc[j%len(output.iloc[:,[1]]),1:5]
                output.iloc[j%len(output.iloc[:,[1]]),1:5]=content_list #入れ替え完了
                i+=1
                j+=1
                print(i)
            else:
                j+=1
                count+=1
                if count>=len(output.iloc[:,[1]]):#1000を超えた場合の処理
                    output.iloc[i,1:5]="-"
                    i+=1
                    count=0
                else:
                    pass
        return(i,j)

    change_order(i,j)
    #以下、スプシに移す作業　

    str_output1 = output[["クリック数"]].to_string(index=False) #クリック数
    str_output2 = output[["表示回数"]].to_string(index=False) #表示回数
    str_output3 = output[["掲載順位"]].to_string(index=False) #掲載順位

    cleaned_text1 = "\n".join(line.strip() for line in str_output1.split("\n"))
    cleaned_text2 = "\n".join(line.strip() for line in str_output2.split("\n"))
    cleaned_text3 = "\n".join(line.strip() for line in str_output3.split("\n"))

    # print(str_output.count())#1000

    #クリップボードにcopy,paste
    pyperclip.copy(cleaned_text1)#追記
    paste1 = pyperclip.paste()
    pyperclip.copy(cleaned_text2)#追記
    paste2 = pyperclip.paste()
    pyperclip.copy(cleaned_text3)#追記
    paste3 = pyperclip.paste()

    #最も右(=最新の)行を取得
    colum_length = len(worksheet.row_values(3))#3行目の最も右
    print(colum_length)
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

    #2列,3列目の取得
    second = num+1
    third = second+1

    #3列の取得
    start_low = "2:"
    end_low = "2000"
    first_colum= str(Alphabet(num)+start_low+Alphabet(num)+end_low)#第1列
    second_colum= str(Alphabet(second)+start_low+Alphabet(second)+end_low)#第2列
    third_colum= str(Alphabet(third)+start_low+Alphabet(third)+end_low)#第3列

    #セルの更新
    worksheet.update(first_colum, values=[[value] for value in paste1.split('\n')])
    worksheet.update(second_colum, values=[[value] for value in paste2.split('\n')])
    worksheet.update(third_colum, values=[[value] for value in paste3.split('\n')])

task()