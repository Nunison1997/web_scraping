import urllib.request
from bs4 import BeautifulSoup
import json 
import pyperclip
from urllib.error import URLError
import tkinter as tk
import time

def trial():
    global output  # output をグローバル変数として宣言
    url = entry.get()

    # urlが適切か判断
    try: 
        html = urllib.request.urlopen(url=url)
        # htmlをBeautifulSoupで扱う
        soup = BeautifulSoup(html, "html.parser")
        title = soup.title.string
        contents_list = url.split("/")  # /で文字列を分割
    except URLError as e:
        print("URLは適切ではありません")

    # カテゴリを取得
    category = soup.find('a',class_="inline-block ml-3 category-btn")
    category_list = ["エンジニア","IT業界","プログラマー・開発","インフラエンジニア","システムエンジニア（SE）","インフラ資格","プログラマー資格","その他IT資格","プログラミング"]

    # 長さが5ならそのままタイトルを返す
    if len(contents_list) == 5:
        join1 = "/".join(contents_list[0:3])
        url1 = join1+"/"
        list_element =         {
                "@type": "ListItem",
                "position": 1,
                "item": {
                    "@id": url1,
                    "name": "HOME"
                }
            }

    # カテゴリの分類
    if category == category_list[0]:
        category_content = "engineer"
        k = 0
    elif category == category_list[1]:
        category_content = "it-industry"
        k = 1
    elif category == category_list[2]:
        category_content = "development"
        k = 2
    elif category == category_list[3]:
        category_content = "infrastructure-engineer"
        k = 3
    elif category == category_list[4]:
        category_content = "system-engineer"
        k = 4
    elif category == category_list[5]:
        category_content = "infura-shikaku"
        k = 5
    elif category == category_list[6]:
        category_content = "programmer-shikaku"
        k = 6
    elif category == category_list[7]:
        category_content = "it-certification"
        k =7 
    else :
        category_content = "programming"
        k = 8


    # 長さが6なら3つのリストに分ける
    if len(contents_list) ==6:
        join1 = "/".join(contents_list[0:4])
        join2 = "/".join(contents_list)
        url1 = join1+"/"
        url2 = join1+"/"+"category"+"/"+category_content+"/"
        url3 = join2+"/"

        # elementの編集
        list_element =          {
                "@type": "ListItem",
                "position": 1,
                "item": {
                    "@id": url1,
                    "name": "TOP"
                }
            },{
                "@type": "ListItem",
                "position": 2,
                "item": {
                    "@id": url2,
                    "name": category_list[k]
                }
            },{
                "@type": "ListItem",
                "position": 3,
                "item": {
                    "@id": url3,
                    "name": title
                }
            }
    # json変換前
    data = {
        "@context": "http://schema.org",
        "@type": "BreadcrumbList",
        "itemListElement": [
            list_element
        ]
    }
    # json形式で変換
    output = json.dumps(data, indent=2, ensure_ascii=False)
    label = tk.Label(text="構造化データを取得できました")
    # time.sleep(5)
    label.place(x=450, y=245)
    win.after(3000, lambda: label.place_forget())
    entry.delete(0, tk.END)   # テキストボックス内の文字を削除

def click():
    global output
    pyperclip.copy(output)  # クリップボードにコピー
    label = tk.Label(text="クリップボードにコピーされました")
    label.place(x=450, y=330)
    win.after(5000, lambda: label.place_forget())

def reset():
    global entry
    entry.delete(0,tk.END)

# ウィンドウの作成
win = tk.Tk()
win.title("構造化データ作成アプリ")
win.geometry("1000x500")

# テキストボックス
entry = tk.Entry(textvariable="URL:")
entry.place(x=350, y=200, width=270,height=30)
# ボタン
button = tk.Button(text="OK", command=trial)
button.place(x=620, y=200, height=30)

#リセットボタン
rebutton = tk.Button(text="リセット", command=reset)
rebutton.place(x=650, y=200, height=30,width=40)

label = tk.Label(win)
label2 = tk.Label(text="構造化データ取得システム",font=("MSゴシック", "20", "bold"))
label2.place(x=350, y=100)

button2 = tk.Button(text="クリップボードにコピーする", command=click)
button2.place(x=450, y=280,height=30)
win.mainloop()
