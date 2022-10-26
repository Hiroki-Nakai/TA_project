import tkinter
import tkinter.filedialog
import glob
import os
from PIL import Image, ImageTk
import tkinter.ttk as ttk
import pandas as pd
import numpy as np
import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import configparser
import shutil
import sys
import csv
import json
import pathlib
from tkinter import messagebox
import subprocess
#config読み込み
config_ini = configparser.ConfigParser()
config_ini.read('config.ini', encoding='utf-8')
USERNAME=config_ini["default"]["username"]
PASSWORD=config_ini["default"]["password"]
CLASS=config_ini["default"]["class"]
CLASS_CODE=config_ini["default"]["class_code"]
BASE_URL="https://ist.ksc.kwansei.ac.jp/~ishiura/cgi-bin/submit/main.cgi?"

#ダウンロードパスを絶対パスに変換する
#パスが存在しなければ現在のディレクトリを設定する
downloadPath= config_ini["default"]["downloadPath"]
try:
    downloadPath=downloadPath.resolve()
except:
    downloadPath=pathlib.Path.cwd()
if not os.path.exists(downloadPath):
    downloadPath=pathlib.Path.cwd()
downloadPath=str(downloadPath)

#option
options = Options()
UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.1 Safari/605.1.15'
options.add_argument('--user-agent=' + UA)
options.add_argument("--headless")
options.add_experimental_option('prefs', {'download.prompt_for_download': False,})

#ログイン認証の突破用にurlの書き換え
def make_login_url(url):
    return "https://"+USERNAME+":"+PASSWORD+"@"+url[8:]
def prepare_file():
    driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)
    driver.command_executor._commands["send_command"] = (
        'POST',
        '/session/$sessionId/chromium/send_command'
    )
    driver.execute(
        'send_command',
        params={
            'cmd': 'Page.setDownloadBehavior',
            'params': {'behavior': 'allow', 'downloadPath': downloadPath}
        }
    )
    info=["S="+CLASS_CODE,"act_class=1","c="+CLASS]
    url=make_login_url(BASE_URL+"&".join(info))
    #テーブルの取得
    driver.get(url)
    trs=driver.find_elements(By.TAG_NAME,"tr")
    #学籍番号と名前のpd.DataFrameを作成
    df=[]    
    for i in trs[4:-1]:
        row=i.text.split()
        df.append([int(row[0])," ".join(row[1:3])])
    df=pd.DataFrame(df,columns=["student number","name"])
    df.to_csv(f"{downloadPath}/student_class{CLASS}.csv",index=False,encoding='shift_jis')

    #課題名をすべて取得し、辞書型で保存
    assign_dic=dict()
    for i in trs[3].text.split()[3:-1]:
        info=["S="+CLASS_CODE,"act_report=1","c="+CLASS,"r="+i]
        url=make_login_url(BASE_URL+"&".join(info))
        driver.get(url)
        ths=driver.find_elements(By.TAG_NAME,"tr")[3].find_elements(By.TAG_NAME,"th")
        new=[]
        for j in ths[6:-1]:
            new.append(j.text.split()[0])
        assign_dic[i]=new
    tf = open(f"{downloadPath}/assign_dic.json", "w")
    json.dump(assign_dic,tf)
    tf.close()
    driver.close()

#ファイルの準備用（初回のみ）
if not (os.path.exists(f"{downloadPath}/student_class{CLASS}.csv") and os.path.exists(f"{downloadPath}/assign_dic.json")):
    prepare_file()


student_dic = {}
student_path = f"{downloadPath}/student_class{CLASS}.csv"
with open(student_path,encoding='shift_jis') as f:
    reader = csv.reader(f)
    for i, row in enumerate(reader):
        if i==0:
            continue
        student_dic[int(row[0])] = row[1]
        
student_list = []
for key, val in student_dic.items():
    student_list.append(" ".join([str(key), val]))
#課題一覧を取得
tf = open(f"{downloadPath}/assign_dic.json", "r")
assign_dic = json.load(tf)
assign_list=list(assign_dic.keys())


class Application(tkinter.Tk):
    def __init__(self):
        super().__init__()
        
        self.canvas_width = 900
        self.canvas_height = 500
        
        self.geometry("1230x800")
        self.title("採点システム")
        
        # 画像表示のキャンバス作成と配置
        self.canvas = tkinter.Canvas(
            self,
            width=self.canvas_width,
            height=self.canvas_height,
            bg="gray",
            scrollregion = (0,0,0,0)
        )
        self.canvas.grid(row=0, column=0)
        
        # スクロールバーを作成
        xbar = tkinter.Scrollbar(orient=tkinter.HORIZONTAL)  # バーの方向
        ybar = tkinter.Scrollbar(orient=tkinter.VERTICAL)  # バーの方向
        # キャンバスにスクロールバーを配置
        xbar.grid(
            row=1, column=0,  # キャンバスの下の位置を指定
            sticky=tkinter.W + tkinter.E  # 左右いっぱいに引き伸ばす
        )
        ybar.grid(
            row=0, column=1,  # キャンバスの右の位置を指定
            sticky=tkinter.N + tkinter.S  # 上下いっぱいに引き伸ばす
        )
        # スクロールバーのスライダーが動かされた時に実行する処理を設定
        xbar.config(command=self.canvas.xview)
        ybar.config(command=self.canvas.yview)
        # キャンバススクロール時に実行する処理を設定
        self.canvas.config(xscrollcommand=xbar.set)
        self.canvas.config(yscrollcommand=ybar.set)
        
        # pyファイル実行結果表示用
        # 画像表示のキャンバス作成と配置
        self.canvas_code = tkinter.Canvas(
            self,
            width=self.canvas_width,
            height=150,
            scrollregion = (0,0,3000,600)
        )
        self.canvas_code.grid(row=2, column=0)
        
        # スクロールバーを作成
        xbar2 = tkinter.Scrollbar(orient=tkinter.HORIZONTAL)  # バーの方向
        ybar2 = tkinter.Scrollbar(orient=tkinter.VERTICAL)  # バーの方向
        # キャンバスにスクロールバーを配置
        xbar2.grid(
            row=3, column=0,  # キャンバスの下の位置を指定
            sticky=tkinter.W + tkinter.E  # 左右いっぱいに引き伸ばす
        )
        ybar2.grid(
            row=2, column=1,  # キャンバスの右の位置を指定
            sticky=tkinter.N + tkinter.S  # 上下いっぱいに引き伸ばす
        )
        # スクロールバーのスライダーが動かされた時に実行する処理を設定
        xbar2.config(command=self.canvas_code.xview)
        ybar2.config(command=self.canvas_code.yview)
        # キャンバススクロール時に実行する処理を設定
        self.canvas_code.config(xscrollcommand=xbar2.set)
        self.canvas_code.config(yscrollcommand=ybar2.set)
        # ボタンを配置するフレームの作成と配置
        self.button_frame = tkinter.Frame(width = 1200-self.canvas_width,
                                          height = self.canvas_height)
        self.button_frame.grid(row=0, column=2)
        
        self.font = ("",12)
         # 課題コード選択ボックス
        self.comb_assign_code = ttk.Combobox(self.button_frame,
                                     values=assign_list, font=("",4))
        
        # 課題選択ボックス
        self.comb_assign = ttk.Combobox(self.button_frame,
                                     values=student_list, font=("",))
        # ファイル読み込みボタンの作成と配置
        self.load_button = tkinter.Button(
            self.button_frame,
            text = "課題選択",
            command = self.push_download_button,
            bg = "light cyan",
            font=self.font
        )
        
        # 学生選択ボックス
        self.comb_student = ttk.Combobox(self.button_frame,
                                     values=student_list, font=("",11))
        # 学生選択ボタン
        self.choice_button = tkinter.Button(
            self.button_frame,
            text = "選択",
            command = self.push_choice_button,
            bg = "light cyan",
            font=self.font
        )
        
        # 前、次へボタンの作成
        self.before_button = tkinter.Button(
            self.button_frame,
            text = "前へ",
            command = self.push_before_button,
            bg = "light cyan",
            font=self.font, width = 10
        )
        self.next_button = tkinter.Button(
            self.button_frame,
            text = "次へ",
            command = self.push_next_button,
            bg = "light cyan",
            font=self.font, width = 10
        )
        
        # 拡大、縮小ボタンの作成と配置
        self.big_button = tkinter.Button(
            self.button_frame,
            text = "拡大",
            command = self.push_big_button,
            bg = "light cyan",
            font = self.font, width = 10
        )
        self.small_button = tkinter.Button(
            self.button_frame,
            text = "縮小",
            command = self.push_small_button,
            bg = "light cyan",
            font = self.font, width = 10
        )
        
        # OK、NGボタンの作成と配置
        self.ok_button = tkinter.Button(
            self.button_frame,
            text = "OK",
            command = self.push_ok_button,
            bg = "LightSkyBlue",
            font = self.font, width = 10
        )
        self.ng_button = tkinter.Button(
            self.button_frame,
            text = "NG",
            command = self.push_ng_button,
            bg = "IndianRed1",
            font = self.font, width = 10
        )
        
        # ファイル出力ボタンの作成と配置
        self.output_button = tkinter.Button(
            self.button_frame,
            text = "ファイル出力",
            command = self.push_output_button,
            bg = "seaGreen1",
            font=self.font
        )
        
        self.corrent_num = 0
        # 画像オブジェクトの設定（初期はNone）
        self.image = None
        self.image_orignal = None

        # キャンバスに描画中の画像（初期はNone）
        self.canvas_obj= None
        
        self.files = None
        self.file_len = 0
        self.ASSIGN=""
        self.ASSIGN_NUM=""
        self.CODE=None
        self.message = "課題番号を入力してください"
        self.lbl_message = tkinter.Label(self.button_frame,
                                         text=self.message, background="white",
                                         font=self.font)
        self.student = ""
        self.lbl_student = tkinter.Label(self.button_frame,
                                         text=self.student, background="white",
                                         font=self.font)
        self.msg_box()
        
        self.txt_set = set()
        self.txt_list = list(self.txt_set)
        self.comb_txt = ttk.Combobox(self.button_frame,
                                     values=self.txt_list, font=self.font)
        #コード実行結果を格納
        self.result_code=None

         # 課題コード選択ボックス
        self.assign_code = ttk.Combobox(self.button_frame,
                                     values=assign_list, font=("",12), width=4)
        self.assign_code.current(0)
        self.assign_code.bind('<<ComboboxSelected>>', lambda event: self.assign.config(values=assign_dic[self.assign_code.get()]))
        # 課題選択ボックス
        self.assign = ttk.Combobox(self.button_frame,
                                     values=assign_dic[assign_list[0]], font=("",12), width=10)
        self.df = None
        self.output_file=None
        self.canvas_code_obj=None
        #self.df = self.mk_df()
        self.state_dict={"1":"OK","0":"NG",np.nan:"未採点","未提出":"未提出"}

        # gridでウェイジェットの配置
        pady = 12
        self.lbl_message.grid(row=0, column=0, columnspan=6, padx=10, pady=pady, 
                              sticky=tkinter.W+tkinter.E)
        self.assign_code.grid(row=1, column=0, columnspan=1, padx=5, pady=pady)
        ttk.Label(self.button_frame, text='の',font=("",12)).grid(row=1, column=1, columnspan=1, padx=0, pady=pady)
        self.assign.grid(row=1, column=2, columnspan=2, padx=10, pady=pady)
        
        self.load_button.grid(row=1, column=4, columnspan=2, padx=10, pady=pady)
        self.lbl_student.grid(row=2, column=0, columnspan=6, padx=10, pady=pady,
                              sticky=tkinter.W+tkinter.E)
        self.comb_student.grid(row=3, column=0, columnspan=5, padx=10, pady=pady)
        self.choice_button.grid(row=3, column=5, columnspan=1, padx=10, pady=pady)
        self.before_button.grid(row=4, column=0, columnspan=3, padx=15, pady=pady)
        self.next_button.grid(row=4, column=3, columnspan=3, padx=10, pady=pady)
        self.small_button.grid(row=5, column=0, columnspan=3, padx=15, pady=pady)
        self.big_button.grid(row=5, column=3, columnspan=3, padx=10, pady=pady)
        self.comb_txt.grid(row=6, column=0, columnspan=6, padx=10, pady=pady,
                      sticky=tkinter.W+tkinter.E)
        self.ok_button.grid(row=7, column=0, columnspan=3, padx=15, pady=pady)
        self.ng_button.grid(row=7, column=3, columnspan=3, padx=10, pady=pady)
        self.output_button.grid(row=8, column=0, columnspan=6, padx=10, pady=pady,
                                sticky=tkinter.W+tkinter.E)
        
    def push_download_button(self):
        num=self.assign_code.get()
        assign = self.assign.get()
        if len(assign)==0:
            self.message = "課題名を入力してください"
        elif len(num)==0:
            self.message = "課題番号を入力してください"
        else:
            try:
                self.ASSIGN=assign
                self.ASSIGN_NUM=num
                self.output_file=f"{downloadPath}/eval_{self.ASSIGN}.xlsx"
                if os.path.exists(downloadPath+"/"+self.ASSIGN+'-'+str(CLASS)):
                    ret = messagebox.askyesno("確認", "すでに課題が存在します。\n再ダウンロードしますか？")
                    if ret == False:
                        self.push_load_button()
                        return
                info=["S="+CLASS_CODE,"act_download_multi=1","c="+CLASS,"r="+self.ASSIGN_NUM,"i="+self.ASSIGN]
                url=make_login_url(BASE_URL+"&".join(info))
                self.message = "ダウンロードを行います"
                
                #webdriver起動
                self.msg_box()
                driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)
                driver.command_executor._commands["send_command"] = (
                    'POST',
                    '/session/$sessionId/chromium/send_command'
                )
                driver.execute(
                    'send_command',
                    params={
                        'cmd': 'Page.setDownloadBehavior',
                        'params': {'behavior': 'allow', 'downloadPath': downloadPath}
                    }
                )
                
                #テーブルの取得
                driver.get(url)
                time.sleep(1)
                driver.close()
                self.message = "ダウンロードが終了しました"
            except:
                self.message = "課題名が不正です"
            
            self.msg_box()
            
            self.push_load_button()
            
    def push_load_button(self,f=True):
        if f:
            shutil.unpack_archive(downloadPath+"/"+self.ASSIGN+'-'+str(CLASS)+'.zip', downloadPath)
        folder_path = downloadPath+"/"+self.ASSIGN+'-'+str(CLASS)
        if self.ASSIGN.split(".")[-1]=="py":
            self.files = glob.glob(f"{folder_path}/*.py")
            # フォルダ内にpyファイルが存在するか確認
            self.file_len = len(self.files)
            if self.file_len==0:
                basename = ""
            else:
                basename = os.path.splitext(os.path.basename(self.files[0]))[0]
            # pyファイルがプロ1のものか判定（ファイル名が8桁の数字であることが条件）
            if self.file_len!=0 and len(basename)==8 and basename.isdigit():
                self.corrent_num = 0
                self.CODE=self.ASSIGN
                self.message = ""
                if os.path.exists(self.output_file):
                    ret = messagebox.askyesno("確認", "すでに出力ファイルが存在します。\n前回の続きから再開しますか？")
                    if ret == True:
                        self.df=self.read_df()
                    else:
                        self.df = self.mk_df()
                else:
                    self.df = self.mk_df()
                self.result_code=self.execute_py()
                if self.result_code is not None:
                    if self.canvas_code_obj is not None:
                        self.canvas_code.delete(self.canvas_code_obj)
                    self.canvas_code_obj=self.canvas_code.create_text(10, 10, anchor="nw", text="コードの実行が完了しました！",fill="black",font=self.font)
            else:
                self.files = None
                self.result_code = None
                self.df = None
                self.message = "フォルダが適切ではありません"

        elif self.ASSIGN.split(".")[-1]=="png":
            self.files = glob.glob(f"{folder_path}/*.png")
            # フォルダ内にpngファイルが存在するか確認
            self.file_len = len(self.files)
            if self.file_len==0:
                basename = ""
            else:
                basename = os.path.splitext(os.path.basename(self.files[0]))[0]
            # 画像ファイルがプロ1のものか判定（ファイル名が8桁の数字であることが条件）
            if self.file_len!=0 and len(basename)==8 and basename.isdigit():
                self.corrent_num = 0
                self.image_original = Image.open(self.files[self.corrent_num])
                self.image = self.image_original.copy()
                self.message = ""
                self.view_image(self.corrent_num)
                if self.result_code is not None:
                    if self.canvas_code_obj is not None:
                        self.canvas_code.delete(self.canvas_code_obj)
                    self.canvas_code_obj=self.canvas_code.create_text(10, 10, anchor="nw", text=self.result_code[int(basename)],fill="black",font=self.font)
                if os.path.exists(self.output_file):
                    ret = messagebox.askyesno("確認", "すでに出力ファイルが存在します。\n前回の続きから再開しますか？")
                    if ret == True:
                        self.df=self.read_df()
                    else:
                        self.df = self.mk_df()
                else:
                    self.df = self.mk_df()
            else:
                self.files = None
                self.image = None
                self.df = None
                # キャンバスに描画中の画像を削除
                if self.canvas_obj is not None:
                    self.canvas.delete(self.canvas_obj)
                self.message = "フォルダが適切ではありません"
                self.student = ""
        self.msg_box()
        
    def push_choice_button(self):
        self.message = "ファイルが存在しません"
        if self.files is not None:
            value = self.comb_student.get()
            value = value.split(" ")[0]
            #print(value)
            #print(type(value))
            for i, file in enumerate(self.files):
                if value in file:
                    #print(i, file)
                    self.image_original = Image.open(file)
                    self.image = self.image_original.copy()
                    self.view_image(i)
                    self.message = ""
                    if self.result_code is not None:
                        if self.canvas_code_obj is not None:
                            self.canvas_code.delete(self.canvas_code_obj)
                        self.canvas_code_obj=self.canvas_code.create_text(10, 10, anchor="nw", text=self.result_code[int(value)],fill="black",font=self.font)
                    if int(value) in self.df.index:
                        self.message=self.state_dict[self.df["判定"][int(value)]]
                    self.corrent_num = i
                    break
        self.msg_box()
        
    def push_before_button(self):
        if self.image is not None:
            self.corrent_num = (self.corrent_num-1) % self.file_len
            filename = self.files[self.corrent_num]
            basename = os.path.splitext(os.path.basename(filename))[0]
            self.image_original = Image.open(self.files[self.corrent_num])
            self.image = self.image_original.copy()
            self.view_image(self.corrent_num)
            self.message = ""
            if self.result_code is not None:
                if self.canvas_code_obj is not None:
                    self.canvas_code.delete(self.canvas_code_obj)
                self.canvas_code_obj=self.canvas_code.create_text(10, 10, anchor="nw", text=self.result_code[int(basename)],fill="black",font=self.font)
            if int(basename) in self.df.index:
                    self.message=self.state_dict[self.df["判定"][int(basename)]]
            ind = self.return_list_index(basename)
            if ind is not None:
                self.comb_student.delete(0, tkinter.END)
                self.comb_student.insert(tkinter.END, student_list[ind])
        else:
            self.message = "ファイルが存在しません"
        self.msg_box()
        
    
    def push_next_button(self):
        if self.image is not None:
            self.corrent_num = (self.corrent_num+1) % self.file_len
            filename = self.files[self.corrent_num]
            basename = os.path.splitext(os.path.basename(filename))[0]
            self.image_original = Image.open(filename)
            self.image = self.image_original.copy()
            self.view_image(self.corrent_num)
            self.message = ""
            if int(basename) in self.df.index:
                    self.message=self.state_dict[self.df["判定"][int(basename)]]
                    if self.result_code is not None:
                        if self.canvas_code_obj is not None:
                            self.canvas_code.delete(self.canvas_code_obj)
                        self.canvas_code_obj=self.canvas_code.create_text(10, 10, anchor="nw", text=self.result_code[int(basename)],fill="black",font=self.font)
            ind = self.return_list_index(basename)
            if ind is not None:
                self.comb_student.delete(0, tkinter.END)
                self.comb_student.insert(tkinter.END, student_list[ind])
        else:
            self.message = "ファイルが存在しません"
        self.msg_box()
        
    def push_big_button(self):
        if self.image is not None:
            width = self.image.width
            height = self.image.height
            new_width = int(width*11/10)
            new_height = int(height * new_width/width)
            self.image = self.image_original.resize((new_width, new_height))
            self.image_tk = ImageTk.PhotoImage(self.image)
            # キャンバスに描画中の画像を削除
            if self.canvas_obj is not None:
                self.canvas.delete(self.canvas_obj)
            # 画像を描画
            self.canvas_obj = self.canvas.create_image(
                0, 0, image=self.image_tk,anchor=tkinter.NW
            )
            # スクロールバーの範囲を設定
            self.canvas.config(scrollregion = (0,0,self.image.width,self.image.height))
    
    def push_small_button(self):
        if self.image is not None:
            width = self.image.width
            height = self.image.height
            new_width = int(width*10/11)
            new_height = int(height * new_width/width)
            self.image = self.image_original.resize((new_width, new_height))
            self.image_tk = ImageTk.PhotoImage(self.image)
            # キャンバスに描画中の画像を削除
            if self.canvas_obj is not None:
                self.canvas.delete(self.canvas_obj)
            # 画像を描画
            self.canvas_obj = self.canvas.create_image(
                0, 0, image=self.image_tk,anchor=tkinter.NW
            )
            
            # スクロールバーの範囲を設定
            self.canvas.config(scrollregion = (0,0,self.image.width,self.image.height))
        
    def push_ok_button(self):
        if self.df is not None:
            filename = self.files[self.corrent_num]
            basename = int(os.path.splitext(os.path.basename(filename))[0])
            comment = self.comb_txt.get()
            # コメントのコンボボックスにコメントを追加
            self.txt_set.add(comment)
            self.txt_list = list(self.txt_set)
            self.comb_txt.config(values=self.txt_list)
            self.df.loc[basename, "判定"] = "1"
            self.df.loc[basename, "コメント"] = comment
            self.df.loc[basename, "system"] = "!ok"
            self.comb_txt.delete(0, tkinter.END)
            self.message = "OK"
            # コンボボックスの表示を変更
            ind = self.return_list_index(str(basename))
            if ind is not None:
                student_list[ind] = student_list[ind].replace(" *", "")
                self.comb_student.config(values=student_list)
                self.comb_student.delete(0, tkinter.END)
                self.comb_student.insert(tkinter.END, student_list[ind])
                self.push_next_button()
            self.msg_box()
          
    def push_ng_button(self):
        if self.df is not None:
            filename = self.files[self.corrent_num]
            basename = int(os.path.splitext(os.path.basename(filename))[0])
            comment = self.comb_txt.get()
            # コメントのコンボボックスにコメントを追加
            self.txt_set.add(comment)
            self.txt_list = list(self.txt_set)
            self.comb_txt.config(values=self.txt_list)
            if len(comment)==0:
                self.message = "コメントを入力してください"
            else:
                self.df.loc[basename, "判定"] = "0"
                self.df.loc[basename, "コメント"] = comment
                self.df.loc[basename, "system"] = "!NG/"+comment
                self.comb_txt.delete(0, tkinter.END)
                self.message = "NG"
                # コンボボックスの表示を変更
                ind = self.return_list_index(str(basename))
                if ind is not None:
                    student_list[ind] = student_list[ind].replace(" *", "")
                    self.comb_student.config(values=student_list)
                    self.comb_student.delete(0, tkinter.END)
                    self.comb_student.insert(tkinter.END, student_list[ind])
                    self.push_next_button()
            self.msg_box()
        
    def push_output_button(self):
        if self.df is not None:
            null_num = self.df.isnull().sum()["判定"]
            if null_num == 0:
                ret = messagebox.askyesno("確認", "excelに出力しますか？")
                if ret == True:
                    self.output()
                else:
                    self.message = "ファイルを出力しませんでした"
            else:
                ret = messagebox.askyesno("確認", "未採点の課題があります。\nexcelに出力しますか？")
                if ret == True:
                    self.output()
                else:
                    self.message = "ファイルを出力しませんでした"
            if self.CODE is not None:
                ret = messagebox.askyesno("確認", self.CODE+"にも同じ結果を出力しますか？")
                if ret == True:
                    self.output()
                else:
                    self.message = self.CODE+"にはファイルを出力しませんでした"
            self.msg_box()
    
    def output(self):
        
        self.df.to_excel(self.output_file)
        self.message = "excelファイルを出力しました"
        self.msg_box()
        ret = messagebox.askyesno("確認", "結果をレポートシステムに出力しますか？")
        if ret == True:
            info=["S="+CLASS_CODE,"act_report=1","c="+CLASS,"r="+self.ASSIGN_NUM,"eval_i="+self.ASSIGN]
            url=make_login_url(BASE_URL+"&".join(info))
            driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)
            driver.get(url)
            flg = False
            for i in self.df.index.values:
                if type(self.df["system"][i]) is not str: 
                    continue
                if len(self.df["system"][i])==0:
                    continue
                flg = True
                a=driver.find_element(By.NAME,"eval:"+str(i)+":"+self.ASSIGN)
                a.clear()
                a.send_keys(self.df["system"][i])   
            if flg:
                a.send_keys(Keys.ENTER)
            driver.close()
            
    def view_image(self, view_num):
        # 画像の描画位置を調節
        x = 0#int(self.canvas_width / 2)
        y = 0#int(self.canvas_height / 2)
        # 画像サイズを変更
        width = self.image.width
        height = self.image.height
        # スクロールバーの範囲を設定
        self.canvas.config(scrollregion = (0,0,width,height))
        
        self.image_tk = ImageTk.PhotoImage(self.image)

        # キャンバスに描画中の画像を削除
        if self.canvas_obj is not None:
            self.canvas.delete(self.canvas_obj)
        # 画像を描画
        self.canvas_obj = self.canvas.create_image(
            x, y,
            image=self.image_tk,anchor=tkinter.NW
        )
        #　ファイル名から学籍番号と氏名を取得し、表示する
        basename = os.path.splitext(os.path.basename(self.files[view_num]))[0]
        self.student = basename + " " + student_dic[int(basename)]
        self.msg_box()
    # code実行
    def execute_py(self):
        result={i:"提出されていません" for i in student_dic.keys()}
        for j in self.files:
            student_num=int(j.split("\\")[-1].split(".")[0])
            #print(student_num)
            try:
                p=subprocess.run("python "+j,input="",stdout=subprocess.PIPE,shell=False,encoding='cp932',timeout=3)
                if p.returncode:
                    result[student_num]="エラーです。"
                else:
                    result[student_num]=p.stdout
            except:
                result[student_num]="終わりません"
        return result
    def msg_box(self):
        self.lbl_message.config(text=self.message)
        self.lbl_student.config(text=self.student)
        
    def read_df(self):
        df=pd.read_excel(self.output_file, index_col=0)
        for file in self.files:
            basename = int(os.path.splitext(os.path.basename(file))[0])
            if df["判定"][basename]!="1":
                # コンボボックスの要素のうち未提出かok以外のものに*を記入
                ind = self.return_list_index(str(basename))
                student_list[ind] = student_list[ind].replace(" *", "") + " *"
        self.comb_student.config(values=student_list)
        self.txt_set=set(df["コメント"].dropna())
        self.txt_list = list(self.txt_set)
        self.comb_txt.config(values=self.txt_list)
        return df

    def mk_df(self):
        nan = np.full((len(student_list), 4), np.nan)
        df = pd.DataFrame(data=nan, index=student_dic.keys(), columns=["名前","判定", "コメント","system"])
        df["名前"] = student_dic.values()
        df["判定"] = "未提出"
        df["コメント"] = ""
        df["system"] = ""
        for file in self.files:
            basename = int(os.path.splitext(os.path.basename(file))[0])
            df.at[basename, "判定"] = np.nan
            # コンボボックスの要素のうち未提出以外のものに*を記入
            ind = self.return_list_index(str(basename))
            student_list[ind] = student_list[ind].replace(" *", "") + " *"
        self.comb_student.config(values=student_list)
        
        return df
    
    # 入力に一致するstudent_listの要素のインデックスを返す関数
    def return_list_index(self, txt):
        for i, val in enumerate(student_list):
            if txt in val:
                return i
        return None
    
app = Application()
app.mainloop()