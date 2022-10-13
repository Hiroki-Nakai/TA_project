import tkinter
import tkinter.filedialog
import glob
import os
import sys
import csv
from PIL import Image, ImageTk
import tkinter.ttk as ttk
import pandas as pd
import numpy as np

# 外部ファイルからの読み込みのための関数
def resource_path(filename):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)

student_dic = {}
with open(resource_path('student.csv')) as f:
    reader = csv.reader(f)
    for i, row in enumerate(reader):
        if i==0:
            continue
        student_dic[int(row[0])] = row[1]
        
student_list = []
for key, val in student_dic.items():
    student_list.append(" ".join([str(key), val]))

class Application(tkinter.Tk):
    def __init__(self):
        super().__init__()
        
        self.canvas_width = 900
        self.canvas_height = 500
        
        self.geometry("1200x530")
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
        
        # ボタンを配置するフレームの作成と配置
        self.button_frame = tkinter.Frame(width = 1200-self.canvas_width,
                                          height = self.canvas_height)
        self.button_frame.grid(row=0, column=2)
        
        self.font = ("",12)
        # ファイル読み込みボタンの作成と配置
        self.load_button = tkinter.Button(
            self.button_frame,
            text = "フォルダ選択",
            command = self.push_load_button,
            bg = "light cyan",
            font=self.font
        )
        
        # 選択ボックス
        self.combobox = ttk.Combobox(self.button_frame,
                                     values=student_list, font=("",11))
        # 選択ボタン
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
        
        self.message = "フォルダを選択してください"
        self.lbl_message = tkinter.Label(self.button_frame,
                                         text=self.message, background="white",
                                         font=self.font)
        self.student = ""
        self.lbl_student = tkinter.Label(self.button_frame,
                                         text=self.student, background="white",
                                         font=self.font)
        self.msg_box()
        
        self.txt = tkinter.Entry(self.button_frame, font=self.font)
        
        self.df = None
        #self.df = self.mk_df()
        
        # gridでウェイジェットの配置
        pady = 12
        self.load_button.grid(row=0, column=0, columnspan=6, padx=10, pady=pady,
                              sticky=tkinter.W+tkinter.E)
        self.lbl_message.grid(row=1, column=0, columnspan=6, padx=10, pady=pady, 
                              sticky=tkinter.W+tkinter.E)
        self.lbl_student.grid(row=2, column=0, columnspan=6, padx=10, pady=pady,
                              sticky=tkinter.W+tkinter.E)
        self.combobox.grid(row=3, column=0, columnspan=5, padx=10, pady=pady)
        self.choice_button.grid(row=3, column=5, padx=10, pady=pady)
        self.before_button.grid(row=4, column=0, columnspan=3, padx=15, pady=pady)
        self.next_button.grid(row=4, column=3, columnspan=3, padx=10, pady=pady)
        self.small_button.grid(row=5, column=0, columnspan=3, padx=15, pady=pady)
        self.big_button.grid(row=5, column=3, columnspan=3, padx=10, pady=pady)
        self.txt.grid(row=6, column=0, columnspan=6, padx=10, pady=pady,
                      sticky=tkinter.W+tkinter.E)
        self.ok_button.grid(row=7, column=0, columnspan=3, padx=15, pady=pady)
        self.ng_button.grid(row=7, column=3, columnspan=3, padx=10, pady=pady)
        self.output_button.grid(row=8, column=0, columnspan=6, padx=10, pady=pady,
                                sticky=tkinter.W+tkinter.E)
        
    def push_load_button(self):
        folder_path = tkinter.filedialog.askdirectory()
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
            value = self.combobox.get()
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
            ind = self.return_list_index(basename)
            if ind is not None:
                self.combobox.delete(0, tkinter.END)
                self.combobox.insert(tkinter.END, student_list[ind])
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
            ind = self.return_list_index(basename)
            if ind is not None:
                self.combobox.delete(0, tkinter.END)
                self.combobox.insert(tkinter.END, student_list[ind])
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
            comment = self.txt.get()
            self.df.loc[basename, "判定"] = "OK"
            self.df.loc[basename, "コメント"] = comment
            self.txt.delete(0, tkinter.END)
            # コンボボックスの表示を変更
            ind = self.return_list_index(str(basename))
            if ind is not None:
                student_list[ind] = student_list[ind].replace(" *", "")
                self.combobox.config(values=student_list)
                self.combobox.delete(0, tkinter.END)
                self.combobox.insert(tkinter.END, student_list[ind])
          
    def push_ng_button(self):
        if self.df is not None:
            filename = self.files[self.corrent_num]
            basename = int(os.path.splitext(os.path.basename(filename))[0])
            comment = self.txt.get()
            if len(comment)==0:
                self.message = "コメントを入力してください"
            else:
                self.df.loc[basename, "判定"] = "NG"
                self.df.loc[basename, "コメント"] = comment
                self.txt.delete(0, tkinter.END)
                # コンボボックスの表示を変更
                ind = self.return_list_index(str(basename))
                if ind is not None:
                    student_list[ind] = student_list[ind].replace(" *", "")
                    self.combobox.config(values=student_list)
                    self.combobox.delete(0, tkinter.END)
                    self.combobox.insert(tkinter.END, student_list[ind])
            self.msg_box()
        
    def push_output_button(self):
        if self.df is not None:
            null_num = self.df.isnull().sum()["判定"]
            if null_num == 0:
                #self.df.to_excel("output.xls", encoding="shift jis")
                try:
                    self.df.to_excel("output.xlsx")
                    self.message = "ファイル(output.xlsx)を出力しました"
                except:
                    self.message = "ファイルを出力できませんでした"
            else:
                self.message = "未採点の課題があります"
            self.msg_box()
            
    def view_image(self, view_num):
        # 画像の描画位置を調節
        x = 0#int(self.canvas_width / 2)
        y = 0#int(self.canvas_height / 2)
        # 画像サイズを変更
        width = self.image.width
        height = self.image.height
        # スクロールバーの範囲を設定
        self.canvas.config(scrollregion = (0,0,width,height))
        
        """
        if width>=self.canvas_width or height>=self.canvas_height:
            if width/height >= self.canvas_width/self.canvas_height:
                new_width = int(self.canvas_width)
                new_height = int(height * new_width/width)
                self.image = self.image.resize((new_width, new_height))
            else:
                new_height = int(self.canvas_height)
                new_width = int(width * new_height/height)
                self.image = self.image.resize((new_width, new_height))
        """
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
        
    def msg_box(self):
        self.lbl_message.config(text=self.message)
        self.lbl_student.config(text=self.student)
    
    def mk_df(self):
        nan = np.full((len(student_list), 3), np.nan)
        df = pd.DataFrame(data=nan, index=student_dic.keys(), columns=["名前","判定", "コメント"])
        df["名前"] = student_dic.values()
        df["判定"] = "未提出"
        df["コメント"] = ""
        for file in self.files:
            basename = int(os.path.splitext(os.path.basename(file))[0])
            df.at[basename, "判定"] = np.nan
            # コンボボックスの要素のうち未提出以外のものに*を記入
            ind = self.return_list_index(str(basename))
            student_list[ind] = student_list[ind].replace(" *", "") + " *"
        self.combobox.config(values=student_list)
        
        return df
    
    # 入力に一致するstudent_listの要素のインデックスを返す関数
    def return_list_index(self, txt):
        for i, val in enumerate(student_list):
            if txt in val:
                return i
        return None
    
app = Application()
app.mainloop()