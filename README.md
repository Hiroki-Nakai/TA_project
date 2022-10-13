# TA_project
## TA(teaching assistant)業務の効率化システム
以下の機能を持つ
- 画像ビューワー
- 採点機能
- 採点結果出力

### exe化
1. 仮想環境を使う

```activate [環境名]```

2. pyinstallerのインストール

```pip install pyinstaller```

3. exe化

```pyinstaller [スクリプトファイル名].py --onefile --noconsole```

### 外部ファイル読み込み

クラスに属する学生情報をstudent.csvから読み込む場合

1. スクリプトファイルの編集
```python:XXX.py
def resource_path(filename):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)
    
with open(resource_path('student.csv')) as f:
```

1. specファイルの編集

```
datas = []
```
を以下に書き加える
```
datas=[('student.csv', '.'),],
```

2. specファイルを使ってexeを作成
```
pyinstaller XXX.spec
```
