# export-trello-task
## インストールするオープンソースライブラリ
### 必須
+ numpy
+ openyxl
### 推奨
+ virtualenv

virtualenvによる開発環境の設定は以下の手順で行う。
```
~$ virtualenv -p python3.7 venv #venvはディレクトリ名なので任意で変更可能
~$ source venv/bin/activate #Linux/MacOSはsourceで、Windowsの場合はsouceを付けずにバッチファイルを実行
~$ #以下pip installでvirtualenvに必要なライブラリを追加
```

あとは必要に応じて各自の環境で追加。setup.shでは開発環境で便利なオープンソースライブラリも勝手に追加される。
## 使い方
1. pip installを実行し、必要なライブラリを追加  
2. jsonsディレクトリを作成  
3. jsonsディレクトリにtrelloボードでエクスポートしたjsonファイルを追加していく  
4. main.pyを実行  
excelファイルはxlsxsディレクトリに自動で追加される。日付をファイル名にするため、その日にエクスポートされるファイルはその都度上書きされる。  
## プロジェクト構成

```
.
├── exportExcel
│　　└── export_excel.py
├── main.py
└── trelloAPI
　　　├── actions.py
　　　├── board.py
　　　├── card.py
　　　├── check_lists.py
　　　└── trello_api.py
``` 

### main.py
メインメソッド
### exportExcel (export_excel.py)
エクセルファイルへのエクスポート処理が記述されている。セルの値の設定、色、罫線などの設定もここで記述されている。
### trelloAPI
Trelloに関するデータの操作はここで行う。
+ actions.py
actionsのモデル
+ board.py
boardのモデル
+ card.py
cardのモデル
+ check_lists.py
check_listsのモデル
+ trello_api.py
trelloのデータを操作するクラス・関数が記述されている。  
ここでboardとその他データとの紐付けを行う。
