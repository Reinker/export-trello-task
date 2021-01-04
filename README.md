# export-trello-task
## 使い方
1. pip installを実行  
2. jsonsディレクトリを作成  
3. jsonsディレクトリにtrelloボードでエクスポートしたjsonファイルを追加していく  
4. main.pyを実行  
excelファイルはxlsxsディレクトリに自動で追加される。日付をファイル名にするため、その日にエクスポートされるファイルはその都度上書きされる。  
## プロジェクト構成
main.py  
|  
| - exportExcel  
|   | - export_excel.py  
|  
| - trelloAPI  
    | - actions.py  
    | - board.py  
    | - card.py  
    | - check_lists.py  
    | - trello_api.py  
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
