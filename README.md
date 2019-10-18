## 自動キャプチャ  

pythonで指定したサイトのキャプチャを取得し、エクセルに保存する  

### Env
`python 3`  

### How to Use
`sh init.sh`  
`source env/bin/activate`  
`pip install -r requirement.txt`  

`python app.py`  

### Dir  
app.py- 実行python  
init.sh- python仮想環境作成シェル
requirement.txt- pipのインストールするライブラリ
url.xlsx- キャプチャ対象URL格納エクセル
capture- 実行後キャプチャ管理ディレクトリ

### エクセル
A列 ID
B列 ページ名
C列 URL

