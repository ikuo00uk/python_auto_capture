## 自動キャプチャ  

pythonで指定したサイトのキャプチャを取得し、エクセルに保存する  

### Env
`python 3`  

### How to Use
初回環境構築　
`sh init.sh`  
`source env/bin/activate`  
`pip install -r requirement.txt`  

スクリプト実行  
`python app.py`  

下記が聞かれるので選択する  
`PC/SPのキャプチャを取得しますか？ [y/N]: y` => PC・SP両方キャプチャする  
`PCのみキャプチャを取得しますか？ NOの場合SPのみ保存します [y/N]` => yesならPC, NoならSPのみ  

仮想環境から出る  
`deactivate`  

### Dir
```
app.py- 実行python    
init.sh- python仮想環境作成シェル  
requirement.txt- pipのインストールするライブラリ  
url.xlsx- キャプチャ対象URL格納エクセル   
capture- 実行後キャプチャ管理ディレクトリ  
```

### エクセル
A列 ID
B列 ページ名
C列 URL
に情報を記載
