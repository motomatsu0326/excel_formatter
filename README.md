## エクセル整形くん

## 事前準備
1. python、pipをインストール
2. pipenvをインストール
```
$ pip install pipenv
```
3. pipenv仮想環境を作成
```
pipenv install
```

## 実行方法
1. ルートパスにExcelファイルを配置
2. run.shを実行
3. Excelファイルが下記ルールでフォーマットされる
```
・全てのシートのフォーカスをA1に変更
・アクティブシートを先頭シートに変更
・表示倍率を全て100％に変更
・全てのシートを表示状態に変更
```

## 参考
https://qiita.com/tasogarei/items/13cd236bce309b23c455
