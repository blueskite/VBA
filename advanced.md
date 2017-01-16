# もう少し詳しく
こちらの内容も含めて知っておくとVBAの作成に役立ちます

## ある部分の処理時間を計測したい
計測開始時点で時刻をとっておき、計測終了時点の時刻と引き算して出力することができます。

```vb
' Utilモジュールに関数を定義します (秒数は実数で管理)
Dim start As Double
Public Function TimeStart()
    start = Timer
End Function

Public Function TimeFinish()
    Debug.Print (Timer - start) & " sec elapsed."
End Function


' 利用方法
Sub Test()
    TimeStart 
    '～ 計測したい対象
    TimeFinish
End Sub
```

TimeFinishが呼ばれたときに経過時間が表示されます。

```
0.08984375 sec elapsed.
```
------------------

# シート全体の文字を削除
シート名を指定してそのセル全体の値を削除する方法です

```vb
Worksheets("シート名").Cells.ClearContents
```
------------------

# 配列の利用
VBAプログラム内で複数の要素をまとめるのにArray関数を利用することができます。

## 複数の要素に対して繰り返し処理をする

例) Arrayに3つのシート名をVBAに直書きしてセル全体の値を初期化する方法

```vb
Dim s As String
For Each s In Array("シート1", "シート3", "シート3")
    Worksheets(s).Cells.ClearContents
Next
```


------------------

# VBAのもととなるVisual Basicの特徴
Visual Basicは、Java ScriptなどC系の言語と比較して書き方に特徴があります。
ざっと理解するために違いを整理します。

* 書き方の違い
  * 関数名やメソッドは先頭を大文字にする
  * If文で値が同じかを比較するのは == ではなく = １つ。異なる場合は!=ではなく <>
  * {} は使わない。 End XXXXX や、Nextなどで 対象の範囲を定義する
  * 変数の定義は Dim 変数名 As 型。 As 型を省略したときは任意の値を入れられるVariantとして定義

* 挙動の違い
  * If文でAnd/Orを使って復数の比較をする時、CなどはAndにて最初のがFalseの場合は以降の比較はしないがVBAは行うので未定義とならない要注意
  * 関数にはSubと Functionの2つがある。マクロから呼び出せるのはSubのみ。値を返せるのはFunctionのみ



---------------------------------

## おまけ

### 未定義の変数があるとき必ずエラーにする
各プログラムの先頭に「Option Explict」を必ず記入してください。
これにより未定義の変数があるとプログラムエラーとなり、タイプミスを軽減することができます。

