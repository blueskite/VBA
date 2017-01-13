[TOP](.)

# 繰り返し処理
よく使う繰り返し処理をまとめます。


### 並んだセルから順番に読み取り合計する
例) A1～A10の数値を合算する

```vb
Dim total As Long, i As Long
For i = 1 To 10
  total = total + Cells(i, 1).Value
Next
```
iを1から10まで順に増加させながら セルの値をtotalに加算しています。

また、単純に1ずつ増やす以外の方法もできます。

```vb
For i = 1 To 10 Step 2   '1,3,5... と1つおきに増やす
For i = 10 To 1 Step -1  '10,9,8... と1つずつ小さくしていく
```

------------------

### 並んだセルから順番に読み取る&別のセルに書き込む
例) A列の値を順に表示。B列にコピーする (今回は10行目まで)

```vb
Dim r as Range
For Each r in Range("A1:A10")
  Debug.Print r.Value
  r.Offset(0, 1).Value = r.Value
Next
```
rにはA1セルから順に入力されます。
Debug.Printにてrの値を、イミディエイトウィンドウに出力します
r.Offset(0, 1)は rのセルから行方向、列方向にずらす数を指定します。この場合は列を+1するのでB列になります

------------------

### VLOOKUPのように2つの表を参照する
例) A列にID, B列に名前が入力されており、 E列にIDが順不同に並んでいるとします。このときF列にA,B列に定義された名前を表示したい場合

```vb
Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
Dim r As Range

' A列のIDに対する名前を辞書として保存
For Each r In Range("A1:A10")
  dic(r.Value) = r.Offset(0, 1).Value 
Next

' F列に、E列のIDからわかる名前を dicから取得する
For Each r In Range("E1:E10")
  r.Offset(0, 1).Value = dic(r.Value) 
Next
```
dic はディクショナリ(もしくは連想配列)です。ディクショナリは一意のキー(key)に対する値(value)をプログラム実行中のみ保存しておくことができます。

------------------

### 並んだセルから順番に読み書き (数字で指定する方法)
例) A列の値を順に表示

```vb
Dim i as Long
For i = 1 to 10
  Debug.Print Cells(i, 1).Value
  Cells(i, 2).Value = Cells(i, 1).Value 'B列にA列の値をコピー
Next
```
iに参照したい行番号を指定します。
Cells(1,1) は Range("A1") と同じです。
