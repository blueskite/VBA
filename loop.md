## 繰り返し処理
よく使う繰り返し処理をまとめます。

### 並んだセルから順番に読み取る&別のセルに書き込む
Ex.) A1～A10までの値を順に表示。B列にコピーする

```
Dim r as Range
For Each r in Range("A1:A10")
  Debug.Print r.Value
  r.Offset(0, 1).Value = r.Value
Next
```
rにはA1セルから順に入力されます。
Debug.Printにてrの値を、イミディエイトウィンドウに出力します
r.Offset(0, 1)は rのセルから行方向、列方向にずらす数を指定します。この場合は列を+1するのでB列になります

### VLOOKUPのように2つの表を参照する
Ex.) A列にID, B列に名前が入力されており、 E列にIDが順不同に並んでいるとします。このときF列にA,B列に定義された名前を表示したい場合

```
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

### 今後追加