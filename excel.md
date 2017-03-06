# Excelで利用する
Excelで計算式が長くなったものをVBAの関数にすることができます。それにより見やすくなったりミスを防ぎやすくなります。

## Excel関数をVBAにする

### 例) sum関数をExcelのVBAで表現する
Excelのセルに = SUM(A1:A10) と入力されているものとします。
この範囲は常に固定して使いたい時 =Total() と自分で定義した関数をExcelの計算式に使うことができます。

```vb
' Excelの計算式をコメントとして記載しておくと便利です
' = SUM(A1:A10)
Function Total() as Long  ' 入力値はなし、出力は整数とします
    Total = WorksheetFunction.Sum(Range("A1:A10")) 
End Function
```

* 関数で値を返すときは 関数名 = 返す値と記述します
* Excelの関数をVBAで利用するときはWorksheetFunction.XXXX とします
* ExcelのセルでA1:A10のようにセルを指定している部分は Range("XXXX") と記述します

--------------------------------

### 例) match, indexで参照しているものをVBAで表現する
id -> 名前 となっている表を参照して あるidに対する名前を取得するユーザ定義関数です。

```vb
' = index("B1:B12", match(D1, "A1:A12", 0), 1)
Function getName(id as String) as String
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Range
    For Each r In Range("A1:A12")
        dic(r.Value) = r.Offset(0, 1).Value
    Next

    getName = dic(id)
End Function
```

* まずは、A列のID → B列の 名前 という辞書を作成します
* あとはセルに対する値を dic(id)として取り出し関数名に返します

