# 条件によって処理を変える (条件分岐)
if文などを利用して条件によって処理内容を変える方法について

## if 文
ある条件に合うときのみ処理を行うときにif文を利用します。

### 例1) 整数の x が 10のときのみ実行する

```vb
If x = 10 Then
    Debug.Print "xは10です"
End if 
```

※ 値が同じであるかをチェックするときは
   他のプログラミング言語は == ですがVBAは= は1つです


### 例2) xが10もしくは 20のとき実行する。そうでないときは別の処理を実行する

```vb
If x = 10 Or x = 20 then 
    Debug.Print "xは10火20です"
Else
    Debug.Print "それ以外です"
End If
```

### 例3) 文字列s に x が10の時 "A"、 xが20以上の時 "B"、 それ以外のときは "other"を入力する

```vb
Dim s As String
If x = 10 Then
    s = "A"
ElseIf x >= 20 Then
    s = "B"
Else
    s = "other"
End If
```

### 例4) xが10以上20未満なら処理1を実行する。そうでないときxが0でなければ別の処理を実行する

```vb
If x >= 10 And x < 20 Then
    Debug.Print "処理1"
ElseIf x <> 0 Then
    Debug.Print "xは0ではない"
End If
```

### 例5) value が数値でなかったら処理を実行する

```vb
Dim value   'valueは何でも入力できるVariant型
~ここでvalueに値を入力する~
If Not IsNumeric(value) Then
   Debug.Print "valueは数値ではない"
End If
```

-------------------------------

## select case 文
ある変数の値によって処理を変えるものが複数ある時select case の利用を検討ください。
if、 then、 elseif がたくさん続くときに便利です。

### 例1) 変数valueが red, blue, green のときにその日本語の色名を文字列sに入力する

```vb
Dim s As String
Select Case value 
  Case "red": s = "赤"
  Case "blue": s = "青"
  Case "green": s = "緑"
End Select
```

### 例2) 日付の値 value が 曜日によって処理を変える

```vb
If IsDate(value) Then  ' valueが日付形式だったら
  Select Case WeekDay(value)     '曜日を数字で取得(1=日曜 7=土曜)
    Case 1: Debug.Print "日曜日です"
    Case 5,6: Debug.Print "週末に近い(木金)です"  
    Case 7: Debug.Print "土曜日です"
    Case Else: Debug.Print "他の平日です"
  End Select
End If
``` 

### Caseで指定する方法のその他の例

```vb
値が 1～3 の場合:   Case 1 To 3, Is >= 10
値が10以上の場合:   Case Is >= 10
```

