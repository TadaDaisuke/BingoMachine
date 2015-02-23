Attribute VB_Name = "Module1"
Option Explicit

Dim NumbersCell As Range
Dim CurrentIndexCell As Range
Dim DisplayCell As Range

Sub Start()

    Dim Result As Integer
    If MsgBox("さいしょからやりますか？", vbYesNo) = vbNo Then Exit Sub

    Dim i As Integer
    Set NumbersCell = Sheets("Hide").Range("Numbers")
    Set CurrentIndexCell = Sheets("Hide").Range("CurrentIndex")
    Set DisplayCell = Sheets(1).Cells(1, 1)
    
    ' 初期値設定
    For i = 1 To 75
        NumbersCell(i).Value = i
        ' 結果表示欄を初期化する
        KekkaCell(i).Value = i
        KekkaCell(i).Interior.Pattern = xlNone
    Next i
    DisplayCell.Value = "スタート！"
    CurrentIndexCell.Value = 0
    
    ' 数字をシャッフルする
    Dim a As Integer
    Dim b As Integer
    Dim work As Integer
    Randomize
    For i = 1 To 3000
        a = Int(Rnd * 75) + 1
        b = Int(Rnd * 75) + 1
        work = NumbersCell(a).Value
        NumbersCell(a).Value = NumbersCell(b).Value
        NumbersCell(b).Value = work
    Next i

End Sub

Private Function KekkaCell(num As Integer) As Range
    Dim r As Integer
    Dim c As Integer
    r = (num - 1) Mod 15 + 3
    c = Int((num - 1) / 15) + 1
    Set KekkaCell = Sheets(1).Cells(r, c)
End Function

Sub NextNumber()

    Set NumbersCell = Sheets("Hide").Range("Numbers")
    Set CurrentIndexCell = Sheets("Hide").Range("CurrentIndex")
    Set DisplayCell = Sheets(1).Cells(1, 1)
    CurrentIndexCell.Value = CurrentIndexCell.Value + 1

    If CurrentIndexCell.Value > 75 Then
        DisplayCell.Value = "おしまい。"
        Exit Sub
    End If
    
    DisplayCell.Value = NumbersCell(CurrentIndexCell.Value).Value
    KekkaCell(NumbersCell(CurrentIndexCell.Value).Value).Interior.Color = RGB(255, 255, 0)
    If CurrentIndexCell.Value > 1 Then
        KekkaCell(NumbersCell(CurrentIndexCell.Value - 1).Value).Interior.Color = RGB(173, 255, 47)
    End If

End Sub
