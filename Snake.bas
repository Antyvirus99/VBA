Attribute VB_Name = "Module2"
Public x As Integer
Public y As Integer
Public score As Integer
Public BodyXIndex(1 To 100) As Integer
Public BodyYIndex(1 To 100) As Integer
Public BodyIndex As Integer
Public MoveIndex As Integer
Public size As Integer
Public szerokosc As Integer
Public board As Integer
Public wysokosc As Integer
Sub blockT()
Cells(1, 1).Select
Selection.Interior.Color = vbBlack
Selection.Offset(0, 1).Interior.Color = vbBlack
Selection.Offset(0, 2).Interior.Color = vbBlack
Selection.Offset(1, 1).Interior.Color = vbBlack
End Sub
Sub blockI()
Cells(4, 1).Select
Selection.Interior.Color = vbBlack
Selection.Offset(0, 1).Interior.Color = vbBlack
Selection.Offset(0, 2).Interior.Color = vbBlack
Selection.Offset(0, 3).Interior.Color = vbBlack
End Sub
Sub blockZ()
Cells(8, 1).Select
Selection.Interior.Color = vbBlack
Selection.Offset(0, 1).Interior.Color = vbBlack
Selection.Offset(1, 1).Interior.Color = vbBlack
Selection.Offset(1, 2).Interior.Color = vbBlack
End Sub
Sub blockS()
Cells(12, 1).Select
Selection.Offset(1, 0).Interior.Color = vbBlack
Selection.Offset(1, 1).Interior.Color = vbBlack
Selection.Offset(0, 1).Interior.Color = vbBlack
Selection.Offset(0, 2).Interior.Color = vbBlack
End Sub
Sub blockL()
Cells(16, 1).Select
Selection.Interior.Color = vbBlack
Selection.Offset(0, 1).Interior.Color = vbBlack
Selection.Offset(0, 2).Interior.Color = vbBlack
Selection.Offset(1, 0).Interior.Color = vbBlack
End Sub
Sub blockO()
Cells(20, 1).Select
Selection.Interior.Color = vbBlack
Selection.Offset(0, 1).Interior.Color = vbBlack
Selection.Offset(1, 0).Interior.Color = vbBlack
Selection.Offset(1, 1).Interior.Color = vbBlack
End Sub
Sub start()
Application.OnKey "{UP}", "UP"
Application.OnKey "{DOWN}", "DOWN"
Application.OnKey "{LEFT}", "LEFT"
Application.OnKey "{RIGHT}", "RIGHT"
Application.OnKey "~", "CHANGE"
With Worksheets("Arkusz1").Range("A1:T20")
   .Interior.Color = xlNone
   .Value = Empty
   .ColumnWidth = 1
   .RowHeight = 10
End With
BodyIndex = 1
score = 0
size = 1
MoveIndex = 1
szerokosc = 10
wysokosc = 10
board = szerokosc * wysokosc
Cells(11, 1).Value = score
x = 1
y = 1
Cells(x, y).Select
Selection.Interior.Color = vbBlue
Clear
BodyXIndex(BodyIndex) = x
BodyYIndex(BodyIndex) = y
food
Cells(15, 1).Value = Cells(1, 1).Width
Cells(16, 1).Value = Cells(1, 1).Height
'Show
End Sub


Sub UP()
If x > 1 Then
    x = x - 1
Else
    x = wysokosc
End If
Cells(BodyXIndex(MoveIndex), BodyYIndex(MoveIndex)).Select
MoveIndex = MoveIndex + 1
If MoveIndex > board Then MoveIndex = 1
Move (MoveIndex)
End Sub

Sub LEFT()
If y > 1 Then
    y = y - 1
Else
    y = szerokosc
End If
Cells(BodyXIndex(MoveIndex), BodyYIndex(MoveIndex)).Select
MoveIndex = MoveIndex + 1
If MoveIndex > board Then MoveIndex = 1
Move (MoveIndex)
End Sub

Sub RIGHT()
If y < szerokosc Then
    y = y + 1
Else
    y = 1
End If
Cells(BodyXIndex(MoveIndex), BodyYIndex(MoveIndex)).Select
MoveIndex = MoveIndex + 1
If MoveIndex > board Then MoveIndex = 1
Move (MoveIndex)
End Sub

Sub DOWN()
If x < wysokosc Then
    x = x + 1
Else
    x = 1
End If
Cells(BodyXIndex(MoveIndex), BodyYIndex(MoveIndex)).Select
MoveIndex = MoveIndex + 1
If MoveIndex > board Then MoveIndex = 1
Move (MoveIndex)
End Sub

Sub CHANGE()
If Selection.Interior.Color = vbRed Then
    Selection.Interior.Color = xlNone
    score = score + 1
    size = size + 1
    Cells(11, 1).Value = score
    food
End If

End Sub

Sub food()
Dim reason As Integer
Do
Randomize
a = Int(wysokosc * Rnd + 1)
b = Int(szerokosc * Rnd + 1)
Cells(a, b).Select
Loop While (Selection.Interior.Color = vbBlue Or Selection.Interior.Color = vbRed Or (BodyXIndex(MoveIndex) = a And BodyYIndex(MoveIndex) = b))
Selection.Interior.Color = vbRed

Cells(x, y).Select
End Sub
Sub Move(n As Integer)

Cells(12, 1) = n
BodyXIndex(n) = x
BodyYIndex(n) = y
If (n > size) Then
    If (BodyXIndex(n - size) <> 0 Or BodyYIndex(n - size) <> 0) Then
        Cells(BodyXIndex(n - size), BodyYIndex(n - size)).Select
        Selection.Interior.Color = xlNone
    End If
    BodyXIndex(n - size) = 0
    BodyYIndex(n - size) = 0
Else
    If (BodyXIndex(board + n - size) <> 0 Or BodyYIndex(board + n - size) <> 0) Then
        Cells(BodyXIndex(board + n - size), BodyYIndex(board + n - size)).Select
        Selection.Interior.Color = xlNone
    End If
    BodyXIndex(board + n - size) = 0
BodyYIndex(board + n - size) = 0
End If
Cells(BodyXIndex(n), BodyYIndex(n)).Select

CHANGE
If (Selection.Interior.Color = vbBlue) Then Cells(1, 1) = "Przegrana"
Selection.Interior.Color = vbBlue
'Show
End Sub
Sub Clear()
i = 1
While (i < 101)
BodyXIndex(i) = 0
BodyYIndex(i) = 0
i = i + 1
Wend
End Sub
Sub Show()
i = 1
While (i < 101)
Cells(13, i).Value = BodyXIndex(i)
Cells(14, i).Value = BodyYIndex(i)
i = i + 1
Wend
End Sub
Sub gradient()
n = 1
While (n < 256)
Cells(256 - n, 8).Select
Selection.Interior.Color = RGB(0, n, n)
n = n + 1
Wend
End Sub


