Attribute VB_Name = "Module1"
Sub number()
Randomize
Do
a = Int(4 * Rnd + 1)
b = Int(4 * Rnd + 1)
Cells(a, b).Select
Loop While (Selection <> 0)
Selection.Value = 2
End Sub
Sub start()
Application.OnKey "{UP}", "UP"
Application.OnKey "{DOWN}", "DOWN"
Application.OnKey "{LEFT}", "LEFT"
Application.OnKey "{RIGHT}", "RIGHT"
number
End Sub
Sub LEFT()
For i = 1 To 4
    For j = 1 To 4
    If Cells(i, j) <> 0 Then
    'przesun liczby
        For n = 1 To (j - 1)
            If Cells(i, j - n) = 0 Then
                Cells(i, j - n) = Cells(i, j - n + 1)
                Cells(i, j - n + 1) = Empty
                'problem w sytuacji 0022 oraz 0202
'            ElseIf Cells(i, j - (j - n)) = Cells(i, j - (j - n) + 1) And Cells(i, j - (j - n)) <> 0 Then
'            Cells(i, j - (j - n)).Select
'                Cells(i, j - (j - n)) = 2 * Cells(i, j - (j - n))
'                Cells(i, j - (j - n) + 1) = Empty
            End If
        Next n
        'problem w sytuacji 8442
'        For m = 1 To (j - 1)
'            If Cells(i, m) = Cells(i, m + 1) And Cells(i, m) <> 0 Then
'                Cells(i, m) = 2 * Cells(i, m)
'                Cells(i, m + 1) = Empty
'            End If
'        Next m
    End If
    Next j
    'polacz liczby
        For m = 1 To 4
            If Cells(i, m) = Cells(i, m + 1) And Cells(i, m) <> 0 Then
                Cells(i, m) = 2 * Cells(i, m)
                Cells(i, m + 1) = Empty
            End If
        Next m
        'przesun liczby po polaczeniu
        For n = 1 To 4
            If Cells(i, n) = 0 Then
                Cells(i, n) = Cells(i, n + 1)
                Cells(i, n + 1) = Empty
            End If
        Next n
Next i
number
End Sub

Sub UP()
For j = 1 To 4
    For i = 1 To 4
    If Cells(i, j) <> 0 Then
        For n = 1 To (i - 1)
            If Cells(i - n, j) = 0 Then
                Cells(i - n, j) = Cells(i - n + 1, j)
                Cells(i - n + 1, j) = Empty
'            ElseIf Cells(i - (i - n), j) = Cells(i - (i - n) + 1, j) And Cells(i - (i - n), j) <> 0 Then
'            Cells(i - (i - n), j).Select
'                Cells(i - (i - n), j) = 2 * Cells(i - (i - n), j)
'                Cells(i - (i - n) + 1, j) = Empty
            End If
        Next n
'        For m = 1 To (i - 1)
'            If Cells(i - (i - m), j) = Cells(i - (i - m) + 1, j) And Cells(i - (i - m), j) <> 0 Then
'                Cells(i - (i - m), j) = 2 * Cells(i - (i - m), j)
'                Cells(i - (i - m) + 1, j) = Empty
'            End If
'        Next m
    End If
    Next i
    For m = 1 To 4
        If Cells(m, j) = Cells(m + 1, j) And Cells(m, j) <> 0 Then
            Cells(m, j) = 2 * Cells(m, j)
            Cells(m + 1, j) = Empty
        End If
    Next m
    For n = 1 To 4
        If Cells(n, j) = 0 Then
            Cells(n, j) = Cells(n + 1, j)
            Cells(n + 1, j) = Empty
        End If
    Next n
Next j
number
End Sub

Sub RIGHT()
For i = 4 To 1 Step -1
    For j = 4 To 1 Step -1
    If Cells(i, j) <> 0 Then
        For n = 1 To (3 - (j - 1))
            If Cells(i, j + n) = 0 Then
                Cells(i, j + n) = Cells(i, j + n - 1)
                Cells(i, j + n - 1) = Empty
            End If
        Next n
'        For m = (3 - (j - 1)) To 1 Step -1
'        If Cells(i, j + m) = Cells(i, j + m - 1) And Cells(i, j + m - 1) <> 0 Then
'            Cells(i, j + m) = 2 * Cells(i, j + m)
'            Cells(i, j + m - 1) = Empty
'        End If
'        Next m
    End If
    Next j
    For m = 4 To 2 Step -1
        If Cells(i, m) = Cells(i, m - 1) And Cells(i, m - 1) <> 0 Then
            Cells(i, m) = 2 * Cells(i, m)
            Cells(i, m - 1) = Empty
    End If
    Next m
    For n = 4 To 2 Step -1
        If Cells(i, n) = 0 Then
            Cells(i, n) = Cells(i, n - 1)
            Cells(i, n - 1) = Empty
        End If
    Next n
Next i
number
End Sub

Sub DOWN()
For j = 4 To 1 Step -1
    For i = 4 To 1 Step -1
    If Cells(i, j) <> 0 Then
        For n = 1 To (3 - (i - 1))
            If Cells(i + n, j) = 0 Then
                Cells(i + n, j) = Cells(i + n - 1, j)
                Cells(i + n - 1, j) = Empty
            End If
        Next n
'        For m = (3 - (i - 1)) To 1 Step -1
'        If Cells(i + m, j) = Cells(i + m - 1, j) And Cells(i + m - 1, j) <> 0 Then
'            Cells(i + m, j) = 2 * Cells(i + m, j)
'            Cells(i + m - 1, j) = Empty
'        End If
'        Next m
    End If
    Next i
    For m = 4 To 2 Step -1
        If Cells(m, j) = Cells(m - 1, j) And Cells(m - 1, j) <> 0 Then
            Cells(m, j) = 2 * Cells(m, j)
            Cells(m - 1, j) = Empty
        End If
    Next m
    For n = 4 To 2 Step -1
        If Cells(n, j) = 0 Then
            Cells(n, j) = Cells(n - 1, j)
            Cells(n - 1, j) = Empty
        End If
    Next n
Next j
number
End Sub
