Attribute VB_Name = "Module4"
Rem М. Э. Абрамян. 1000 ЗАДАЧ ПО ПРОГРАММИРОВАНИЮ, Ростов-на-Дону 2004.
Rem Решения: Митюшин Пётр
Rem Boolean 12. Даны три целых числа: A, B, C. Проверить истиность высказывания "Каждое из чисел A, B, C положительное"

Sub Boolean12()
    Dim a As Integer, b As Integer, c As Integer
    a = Inputbox("")
    b = Inputbox("")
    c = Inputbox("")
    If a >= 0 And b >= 0 And c >= 0 Then
        MsgBox (True)
    Else
        MsgBox (False)
    End If
End Sub
