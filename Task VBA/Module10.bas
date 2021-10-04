Attribute VB_Name = "Module10"
Rem М. Э. Абрамян. 1000 ЗАДАЧ ПО ПРОГРАММИРОВАНИЮ, Ростов-на-Дону 2004.
Rem Решения: Митюшин Пётр
Rem Func12. Описать функцию IsPowerN(K, N) логического типа, возвращающую True, если целый параметр K(> 0) является степенью числа
Rem N (> 1), и False в противном случае. Дано число N (> 1) и набор из 10 целых положительных чисел. С помощью  функции IsPowerN
Rem найти количество степеней числа N в данном наборе.

Function IsPowerN(k As Integer, n As Integer) As Boolean
    Dim pow As Integer
    pow = 1
    While (pow < n)
        pow = pow * k
    Wend
    If pow = n Then
        IsPowerN = True
        MsgBox (True)
    Else
        IsPowerN = False
        MsgBox (False)
    End If
End Function
