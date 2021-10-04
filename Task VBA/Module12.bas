Attribute VB_Name = "Module12"
Rem М. Э. Абрамян. 1000 ЗАДАЧ ПО ПРОГРАММИРОВАНИЮ, Ростов-на-Дону 2004.
Rem Решения: Митюшин Пётр
Rem Minmax12. Дано целое число N и набор из N чисел. Найти минимальное положительное число из данного набора.
Rem Если положительные числа в наборе отсутсвуют, то вывести 0

Sub Minmax12()
    Dim minNum, i, n, num As Integer
    minNum = 10000
    n = Inputbox("")
    For i = 1 To n
        num = Inputbox("")
        If minNum > num Then
            minNum = num
        End If
        Next i
    MsgBox (minNum)
End Sub
