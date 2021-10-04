Attribute VB_Name = "Module9"
Rem М. Э. Абрамян. 1000 ЗАДАЧ ПО ПРОГРАММИРОВАНИЮ, Ростов-на-Дону 2004.
Rem Решения: Митюшин Пётр
Rem Series12. Дан набор ненулевых целых чисел; признак его завершения — число 0. Вывести количество чисел в наборе.

Sub Series12()
    Dim num As Integer
    num = Inputbox("")
    While Not num = 0
        num = Inputbox("")
End Sub
