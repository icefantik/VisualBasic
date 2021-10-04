Attribute VB_Name = "Module8"
Rem М. Э. Абрамян. 1000 ЗАДАЧ ПО ПРОГРАММИРОВАНИЮ, Ростов-на-Дону 2004.
Rem Решения: Митюшин Пётр
Rem While12. Дано целое число N( > 1). Вывести наибольшее из целых чисел K, для которых сумма 1 + 2 + ... + K будет меньше или равна N, и саму эту сумму.

Sub While12()
    Dim n As Integer, k As Integer, temp As Integer
    n = Inputbox("")
    While Not (temp + k + 1) > n
        k = k + 1
        temp = temp + k
    MsgBox (k)
    MsgBox (temp)
End Sub

