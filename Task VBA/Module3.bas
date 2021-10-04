Attribute VB_Name = "Module3"
Rem М. Э. Абрамян. 1000 ЗАДАЧ ПО ПРОГРАММИРОВАНИЮ, Ростов-на-Дону 2004.
Rem Решения: Митюшин Пётр
Rem Integer12. Дано трехзначное число. Вывести число, полученное при прочтении исходного числа справа налево.

Sub Integer12()
    Dim num As String
    num = Inputbox("")
    MsgBox (StrReverse(num))
End Sub
