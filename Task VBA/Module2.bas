Attribute VB_Name = "Module2"
Rem М. Э. Абрамян. 1000 ЗАДАЧ ПО ПРОГРАММИРОВАНИЮ, Ростов-на-Дону 2004.
Rem Решения: Митюшин Пётр
Rem Begin12. Даны катеты прямоугольного треугольника a и b. Найти его гипотенузу c и периметр P:
Rem c = (a2 + b2)1/2, P = a + b + c.

Sub Begin12()
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    a = Inputbox("")
    b = Inputbox("")
    c = (a ^ 2 + b ^ 2) ^ 0.5
    p = a + b + c
    MsgBox (p)
End Sub
