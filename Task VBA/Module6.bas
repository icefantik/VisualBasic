Attribute VB_Name = "Module6"
Rem М. Э. Абрамян. 1000 ЗАДАЧ ПО ПРОГРАММИРОВАНИЮ, Ростов-на-Дону 2004.
Rem Решения: Митюшин Пётр
Rem Case12. Элементы окружности пронумерованы следующим образом: 1 - радиус R, 2 - диаметр D = 2 * R, 3 - длина L = 2 * PI * R
Rem 4 - площадь круга S = PI * R ^ 2. Дан номер одного из этих элементов и его значение. Вывести значения остальных элементов данной
Rem окружности (В том же порядке). В качестве значения PI использовать 3.14.

Sub Case12()
    Dim a As Integer, d As Integer, p As Integer, r As Integer, L As Integer, s As Integer
    a = Inputbox("")
    Select Case a
        Case 1
            r = x
            d = 2 * r
            L = 2 * 3.14 * r
            s = 3.14 * Sqr(r)
            MsgBox (d)
            MsgBox (L)
            MsgBox (s)
        Case 2
            r = x / 2
            d = x
            L = 2 * 3.14 * r
            s = 3.14 * Sqr(r)
            MsgBox (r)
            MsgBox (L)
            MsgBox (s)
        Case 3
            r = x / 2 * 3.14
            d = 2 * r
            L = x
            s = 3.14 * Sqr(r)
            MsgBox (r)
            MsgBox (d)
            MsgBox (s)
        Case 4
            r = Sqr(x / 3.14)
            d = 2 * r
            L = 2 * 3.14 * r
            s = x
            MsgBox (r)
            MsgBox (d)
            MsgBox (L)
End Sub
