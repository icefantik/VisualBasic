Attribute VB_Name = "Module7"
Rem М. Э. Абрамян. 1000 ЗАДАЧ ПО ПРОГРАММИРОВАНИЮ, Ростов-на-Дону 2004.
Rem Решения: Митюшин Пётр
Rem For12. Дано целое число N (> 0). Найти произведение: 1.1 * 1.2 * 1.3 * ...

Sub For12()
    Dim n As Integer, i As Integer, znach As Double, res As Double
    res = 1
    znach = 1.1
    For i = 0 To n Step 1
        res = res * (znach + 0.1)
    MsgBox (res)
End Sub

