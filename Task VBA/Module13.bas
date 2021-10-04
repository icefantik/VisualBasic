Attribute VB_Name = "Module13"
Rem М. Э. Абрамян. 1000 ЗАДАЧ ПО ПРОГРАММИРОВАНИЮ, Ростов-на-Дону 2004.
Rem Митюшин Пётр
Rem Array 12. Дан массив A размера N (N - чётное число). Вывести его элементы с чётными номерами в порядке возрастания номеров
Rem A2, A4, A6, ..., An. Условные выражения не использовать.
 
Sub Array12()
    Dim i As Integer, n As Integer, num As Integer
    n = Inputbox("")
    ReDim a(n) As Integer
    For i = 0 To n
        num = Inputbox("")
        a(i) = num
    Next i
    
    For i = 0 To n Step 2
        MsgBox (a(i))
    Next i
End Sub
