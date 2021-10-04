Attribute VB_Name = "Module15"
Rem М. Э. Абрамян. 1000 ЗАДАЧ ПО ПРОГРАММИРОВАНИЮ, Ростов-на-Дону 2004.
Rem Митюшин Пётр
Rem Дана непустая строка S и целое число N (> 0). Вывести строку, содержащую символы строки S, между которыми вставлено по N символов «*» (звездочка).

Sub String12()
    Dim str, zv, nwStr As String, i, n As Integer
    n = Inputbox("")
    str = Inputbox("")
    For i = 1 To n
        zv = zv + "*"
    Next i
    For i = 1 To Len(str)
        Rem MsgBox (Mid(str, i, 1))
        nwStr = nwStr + Mid(str, i, 1) + zv
    Next i
    MsgBox (nwStr)
End Sub
