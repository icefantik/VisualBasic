Attribute VB_Name = "Module11"
Rem ћ. Ё. јбрам€н. 1000 «јƒј„ ѕќ ѕ–ќ√–јћћ»–ќ¬јЌ»ё, –остов-на-ƒону 2004.
Rem –ешени€: √ус€тинер Ћ.Ѕ., 22.12.2014

Rem /*************************** Case ***************************/
Rem Case3. ƒан номер мес€ца Ч целое число в диапазоне 1Ц12 (1 Ч €нварь, 2 Ч февраль и т. д.).
Rem ¬ывести название соответствующего времени года (Ђзимаї, Ђвеснаї, Ђлетої, Ђосеньї).

Sub case3()
Dim x As Integer, s As String
    x = Inputbox("")
    Select Case x
        Case 1, 11, 12
            s = "зима"
        Case 3 To 5
            s = "весна"
        Case 6 To 8
            s = "лето"
        Case 9 To 11
            s = "осень"
        Case Else
            s = "неизвестное"
    End Select
    MsgBox (s)
End Sub

Rem Case4. ƒан номер мес€ца Ч целое число в диапазоне 1Ц12 (1 Ч €нварь, 2 Ч февраль и т. д.).
Rem ќпределить количество дней в этом мес€це дл€ невисокосного года.

Sub case4()
Dim x As Integer, d As Integer
    x = Inputbox("")
    Select Case x
        Case 1, 3, 5, 7, 8, 10, 12
            d = 31
        Case 2
            d = 28
        Case 4, 6, 9, 11
            d = 30
        Case Else
            d = 0
    End Select
    MsgBox (d)
End Sub
