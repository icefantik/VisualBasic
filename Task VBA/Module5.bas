Attribute VB_Name = "Module5"
Rem �. �. �������. 1000 ����� �� ����������������, ������-��-���� 2004.
Rem �������: ������� ϸ��
Rem If12. ���� ��� �����. ����� ���������� �� ���.

Function Min(a As Integer, b As Integer) As Integer
    If a < b Then
        Min = a
    Else
        Min = b
    End If
End Function

Sub If12()
    Dim a As Integer, b As Integer, c As Integer
    a = Inputbox("")
    b = Inputbox("")
    c = Inputbox("")
    a = Min(a, b)
    b = Min(b, c)
    If a < b Then
        MsgBox (a)
    Else
        MsgBox (b)
    End If
End Sub
