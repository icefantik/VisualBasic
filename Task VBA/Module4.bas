Attribute VB_Name = "Module4"
Rem �. �. �������. 1000 ����� �� ����������������, ������-��-���� 2004.
Rem �������: ������� ϸ��
Rem Boolean 12. ���� ��� ����� �����: A, B, C. ��������� ��������� ������������ "������ �� ����� A, B, C �������������"

Sub Boolean12()
    Dim a As Integer, b As Integer, c As Integer
    a = Inputbox("")
    b = Inputbox("")
    c = Inputbox("")
    If a >= 0 And b >= 0 And c >= 0 Then
        MsgBox (True)
    Else
        MsgBox (False)
    End If
End Sub
