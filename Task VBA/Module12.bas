Attribute VB_Name = "Module12"
Rem �. �. �������. 1000 ����� �� ����������������, ������-��-���� 2004.
Rem �������: ������� ϸ��
Rem Minmax12. ���� ����� ����� N � ����� �� N �����. ����� ����������� ������������� ����� �� ������� ������.
Rem ���� ������������� ����� � ������ ����������, �� ������� 0

Sub Minmax12()
    Dim minNum, i, n, num As Integer
    minNum = 10000
    n = Inputbox("")
    For i = 1 To n
        num = Inputbox("")
        If minNum > num Then
            minNum = num
        End If
        Next i
    MsgBox (minNum)
End Sub
