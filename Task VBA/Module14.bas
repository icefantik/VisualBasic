Attribute VB_Name = "Module14"
Rem �. �. �������. 1000 ����� �� ����������������, ������-��-���� 2004.
Rem ������� ϸ��
Rem ���� ������� ������� M x N. ������� �� �������� � ��������� �������: ������ ������� ������ ����, ������ ������� ����� �����,
Rem ������ ������� ������ ����, �������� ������� ����� ����� � �.�.

Sub Matrix12()
    Dim i As Integer, j As Integer, m As Integer, n As Integer
    m = Inputbox("")
    n = Inputbox("")
    Dim Matrix(m, n) As Integer
    For i = 0 To m
        For j = 0 To n
            Matrix(i, j) = Inputbox("")
        Next j
    Next i
    For i = 0 To m
        For j = 0 To n
            If j Mod 2 = 0 Then
                MsgBox (Matrix(i, j))
            Else
                MsgBox (Matrix(n - i - 1, j))
        Next j
    Next i
End Sub

