Attribute VB_Name = "Module10"
Rem �. �. �������. 1000 ����� �� ����������������, ������-��-���� 2004.
Rem �������: ������� ϸ��
Rem Func12. ������� ������� IsPowerN(K, N) ����������� ����, ������������ True, ���� ����� �������� K(> 0) �������� �������� �����
Rem N (> 1), � False � ��������� ������. ���� ����� N (> 1) � ����� �� 10 ����� ������������� �����. � �������  ������� IsPowerN
Rem ����� ���������� �������� ����� N � ������ ������.

Function IsPowerN(k As Integer, n As Integer) As Boolean
    Dim pow As Integer
    pow = 1
    While (pow < n)
        pow = pow * k
    Wend
    If pow = n Then
        IsPowerN = True
        MsgBox (True)
    Else
        IsPowerN = False
        MsgBox (False)
    End If
End Function
