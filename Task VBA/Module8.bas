Attribute VB_Name = "Module8"
Rem �. �. �������. 1000 ����� �� ����������������, ������-��-���� 2004.
Rem �������: ������� ϸ��
Rem While12. ���� ����� ����� N( > 1). ������� ���������� �� ����� ����� K, ��� ������� ����� 1 + 2 + ... + K ����� ������ ��� ����� N, � ���� ��� �����.

Sub While12()
    Dim n As Integer, k As Integer, temp As Integer
    n = Inputbox("")
    While Not (temp + k + 1) > n
        k = k + 1
        temp = temp + k
    MsgBox (k)
    MsgBox (temp)
End Sub

