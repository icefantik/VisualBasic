Attribute VB_Name = "Module7"
Rem �. �. �������. 1000 ����� �� ����������������, ������-��-���� 2004.
Rem �������: ������� ϸ��
Rem For12. ���� ����� ����� N (> 0). ����� ������������: 1.1 * 1.2 * 1.3 * ...

Sub For12()
    Dim n As Integer, i As Integer, znach As Double, res As Double
    res = 1
    znach = 1.1
    For i = 0 To n Step 1
        res = res * (znach + 0.1)
    MsgBox (res)
End Sub

