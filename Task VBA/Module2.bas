Attribute VB_Name = "Module2"
Rem �. �. �������. 1000 ����� �� ����������������, ������-��-���� 2004.
Rem �������: ������� ϸ��
Rem Begin12. ���� ������ �������������� ������������ a � b. ����� ��� ���������� c � �������� P:
Rem c = (a2 + b2)1/2, P = a + b + c.

Sub Begin12()
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    a = Inputbox("")
    b = Inputbox("")
    c = (a ^ 2 + b ^ 2) ^ 0.5
    p = a + b + c
    MsgBox (p)
End Sub
