Attribute VB_Name = "Module13"
Rem �. �. �������. 1000 ����� �� ����������������, ������-��-���� 2004.
Rem ������� ϸ��
Rem Array 12. ��� ������ A ������� N (N - ������ �����). ������� ��� �������� � ������� �������� � ������� ����������� �������
Rem A2, A4, A6, ..., An. �������� ��������� �� ������������.
 
Sub Array12()
    Dim i As Integer, n As Integer, num As Integer
    n = Inputbox("")
    ReDim a(n) As Integer
    For i = 0 To n
        num = Inputbox("")
        a(i) = num
    Next i
    
    For i = 0 To n Step 2
        MsgBox (a(i))
    Next i
End Sub
