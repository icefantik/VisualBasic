Attribute VB_Name = "Module11"
Rem �. �. �������. 1000 ����� �� ����������������, ������-��-���� 2004.
Rem �������: ��������� �.�., 22.12.2014

Rem /*************************** Case ***************************/
Rem Case3. ��� ����� ������ � ����� ����� � ��������� 1�12 (1 � ������, 2 � ������� � �. �.).
Rem ������� �������� ���������������� ������� ���� (�����, ������, �����, �������).

Sub case3()
Dim x As Integer, s As String
    x = Inputbox("")
    Select Case x
        Case 1, 11, 12
            s = "����"
        Case 3 To 5
            s = "�����"
        Case 6 To 8
            s = "����"
        Case 9 To 11
            s = "�����"
        Case Else
            s = "�����������"
    End Select
    MsgBox (s)
End Sub

Rem Case4. ��� ����� ������ � ����� ����� � ��������� 1�12 (1 � ������, 2 � ������� � �. �.).
Rem ���������� ���������� ���� � ���� ������ ��� ������������� ����.

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
