Attribute VB_Name = "Module15"
Rem �. �. �������. 1000 ����� �� ����������������, ������-��-���� 2004.
Rem ������� ϸ��
Rem ���� �������� ������ S � ����� ����� N (> 0). ������� ������, ���������� ������� ������ S, ����� �������� ��������� �� N �������� �*� (���������).

Sub String12()
    Dim str, zv, nwStr As String, i, n As Integer
    n = Inputbox("")
    str = Inputbox("")
    For i = 1 To n
        zv = zv + "*"
    Next i
    For i = 1 To Len(str)
        Rem MsgBox (Mid(str, i, 1))
        nwStr = nwStr + Mid(str, i, 1) + zv
    Next i
    MsgBox (nwStr)
End Sub
