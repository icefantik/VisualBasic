Attribute VB_Name = "Module16"
Rem �. �. �������. 1000 ����� �� ����������������, ������-��-���� 2004.
Rem ������� ϸ��
Rem File 12. ��� ���� ����� �����. ������� ��� ����� �����, ������ �� ������� �������� ������ ����� �� ��������� �����, � ������ � �������� (� ��� �� �������). ���� ������ ��� �������� ����� � �������� ����� �����������, �� ��������������� �������������� ���� �������� ������.

Sub File12()
Dim strLine, str As String, nums() As String, i, n As Integer
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set file1 = fs.CreateTextFile("D:\file1.txt", False)
    Set file2 = fs.CreateTextFile("D:\file2.txt", False)
    
    Open "D:\numbers.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, strLine
        nums = Split(strLine, " ")
        n = UBound(nums)
        For i = 0 To n
            If nums(i) Mod 2 = 0 Then
                file1.Write nums(i) + " "
            Else
                file2.Write nums(i) + " "
            End If
        Next i
    Loop
    Close #1
End Sub
