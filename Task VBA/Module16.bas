Attribute VB_Name = "Module16"
Rem М. Э. Абрамян. 1000 ЗАДАЧ ПО ПРОГРАММИРОВАНИЮ, Ростов-на-Дону 2004.
Rem Митюшин Пётр
Rem File 12. Дан файл целых чисел. Создать два новых файла, первый из которых содержит четные числа из исходного файла, а второй — нечетные (в том же порядке). Если четные или нечетные числа в исходном файле отсутствуют, то соответствующий результирующий файл оставить пустым.

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
