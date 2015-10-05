Attribute VB_Name = "Module1"
Public Sub BacaFile()
Dim MyString As String
Open App.Path & "\database\voucher.txt" For Input As #1 ' Open file for input.
Do While Not EOF(1) ' Loop until end of file.
    Input #1, MyString
    MsgBox MyString
Loop
Close #1    ' Close file.
End Sub
