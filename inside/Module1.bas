Attribute VB_Name = "Module1"
Sub protect_sheet_check()
    Dim strPassword As String
    strPassword = InputBox("Enter the password for the worksheet")
    MsgBox strPassword
    Worksheets("Sheet2").Protect Password:=strPassword, Scenarios:=True
End Sub

Sub unprotect_sheet_check()
    ActiveSheet.Unprotect
End Sub

