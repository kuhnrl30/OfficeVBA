Attribute VB_Name = "Module2"
Sub Protect_sheets()
 
    Dim wSheet          As Worksheet
    Dim Pwd             As String
 
    Pwd = InputBox("Enter your password to protect all worksheets", "Password Input")
    For Each wSheet In Worksheets
        wSheet.Protect Password:=Pwd
    Next wSheet
 
End Sub

Sub Unprotect_sheets()
 
    Dim wSheet          As Worksheet
    Dim Pwd             As String
 
    Pwd = InputBox("Enter your password to unprotect all worksheets", "Password Input")
    On Error Resume Next
    For Each wSheet In Worksheets
        wSheet.Unprotect Password:=Pwd
    Next wSheet
    If Err <> 0 Then
        MsgBox "You have entered an incorect password. All worksheets could not " & _
        "be unprotected.", vbCritical, "Incorect Password"
    End If
    On Error GoTo 0
 
End Sub


