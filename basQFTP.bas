Attribute VB_Name = "basQFTP"
Option Explicit

Public oLog As Collection

Sub Main()
On Error GoTo cmdClose_Click_Error

Set oLog = New Collection

    frmMain.Show

cmdClose_Click_Resume:
    Exit Sub
cmdClose_Click_Error:
    MsgBox Error$, vbInformation
    Resume cmdClose_Click_Resume
End Sub

Function ftpConnectToSite() As Boolean
    
End Function
