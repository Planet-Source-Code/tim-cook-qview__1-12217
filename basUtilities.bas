Attribute VB_Name = "basUtilities"
Option Explicit

Public Const SHGFI_SMALLICON = &H1
Public Const MAX_PATH = 260
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const ILD_TRANSPARENT = &H1
Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
   Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
   Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal flags&) As Long

Public SHELL_FILE_INFO As SHFILEINFO

Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Sub ftpCenterForm(frmForm As Form)
On Error GoTo ftpCenterFormError
Dim lAdjustedHeight As Long

    frmForm.Left = Screen.Width / 2 - frmForm.Width / 2
    frmForm.Top = Screen.Height / 2 - frmForm.Height / 2 - 500

ftpCenterFormContinue:
    Exit Sub
ftpCenterFormError:
    MsgBox Error$, vbExclamation
    Resume ftpCenterFormContinue
End Sub

Public Function ftpRetractDirectory(sPathName As String) As String
On Error GoTo ftpRetractDirectory_Error
Dim sRetval As String
Dim lLength As String
Dim sTemp As String

lLength = Len(sPathName) - 1
sTemp = Mid(sPathName, 1, lLength)
sRetval = Mid(sTemp, 1, InStrRev(sTemp, "\"))

ftpRetractDirectory = sRetval

ftpRetractDirectory_Resume:
    Exit Function
ftpRetractDirectory_Error:
    MsgBox Error$, vbInformation
    Resume ftpRetractDirectory_Resume
End Function


Function ftpCurrentMachine() As Variant
On Error GoTo ftpCurrentMachineError
Dim sMachineName As String
Dim lRetVal As Long
    
    sMachineName = String(2048, 32)
    lRetVal = GetComputerName(sMachineName, Len(sMachineName) - 1)
    ftpCurrentMachine = Mid(sMachineName, 1, InStr(sMachineName, Chr(0)) - 1)

ftpCurrentMachineContinue:
    Exit Function
ftpCurrentMachineError:
    MsgBox Error$, vbExclamation
    Resume ftpCurrentMachineContinue
End Function
