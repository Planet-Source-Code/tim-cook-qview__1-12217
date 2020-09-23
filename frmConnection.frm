VERSION 5.00
Begin VB.Form frmConnection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connection too..."
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3825
      TabIndex        =   2
      Top             =   1500
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   2565
      TabIndex        =   1
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Site Info:"
      Height          =   1380
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   5010
   End
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_Click_Error
    
    Unload Me

cmdCancel_Click_Resume:
    Exit Sub
cmdCancel_Click_Error:
    MsgBox Error$, vbInformation
    Resume cmdCancel_Click_Resume
End Sub

Private Sub Form_Load()
On Error GoTo Form_Load_Error

    ftpCenterForm Me

Form_Load_Resume:
    Exit Sub
Form_Load_Error:
    MsgBox Error$, vbInformation
    Resume Form_Load_Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdCancel_Click
End Sub
