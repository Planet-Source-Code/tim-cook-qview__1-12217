VERSION 5.00
Begin VB.Form frmLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connection Log"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClearLog 
      Caption         =   "Clear Log"
      Height          =   360
      Left            =   2355
      TabIndex        =   4
      Top             =   4590
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveLog 
      Caption         =   "&Save log"
      Height          =   360
      Left            =   3600
      TabIndex        =   3
      Top             =   4590
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   360
      Left            =   4845
      TabIndex        =   2
      Top             =   4590
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Log Entries"
      Height          =   4485
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   6015
      Begin VB.ListBox List1 
         Height          =   4155
         Left            =   75
         TabIndex        =   1
         Top             =   240
         Width           =   5865
      End
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
On Error GoTo cmdClose_Click_Error

    Unload Me

cmdClose_Click_Resume:
    Exit Sub
cmdClose_Click_Error:
    MsgBox Error$, vbInformation
    Resume cmdClose_Click_Resume
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
