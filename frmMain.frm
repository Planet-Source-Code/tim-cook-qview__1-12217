VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "QView"
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6285
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageCombo cboDrives 
      Height          =   330
      Left            =   15
      TabIndex        =   1
      Top             =   60
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      ImageList       =   "imgListTree"
   End
   Begin MSComctlLib.ImageList imgListTree 
      Left            =   5595
      Top             =   5535
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5272
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A1DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C98E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DC12
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":103C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12B7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1532E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17DFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvHostFiles 
      Height          =   5925
      Left            =   0
      TabIndex        =   0
      Top             =   405
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   10451
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgListTree"
      SmallIcons      =   "imgListTree"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Directory Contents"
         Object.Width           =   8916
      EndProperty
   End
   Begin VB.Menu mnuFileMain 
      Caption         =   "&File"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuMainSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditLocalHost 
         Caption         =   "&Local Host"
         Begin VB.Menu mnuLocalCreateDir 
            Caption         =   "&Create Directory"
         End
         Begin VB.Menu mnuLocalDeleteDir 
            Caption         =   "&Delete Directory"
         End
      End
      Begin VB.Menu mnuEditRemotehost 
         Caption         =   "&Remote Host"
         Begin VB.Menu mnuRemoteCreateDir 
            Caption         =   "&Create Directory"
         End
         Begin VB.Menu mnuRemoteDeleteDir 
            Caption         =   "&Delete Directory"
         End
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewSysTray 
         Caption         =   "&Enable System Tray"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSiteManager 
         Caption         =   "&Site Manager"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuToolBar 
         Caption         =   "&Toolbar"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewTransferQueue 
         Caption         =   "Transfer &Queue"
         Checked         =   -1  'True
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuViewLog 
         Caption         =   "&View Log"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuQueue 
      Caption         =   "&Queue"
      Begin VB.Menu mnuStartTransfer 
         Caption         =   "&Start Transfer"
      End
      Begin VB.Menu mnuClearQueue 
         Caption         =   "&Clear Queued Items"
      End
   End
   Begin VB.Menu mnuHelpMain 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Q-FTP"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvDrivePath As String
Private fsoHost As FileSystemObject

Public Sub ftpBuildListView(sStartDrive As String, Optional fRecurse As Boolean)
On Error GoTo ftpBuildListView_Error

Dim oFolder As Folder
Dim oFolders As Folders
Dim oDrive As Drive
Dim MyItem As ListItem
Dim MyItems As ListItems
Dim oFile As File
Dim oFiles As Files
Dim iImage As Integer

Set MyItems = Me.lvHostFiles.ListItems
Set fsoHost = New FileSystemObject

MyItems.Clear

If sStartDrive = "" Then
   sStartDrive = Mid(Trim(cboDrives.Text), 1, Len(cboDrives.Text) - 1) & "\"
   cboDrives.Text = sStartDrive
End If

Set oFolders = fsoHost.GetFolder(sStartDrive).SubFolders
Set oFiles = fsoHost.GetFolder(sStartDrive).Files
Set MyItem = MyItems.Add(, "UPDIR", "...", 12, 12)
    
    For Each oFolder In oFolders
        Set MyItem = MyItems.Add(, sStartDrive & oFolder.Name & "\", oFolder.Name, 9, 9)
    Next
    
    For Each oFile In oFiles
        Set MyItem = MyItems.Add(, sStartDrive & oFile.Name, oFile.Name, 11, 11)
    Next
    
ftpBuildListView_Resume:
    Exit Sub
ftpBuildListView_Error:
    Select Case Err.Number
        Case 76 'Path not found
            MyItems.Clear
            Set MyItem = MyItems.Add(, "BAD", "Check Drive Media", 1, 1)
        Case Else
            MsgBox Error$, vbInformation
    End Select
    Resume ftpBuildListView_Resume
End Sub

Private Sub cboDrives_Click()
    mvDrivePath = Trim(Mid(cboDrives.SelectedItem.Key, 1, 3)) & "\"
    cboDrives.Text = mvDrivePath
    ftpBuildListView mvDrivePath
End Sub

Private Sub cboDrives_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 13
        If Me.cboDrives.Text <> "" Then
            If Right(Me.cboDrives.Text, 1) = "\" Then
                mvDrivePath = Mid(Trim(cboDrives.Text), 1, Len(cboDrives.Text) - 1)
            End If
            mvDrivePath = Trim(cboDrives.Text) & "\"
            ftpBuildListView mvDrivePath
            Me.lvHostFiles.SetFocus
        Else
            MsgBox "Please enter a valid folder path!", vbInformation
        End If
    Case Else
        
End Select

End Sub

Private Sub Form_Load()
On Error GoTo Form_Load_Error

Dim MyFSO As FileSystemObject
Dim oDrive As Drive
Dim oDrives As Drives
Dim iImage As Integer


Set MyFSO = New FileSystemObject
Set oDrives = MyFSO.Drives

    ftpCenterForm Me
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    For Each oDrive In oDrives
        Select Case oDrive.DriveType
            Case 0 '"Unknown"
                iImage = 1
            Case 1 '"Removable"
                iImage = 2
            Case 2 '"Fixed"
                iImage = 3
            Case 3 '"Network"
                iImage = 4
            Case 4 '"CD-ROM"
                iImage = 5
            Case 5 '"RAM Disk"
                iImage = 6
        End Select
        
        cboDrives.ComboItems.Add , oDrive.Path, oDrive.Path & "\", iImage
    Next
    
Form_Load_Resume:
    Exit Sub
Form_Load_Error:
    MsgBox Error$, vbInformation
    Resume Form_Load_Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mnuFileExit_Click
End Sub

Private Sub lvHostFiles_DblClick()
On Error GoTo lvHostFiles_DblClick_Error
Dim MyItem As ListItem

If Me.lvHostFiles.ListItems.Count Then
    Select Case lvHostFiles.SelectedItem.Key
        Case "BAD"
               
        Case "UPDIR"
            mvDrivePath = ftpRetractDirectory(mvDrivePath)
            ftpBuildListView mvDrivePath
        Case Else
            If Right(lvHostFiles.SelectedItem.Key, 1) = "\" Then
                mvDrivePath = lvHostFiles.SelectedItem.Key
                ftpBuildListView lvHostFiles.SelectedItem.Key
            End If
    End Select
End If

Me.cboDrives.Text = mvDrivePath

lvHostFiles_DblClick_Resume:
    Exit Sub
lvHostFiles_DblClick_Error:
    MsgBox Error$, vbInformation
    Resume lvHostFiles_DblClick_Resume
End Sub

Private Sub mnuConnect_Click()
On Error GoTo mnuConnect_Click_Error

    frmConnection.Show 1

mnuConnect_Click_Resume:
    Exit Sub
mnuConnect_Click_Error:
    MsgBox Error$, vbInformation
    Resume mnuConnect_Click_Resume
End Sub

Private Sub mnuFileExit_Click()
On Error GoTo mnuFileExit_Click_Error
    
    Set oLog = Nothing
    Set fsoHost = Nothing

If Me.WindowState <> vbMinimized Then
    SaveSetting App.Title, "Settings", "MainLeft", Me.Left
    SaveSetting App.Title, "Settings", "MainTop", Me.Top
    SaveSetting App.Title, "Settings", "MainWidth", Me.Width
    SaveSetting App.Title, "Settings", "MainHeight", Me.Height
End If
    
    End

mnuFileExit_Click_Resume:
    Exit Sub
mnuFileExit_Click_Error:
    MsgBox Error$, vbInformation
    Resume mnuFileExit_Click_Resume
End Sub

Private Sub mnuHelpAbout_Click()
On Error GoTo mnuHelpAbout_Click_Error

    frmAbout.Show 1

mnuHelpAbout_Click_Resume:
    Exit Sub
mnuHelpAbout_Click_Error:
    MsgBox Error$, vbInformation
    Resume mnuHelpAbout_Click_Resume
End Sub

Private Sub mnuViewLog_Click()
On Error GoTo mnuViewLog_Click_Error

    frmLog.Show 1

mnuViewLog_Click_Resume:
    Exit Sub
mnuViewLog_Click_Error:
    MsgBox Error$, vbInformation
    Resume mnuViewLog_Click_Resume
End Sub
