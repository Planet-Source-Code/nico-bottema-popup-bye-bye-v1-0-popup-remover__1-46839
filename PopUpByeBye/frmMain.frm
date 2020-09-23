VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PopUp Bye Bye"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrWindowCheck 
      Interval        =   300
      Left            =   2640
      Top             =   5640
   End
   Begin VB.CommandButton cmdUnlExit 
      Caption         =   "Unload and exit"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Explorer windows"
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   5295
      Begin VB.CommandButton cmdCloseAll 
         Caption         =   "Close all"
         Height          =   255
         Left            =   4200
         TabIndex        =   9
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdAddAll 
         Caption         =   "Add all"
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   975
      End
      Begin VB.ListBox lstPopUps 
         Height          =   1425
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Remove list"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdImportList 
         Caption         =   "Import list"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdCleanList 
         Caption         =   "Clean list"
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdExportList 
         Caption         =   "Export list"
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2640
         Width           =   975
      End
      Begin VB.ListBox lstRemoves 
         Height          =   2205
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Label Label1 
      Caption         =   "nicobottema@hotmail.com"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Menu mnuTray 
      Caption         =   "None"
      Visible         =   0   'False
      Begin VB.Menu mnuShowWindow 
         Caption         =   "Show PopUp Bye Bye"
      End
      Begin VB.Menu mnuUnloadExit 
         Caption         =   "Unload and exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
For i = 0 To lstPopUps.ListCount - 1
    If lstPopUps.Selected(i) = True Then
        lstRemoves.AddItem lstPopUps.List(i)
        lstPopUps.Selected(i) = False
        Call SaveList
        Exit Sub
    End If
Next i
End Sub

Private Sub cmdAddAll_Click()
If MsgBox("This will close all Internet Explorer windows, are you sure ?", vbYesNo, "Add all popups") = vbYes Then
    For i = 0 To lstPopUps.ListCount - 1
        lstRemoves.AddItem lstPopUps.List(i)
    Next i
    Call SaveList
End If
End Sub

Private Sub cmdCleanList_Click()
If MsgBox("Are you sure ?", vbYesNo, "Clean list") = vbYes Then
    lstRemoves.Clear
    Call SaveList
End If
End Sub

Private Sub cmdCloseAll_Click()
If MsgBox("Are you sure you want to close all Internet Explorer windows ?", vbYesNo, "Close all Popups") = vbYes Then
   Call EnumWindows(AddressOf CloseWindows, hWnd)
End If
End Sub

Private Sub cmdExportList_Click()
Load frmExport
frmExport.Show
End Sub

Private Sub cmdImportList_Click()
Load frmImport
frmImport.Show
End Sub

Private Sub cmdRemove_Click()
For i = 0 To lstRemoves.ListCount - 1
    If lstRemoves.Selected(i) = True Then
        lstRemoves.RemoveItem (i)
        Call SaveList
        Exit Sub
    End If
Next i
End Sub

Private Sub cmdUnlExit_Click()
If MsgBox("Are you sure you want to close PopUp ByeBye ?", vbYesNo, "Close PopUp ByeBye") = vbYes Then
    End
End If
End Sub

Private Sub Form_Load()
Dim TempString As String
On Error GoTo NoFile
Open App.Path & "\RemList.DAT" For Input As #1
While Not EOF(1)
    Input #1, TempString
    lstRemoves.AddItem TempString
Wend
Close #1
NoFile:
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RespondToTray(X, Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Call AddToTray(Me.Icon, "PopUp Bye Bye", Me)
End Sub

Private Sub mnuShowWindow_Click()
Me.Visible = True
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mnuUnloadExit_Click()
If MsgBox("Are you sure you want to close PopUp ByeBye ?", vbYesNo, "Close PopUp ByeBye") = vbYes Then
    Shell_NotifyIcon NIM_DELETE, nid
    End
End If
End Sub

Private Sub tmrWindowCheck_Timer()
Rego:
For i = 0 To lstPopUps.ListCount - 1
    If lstPopUps.Selected(i) = False Then
        lstPopUps.RemoveItem (i)
        GoTo Rego:
    End If
Next i
Call EnumWindows(AddressOf CheckWindows, hWnd)
End Sub

