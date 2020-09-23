VERSION 5.00
Begin VB.Form frmImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import list"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.ListBox lstPopUpLists 
      Height          =   2985
      Left            =   2760
      TabIndex        =   2
      Top             =   0
      Width           =   2535
   End
   Begin VB.DirListBox DirList 
      Height          =   2565
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.DriveListBox drv 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdImport_Click()
Dim TempString As String
For i = 0 To lstPopUpLists.ListCount - 1
    If lstPopUpLists.Selected(i) = True Then
        Open DirList.List(DirList.ListIndex) & lstPopUpLists.List(i) & ".PPS" For Input As #1
        While Not EOF(1)
            Input #1, TempString
            For j = 0 To frmMain.lstRemoves.ListCount - 1
                If frmMain.lstRemoves.List(j) = TempString Then
                    GoTo NextPopUp
                End If
            Next j
            frmMain.lstRemoves.AddItem TempString
NextPopUp:
        Wend
        Close #1
        Call SaveList
        Unload Me
        Exit Sub
    End If
Next i
MsgBox "You must select a file first", vbOKOnly, "Cannot import"
End Sub

Private Sub DirList_Click()
Dim a As String
lstPopUpLists.Clear
a = dir(DirList.List(DirList.ListIndex), vbDirectory)
While a <> ""
    DoEvents
    If Right(a, 4) = ".PPS" Then
        lstPopUpLists.AddItem Left(a, Len(a) - 4)
    End If
    a = dir
Wend
End Sub

Private Sub drv_Change()
DirList.Path = drv
Call DirList_Click
End Sub

