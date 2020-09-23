VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export popup list"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   2655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "PopUps"
      Top             =   3240
      Width           =   1815
   End
   Begin VB.DriveListBox drv 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.DirListBox dir 
      Height          =   2565
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   ".PPS"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Choose filename:"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Open dir.Path & "\" & txtFileName.Text & ".PPS" For Output As #1
For i = 0 To frmMain.lstRemoves.ListCount - 1
    Print #1, frmMain.lstRemoves.List(i)
Next i
Close #1
Unload Me
End Sub

Private Sub drv_Change()
dir.Path = drv
End Sub
