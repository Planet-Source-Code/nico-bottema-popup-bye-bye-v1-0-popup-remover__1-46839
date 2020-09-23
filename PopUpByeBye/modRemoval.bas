Attribute VB_Name = "ModAll"
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Const WM_CLOSE = &H10

Public Function CheckWindows(ByVal hWnd As Long, ByVal lpData As Long) As Long
CheckWindows = 1
Dim WindowCaption As String, Ret As Long

WindowCaption = Space(GetWindowTextLength(hWnd) + 1)
Ret = GetWindowText(hWnd, WindowCaption, GetWindowTextLength(hWnd) + 1)
WindowCaption = Left(WindowCaption, Ret)
If InStr(WindowCaption, "Internet Explorer") Then
    For i = 0 To frmMain.lstRemoves.ListCount - 1
        If WindowCaption = frmMain.lstRemoves.List(i) Then
            Call PostMessage(hWnd, WM_CLOSE, 0, 0)
            Exit Function
        End If
    Next i
    For i = 0 To frmMain.lstPopUps.ListCount - 1
        If WindowCaption = frmMain.lstPopUps.List(i) Then
            Exit Function
        End If
    Next i
    frmMain.lstPopUps.AddItem WindowCaption
End If
End Function

Public Function CloseWindows(ByVal hWnd As Long, ByVal lpData As Long) As Long
CloseWindows = 1
Dim WindowCaption As String, Ret As Long

WindowCaption = Space(GetWindowTextLength(hWnd) + 1)
Ret = GetWindowText(hWnd, WindowCaption, GetWindowTextLength(hWnd) + 1)
WindowCaption = Left(WindowCaption, Ret)
If InStr(WindowCaption, "Internet Explorer") Then
    Call PostMessage(hWnd, WM_CLOSE, 0, 0)
    Exit Function
End If
End Function

Public Sub SaveList()
Open App.Path & "\RemList.DAT" For Output As #1
For i = 0 To frmMain.lstRemoves.ListCount - 1
    Print #1, frmMain.lstRemoves.List(i)
Next i
Close #1
End Sub

