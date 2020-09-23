Attribute VB_Name = "Module1"
Declare Function WindowFromPoint& Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long)
Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)

Public Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type

Public Type POINTAPI
 X As Long
 Y As Long
End Type

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Const SE_ERR_NOASSOC = 31
Const sOperation As String = "open"     ' Constants for shell operations
Const sRun As String = "RUNDLL32.EXE"
Const sParameters As String = "shell32.dll,OpenAs_RunDLL "
    Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter_ As Long, ByVal X As Long, ByVal y_ As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags_ As Long) As Long
    Const conHwndTopmost = -1
    Const conHwndNoTopmost = -2
    Const conSwpNoActivate = &H10
    Const conSwpShowWindow = &H40

Public Function shelldoc(sfile As String)
    Dim sPath As String, RetVal As Long, _
    lRet As Long
    lRet = ShellExecute(GetDesktopWindow(), sOperation, sfile, _
                        vbNullString, vbNullString, SW_SHOWNORMAL)
    If lRet = SE_ERR_NOASSOC Then ' No association exists
        'Create a buffer
        sPath = Space(255)
        'Get the system directory
        RetVal = GetSystemDirectory(sPath, 255)
        'Remove all unnecessary chr$(0)'s
        'and move on the stack
        sPath = Left$(sPath, RetVal)
    
        lRet = ShellExecute(GetDesktopWindow(), "open", sRun, _
                            sParaters + sfile, sPath, SW_SHOWNORMAL)
    End If
End Function
Public Function AlwaysOnTop(ByVal H, FrmX As Long, FrmY As Long, Hght As Long, Wdth As Long, YesAOT As Boolean)

    If YesAOT = True Then
        SetWindowPos H, conHwndTopmost, FrmX, FrmY, Wdth, Hght, conSwpNoActivate
        ElseIf YesAOT = False Then
        SetWindowPos H, conHwndNoTopmost, FrmX, FrmY, Wdth, Hght, conSwpShowWindow
    End If

End Function

Public Sub ShapeResize(FormName)

'resize formname.shape1
FormName.Shape1.Width = FormName.ScaleWidth - 10
FormName.Shape1.Height = FormName.ScaleHeight - 10

'resize formname.shape2
FormName.Shape2.Width = FormName.ScaleWidth - 85
FormName.Shape2.Height = FormName.ScaleHeight - 85

'resize formname.shape3
FormName.Shape3.Width = FormName.ScaleWidth - 160
FormName.Shape3.Height = FormName.ScaleHeight - 160

'resize formname.shape4
FormName.Shape4.Width = FormName.ScaleWidth - 235
FormName.Shape4.Height = FormName.ScaleHeight - 235

'resize formname.shape5
FormName.Shape5.Width = FormName.ScaleWidth - 310
FormName.Shape5.Height = FormName.ScaleHeight - 310

'resize lblTop
FormName.lblTop.Width = FormName.ScaleWidth - 310 - 70
End Sub
