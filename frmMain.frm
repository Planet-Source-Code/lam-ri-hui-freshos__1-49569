VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "FreshOS"
   ClientHeight    =   10050
   ClientLeft      =   2700
   ClientTop       =   645
   ClientWidth     =   10290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   10050
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Transparent"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   840
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4920
      Top             =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © Lam Ri Hui 2003"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   12
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label lblDraw 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FreshOS Draw"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   11
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Image imgDraw 
      Height          =   600
      Left            =   840
      Picture         =   "frmMain.frx":11C2
      Top             =   6360
      Width           =   555
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About FreshOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4320
      TabIndex        =   10
      Top             =   240
      Width           =   1425
   End
   Begin VB.Label lblRun 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FreshOS Run"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   9
      Top             =   8880
      Width           =   1575
   End
   Begin VB.Image imgRun 
      Height          =   480
      Left            =   840
      Picture         =   "frmMain.frx":15DD
      ToolTipText     =   "Run - Select a file or appilcation and run it."
      Top             =   8280
      Width           =   480
   End
   Begin VB.Label lblGames 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FreshOS Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   450
      TabIndex        =   8
      Top             =   6000
      Width           =   1395
   End
   Begin VB.Image imgGames 
      Height          =   480
      Left            =   840
      Picture         =   "frmMain.frx":1EA7
      ToolTipText     =   "FreshOS Games - Play games provided free by FreshOS!"
      Top             =   5400
      Width           =   480
   End
   Begin VB.Label lblFreshPad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FreshOS Pad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   7
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Image imgFreshPad 
      Height          =   480
      Left            =   840
      Picture         =   "frmMain.frx":2771
      ToolTipText     =   "FreshOS Pad - Edit your files using this tool!"
      Top             =   4440
      Width           =   480
   End
   Begin VB.Label lblShutdown 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shut Down"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6000
      TabIndex        =   6
      Top             =   240
      Width           =   1035
   End
   Begin VB.Label lblFreshFind 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FreshOS Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   5
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label lblInternetBrowser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FreshOS Internet Browser"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Image imgFind 
      Height          =   480
      Left            =   840
      Picture         =   "frmMain.frx":2BB3
      ToolTipText     =   "FreshOS Find - Search for a file on your harddisk using this tool!"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image imgInternetBrowser 
      Height          =   480
      Left            =   840
      Picture         =   "frmMain.frx":347D
      ToolTipText     =   "FreshOS Internet Browser - Use this tool to access the world wide web!"
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label lblMediaPlayer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FreshOS Media Player"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Image imgMediaPlayer 
      Height          =   480
      Left            =   840
      Picture         =   "frmMain.frx":3D47
      ToolTipText     =   "FreshOS Media Player - Use this to play audio files or movies!"
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label lblDriveExplorer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FreshOS Drive Explorer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Image imgDriveExplorer 
      Height          =   480
      Left            =   840
      Picture         =   "frmMain.frx":4611
      ToolTipText     =   "FreshOS Drive Explorer - Explore your computer using this powerfull tool!"
      Top             =   720
      Width           =   480
   End
   Begin VB.Label lblTimeAndDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label lblTop 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   195
      TabIndex        =   0
      Top             =   195
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   1905
      Left            =   5400
      Picture         =   "frmMain.frx":4EDB
      Top             =   720
      Width           =   3150
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      Height          =   525
      Left            =   165
      Top             =   165
      Width           =   525
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   690
      Left            =   90
      Top             =   90
      Width           =   690
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Height          =   780
      Left            =   45
      Top             =   45
      Width           =   780
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      Height          =   840
      Left            =   0
      Top             =   15
      Width           =   855
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mouse As POINTAPI
Dim Dragging As Boolean

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Public Function Transparent(ByVal hwnd As Long, ByVal Perc As Integer) As Long
   Dim Msg As Long
   On Error Resume Next
    
   Perc = ((100 - Perc) / 100) * 255
   If Perc < 0 Or Perc > 255 Then
     Transparent = 1
   Else
     Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
     Msg = Msg Or WS_EX_LAYERED
     SetWindowLong hwnd, GWL_EXSTYLE, Msg
     SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
     Transparent = 0
   End If
   If Err Then
     Transparent = 2
   End If
End Function


Public Function Opaque(ByVal hwnd As Long) As Long
   Dim Msg As Long
   On Error Resume Next
   Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
   Msg = Msg And Not WS_EX_LAYERED
   SetWindowLong hwnd, GWL_EXSTYLE, Msg
   SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
   Opaque = 0
   If Err Then
   Opaque = 2
   End If
End Function

Private Sub Check1_Click()
'check1 is checked,
If Check1.Value = 1 Then
'make form transparent
Transparent Me.hwnd, 40
'otherwise
Else
'make form opaque
Opaque Me.hwnd
'end if statement
End If
End Sub

Private Sub Form_Resize()
ShapeResize Me

'move image1
Image1.Left = Me.ScaleWidth - 400 - Image1.Width

'reposition label1
Label1.Left = Me.ScaleWidth - 700 - Label1.Width

'reposition lblshutdown
lblShutdown.Left = Me.ScaleWidth - 310 - lblShutdown.Width

'reposition check1
Check1.Top = Me.ScaleHeight - 310 - Check1.Height
Check1.Left = Me.ScaleWidth - 310 - Check1.Width

'reposition lblabout
lblAbout.Left = Me.ScaleWidth - 310 - lblShutdown.Width - 400 - lblAbout.Width
End Sub

Private Sub Image1_Click()
AboutFreshOS.Show
End Sub

Private Sub imgDraw_Click()
RemoveBold
lblDraw.FontBold = True
End Sub

Private Sub imgDraw_DblClick()
FreshOSDraw.Show
End Sub

Private Sub imgDriveExplorer_Click()
RemoveBold
lblDriveExplorer.FontBold = True
End Sub

Private Sub imgDriveExplorer_DblClick()
FreshOSDriveExplorer.Show
End Sub

Private Sub imgFind_Click()
RemoveBold
lblFreshFind.FontBold = True
End Sub

Private Sub imgFind_DblClick()
FreshOSFind.Show
End Sub

Private Sub imgFreshPad_Click()
RemoveBold
lblFreshPad.FontBold = True
End Sub

Private Sub imgFreshPad_DblClick()
FreshOSPad.Show
End Sub

Private Sub imgGames_Click()
RemoveBold
lblGames.FontBold = True
End Sub

Private Sub imgGames_DblClick()
FreshOSGames.Show
End Sub

Private Sub imgInternetBrowser_Click()
RemoveBold
lblInternetBrowser.FontBold = True
End Sub

Private Sub imgInternetBrowser_DblClick()
FreshOSInternetBrowser.Show
End Sub

Private Sub imgMediaPlayer_Click()
RemoveBold
lblMediaPlayer.FontBold = True
End Sub

Private Sub imgMediaPlayer_DblClick()
FreshOSMediaPlayer.Show
End Sub

Private Sub imgRun_Click()
RemoveBold
lblRun.FontBold = True
End Sub

Private Sub imgRun_DblClick()
FreshOSRun.Show
End Sub

Private Sub lblAbout_Click()
AboutFreshOS.Show
End Sub

Private Sub lblAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAbout.ForeColor = vbRed
End Sub

Private Sub lblAbout_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAbout.ForeColor = vbBlack
End Sub

Private Sub lblShutdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
lblShutdown.ForeColor = vbRed
End If
End Sub

Private Sub lblShutdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
lblShutdown.ForeColor = vbBlack
End If
ShutDown.Show 1, Main
End Sub

Private Sub Timer1_Timer()
lblTimeAndDate.Caption = Format(Date, "dd/mmmm/yyyy") & " " & Format(Time, "h:mm:ss")
End Sub

Private Sub RemoveBold()
Dim lbl
For Each lbl In Me
If TypeOf lbl Is Label Then
lbl.FontBold = False
End If
Next
lblShutdown.FontBold = True
lblAbout.FontBold = True
End Sub

Private Sub Timer2_Timer()

End Sub
