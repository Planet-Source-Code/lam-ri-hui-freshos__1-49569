VERSION 5.00
Begin VB.Form ShutDown 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FreshOS.FreshButton FreshButton1 
      Height          =   735
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      Caption         =   "Shut Down"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shut Down FreshOS?"
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
      Width           =   1890
   End
   Begin VB.Label lblTop 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   3375
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      Height          =   645
      Left            =   150
      Top             =   150
      Width           =   1005
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      Height          =   735
      Left            =   105
      Top             =   105
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   810
      Left            =   75
      Top             =   75
      Width           =   1170
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Height          =   900
      Left            =   30
      Top             =   30
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "ShutDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
ShapeResize Me
'resize shape1
Shape1.Width = Me.ScaleWidth - 10
Shape1.Height = Me.ScaleHeight - 10

'resize shape2
Shape2.Width = Me.ScaleWidth - 85
Shape2.Height = Me.ScaleHeight - 85

'resize shape3
Shape3.Width = Me.ScaleWidth - 160
Shape3.Height = Me.ScaleHeight - 160

'resize shape4
Shape4.Width = Me.ScaleWidth - 235
Shape4.Height = Me.ScaleHeight - 235

'resize shape5
Shape5.Width = Me.ScaleWidth - 310
Shape5.Height = Me.ScaleHeight - 310

'resize lblTop
lblTop.Width = Me.ScaleWidth - 310 - 70

'reposition lblExit
lblExit.Left = Me.ScaleWidth - 360
End Sub
Private Sub lblExit_Click()
'unload shutdown
Unload Me
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'set the forecolor to red instead of using RGB(255,0,0)
lblExit.ForeColor = vbRed
End Sub

Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'set the forecolor of the label to black again
'when you release the mouse button
lblExit.ForeColor = vbBlack
End Sub
Private Sub FreshButton1_Click()
'Unload AboutFreshOS
'Unload FreshOSDraw
'Unload FreshOSDriveExplorer
'Unload FreshOSFind
'Unload FreshOSGames
'Unload FreshOSInternetBrowser
'Unload FreshOSMediaPlayer
'Unload FreshOSPad
'Unload FreshOSRun
'Unload Splash
'Unload Main
'Unload ShutDown

'Another way of unloading all forms instead
'of using the above code

'Declare variable 'form'
Dim Form

'count how many object in this project
For Each Form In Forms
'check if the object is form or not
If TypeOf Form Is Form Then
'if it is, then unload it
Unload Form
'end if statement
End If
'loop
Next
End Sub



Private Sub lblTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
'if the button pressed is left button,
If Button = 1 Then
Call ReleaseCapture
'move the window to the position
lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub
