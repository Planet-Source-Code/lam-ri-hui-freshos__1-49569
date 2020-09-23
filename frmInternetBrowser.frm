VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form FreshOSInternetBrowser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13350
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   9735
   ScaleWidth      =   13350
   StartUpPosition =   2  'CenterScreen
   Begin FreshOS.FreshButton FreshButton2 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Back"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FreshOS.FreshButton FreshButton1 
      Default         =   -1  'True
      Height          =   375
      Left            =   12120
      TabIndex        =   5
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Go"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   4
      Text            =   "http://www.yahoo.com"
      Top             =   1080
      Width           =   11775
   End
   Begin SHDocVwCtl.WebBrowser wb1 
      Height          =   7935
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   12855
      ExtentX         =   22675
      ExtentY         =   13996
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin FreshOS.FreshButton FreshButton3 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Forward"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FreshOS.FreshButton FreshButton4 
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Stop"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FreshOS.FreshButton FreshButton5 
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Refresh"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FreshOS.FreshButton FreshButton6 
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Home"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FreshOS Internet Browser version 1.0"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3240
   End
   Begin VB.Label lblTop 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Height          =   900
      Left            =   30
      Top             =   30
      Width           =   1260
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   810
      Left            =   75
      Top             =   75
      Width           =   1170
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      Height          =   735
      Left            =   105
      Top             =   105
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      Height          =   645
      Left            =   150
      Top             =   150
      Width           =   1005
   End
End
Attribute VB_Name = "FreshOSInternetBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
wb1.Navigate Combo1.List(Combo1.ListIndex)
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
wb1.Navigate Combo1.Text
Combo1.AddItem Combo1.Text
End If
End Sub

Private Sub Form_Load()
ShapeResize Me

'reposition lblExit
lblExit.Left = Me.ScaleWidth - 360

wb1.Navigate Combo1.Text
Combo1.AddItem Combo1.Text
End Sub

Private Sub FreshButton1_Click()
wb1.Navigate Combo1.Text
Combo1.AddItem Combo1.Text
End Sub

Private Sub FreshButton2_Click()
wb1.GoBack
End Sub

Private Sub FreshButton3_Click()
wb1.GoForward
End Sub

Private Sub FreshButton4_Click()
wb1.Stop
End Sub

Private Sub FreshButton5_Click()
wb1.Refresh
End Sub

Private Sub FreshButton6_Click()
wb1.GoHome
End Sub

Private Sub lblExit_Click()
Unload Me
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub
Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = vbRed
End Sub

Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = vbBlack
End Sub
Private Sub lblTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub
