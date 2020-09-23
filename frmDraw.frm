VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FreshOSDraw 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6480
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FreshOS.FreshButton FreshButton1 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Save"
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
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   240
      ScaleHeight     =   6465
      ScaleWidth      =   11265
      TabIndex        =   3
      Top             =   1080
      Width           =   11295
   End
   Begin FreshOS.FreshButton FreshButton3 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Clear"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FreshOS Draw version 1.0"
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
      TabIndex        =   1
      Top             =   240
      Width           =   2265
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
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   135
   End
   Begin VB.Label lblTop 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   180
      TabIndex        =   2
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
Attribute VB_Name = "FreshOSDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ShapeResize Me

'reposition lblExit
lblExit.Left = Me.ScaleWidth - 360
End Sub

Private Sub FreshButton1_Click()
On Error Resume Next
cd1.CancelError = True
cd1.Filter = "Bitmap Files(*.bmp)|*.bmp"
cd1.Action = 2
SavePicture pic.Image, cd1.FileName
End Sub

Private Sub FreshButton3_Click()
pic.Cls
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub lblExit_Click()
Unload Me
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

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pic.DrawWidth = 10
    If Button = 1 Then pic.ForeColor = vbBlack: pic.PSet (X, Y)
    If Button = 2 Then pic.ForeColor = vbWhite: pic.PSet (X, Y)
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then pic.ForeColor = vbBlack: pic.PSet (X, Y)
    If Button = 2 Then pic.ForeColor = vbWhite: pic.PSet (X, Y)
End Sub
