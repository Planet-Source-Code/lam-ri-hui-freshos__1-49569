VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3255
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4470
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin FreshOS.PB_Blue PB_Blue1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
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
      Left            =   840
      TabIndex        =   1
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   2025
      Left            =   240
      Picture         =   "frmSplash.frx":0000
      Top             =   240
      Width           =   3855
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
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call AlwaysOnTop(Me.hWnd, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Height / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, True)
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

'Show the form
Me.Visible = True

'declare variables
Dim Work(2000) As Byte
Dim Count As Integer
Dim X As Integer

'set the blue gradient progress bar's max value to 100
PB_Blue1.Max = 100

'for loop
For Count = 0 To 100

For X = LBound(Work) To UBound(Work)
DoEvents
Next X

PB_Blue1.Value = Count / 100 * 100
Next Count

Main.Show
Unload Me

End Sub

Private Sub Timer1_Timer()

End Sub

