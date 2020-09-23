VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form FreshOSPad 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RTBmain 
      Height          =   6495
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11456
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmTextPad.frx":0000
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   6360
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6960
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0082
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0194
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":02A6
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":03B8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":04CA
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":05DC
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":06EE
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0800
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0912
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0A24
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0B36
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0C48
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0D5A
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextPad.frx":0F80
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Height          =   360
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
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
      Left            =   5760
      TabIndex        =   2
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FreshOS Pad version 1.0"
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
      Width           =   2175
   End
   Begin VB.Label lblTop 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5775
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
Attribute VB_Name = "FreshOSPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
ShapeResize Me
'reposition lblExit
lblExit.Left = Me.ScaleWidth - 360
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

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Dim sfile As String
    Select Case Button.Key
        Case "New"
            Dim SaveCheck As Integer
            If RTBmain.Text <> "" Then SaveCheck = MsgBox("Do you want to quit without saving?", vbYesNo, "Nompadd")
                If SaveCheck = vbYes Then
                    RTBmain.Text = ""
                Else
                    Exit Sub
                End If
        Case "Font"
            With dlgCommonDialog
                .Flags = 3
                .DialogTitle = "Fonts"
                .CancelError = True
                .ShowFont
            End With
            RTBmain.SelFontName = dlgCommonDialog.FontName
            RTBmain.SelFontSize = dlgCommonDialog.FontSize
            RTBmain.SelBold = dlgCommonDialog.FontBold
            RTBmain.SelItalic = dlgCommonDialog.FontItalic
            RTBmain.SelUnderline = dlgCommonDialog.FontUnderline
        Case "Open"
            With dlgCommonDialog
                .DialogTitle = "Open"
                .CancelError = False
                .Filter = "Text File|*.txt|DAT file|*.dat|All Files|*.*"
                .ShowOpen
                If Len(.FileName) = 0 Then
                    Exit Sub
                End If
            sfile = .FileName
            End With
            RTBmain.LoadFile sfile
        Case "Save"
            With dlgCommonDialog
                .DialogTitle = "Save"
                .CancelError = False
                .Filter = "Text File|*.txt|DAT file|*.dat"
                .ShowSave
                If Len(.FileName) = 0 Then
                    Exit Sub
                End If
                sfile = .FileName
            End With
                RTBmain.SaveFile sfile
                RTBmain.SaveFile sfile
        Case "Print"
            With dlgCommonDialog
                .DialogTitle = "Print"
                .CancelError = True
                .Flags = cdlPDReturnDC + cdlPDNoPageNums
                .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            RTBmain.SelPrint .hDC
        End If
    End With
        Case "Cut"
            Clipboard.SetText RTBmain.SelRTF
            RTBmain.SelText = vbNullString
        Case "Copy"
           Clipboard.SetText RTBmain.SelRTF
        Case "Paste"
            RTBmain.SelRTF = Clipboard.GetText
        Case "Bold"
            RTBmain.SelBold = Not RTBmain.SelBold
            Button.Value = IIf(RTBmain.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            RTBmain.SelItalic = Not RTBmain.SelItalic
            Button.Value = IIf(RTBmain.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            RTBmain.SelUnderline = Not RTBmain.SelUnderline
            Button.Value = IIf(RTBmain.SelUnderline, tbrPressed, tbrUnpressed)
        Case "Center"
            RTBmain.SelAlignment = rtfCenter
            Button.Value = IIf(RTBmain.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
        Case "Align Left"
            RTBmain.SelAlignment = rtfLeft
        Case "Align Right"
            RTBmain.SelAlignment = rtfRight
    End Select
End Sub
