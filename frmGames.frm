VERSION 5.00
Begin VB.Form FreshOSGames 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   ScaleHeight     =   3525
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin FreshOS.FreshButton FreshButton1 
      Height          =   615
      Left            =   4560
      TabIndex        =   17
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      Caption         =   "New"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   1920
      TabIndex        =   6
      Top             =   600
      Width           =   2535
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   120
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   2
         Left            =   1800
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   3
         Left            =   120
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   13
         Top             =   1080
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   4
         Left            =   960
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   5
         Left            =   1800
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   11
         Top             =   1080
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   6
         Left            =   120
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   10
         Top             =   1920
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   7
         Left            =   960
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   9
         Top             =   1920
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   1
         Left            =   960
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   8
         Left            =   1800
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   7
         Top             =   1920
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2400
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   2400
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line3 
         X1              =   840
         X2              =   840
         Y1              =   240
         Y2              =   2520
      End
      Begin VB.Line Line4 
         X1              =   1680
         X2              =   1680
         Y1              =   240
         Y2              =   2520
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Opponent"
      Height          =   1215
      Left            =   4560
      TabIndex        =   4
      Top             =   600
      Width           =   1335
      Begin VB.OptionButton Gtype 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CPU"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Gtype 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CPU First"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.ListBox ScoreBoard 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      ItemData        =   "frmGames.frx":0000
      Left            =   240
      List            =   "frmGames.frx":000A
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin FreshOS.FreshButton FreshButton2 
      Height          =   615
      Left            =   4560
      TabIndex        =   18
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      Caption         =   "Clear History"
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
      TabIndex        =   2
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FreshOS Game version 1.0"
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
      Width           =   2355
   End
   Begin VB.Label lblTop 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   5775
   End
End
Attribute VB_Name = "FreshOSGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CPUplay As Boolean
Private whowin
Private playstr As String
Private NumBoxFilled
Private NGP
Private games_played As Integer
Private MovHis() As String
Private loseweight(8) As Integer
'Numbering of the Picture Box
' 0 1 2
' 3 4 5
' 6 7 8
'
'Draws X on the desinated Picture Box, Also Sets the tag property to 'X'
Private Sub drawX(Index As Integer)

On Error Resume Next
    Box(Index).Line (0, 1)-(38, 39)
    Box(Index).Line (0, 0)-(39, 39)
    Box(Index).Line (1, 0)-(39, 38)
    Box(Index).Line (38, 1)-(1, 38)
    Box(Index).Line (38, 0)-(0, 38)
    Box(Index).Line (37, 0)-(0, 37)
    Box(Index).Tag = "X"

End Sub
'Draws O on the desinated Picture Box, Also sets the tag property to 'O'
Private Sub drawO(Index As Integer)

On Error Resume Next
    Box(Index).FillStyle = 0
    Box(Index).FillColor = vbBlack
    Box(Index).Circle (20, 20), 19
    Box(Index).FillColor = vbWhite
    Box(Index).Circle (20, 20), 17
    Box(Index).Tag = "O"

End Sub

'Function that returns true if a winning combination is found
Private Function win() As Boolean

On Error Resume Next
    If rowsame(0, 1, 2) Or rowsame(3, 4, 5) Or rowsame(6, 7, 8) Or rowsame(0, 3, 6) Or rowsame(1, 4, 7) Or rowsame(2, 5, 8) Or rowsame(0, 4, 8) Or rowsame(2, 4, 6) Then win = True

End Function
'Sub Function used by Win() to check wheter the 3 PictureBox has the same Symbol
Private Function rowsame(i1, i2, i3 As Integer) As Boolean

On Error Resume Next
    If (Box(i1).Tag = Box(i2).Tag) And (Box(i2).Tag = Box(i3).Tag) And (Box(i1).Tag = "X" Or Box(i1).Tag = "O") Then
        rowsame = True
        whowin = Box(i1).Tag
        Else
        rowsame = False
    End If

End Function


'Executes when the Picture Boxes are clicked
Private Sub Box_Click(Index As Integer)

On Error Resume Next
Dim mvpos
Dim rp As Integer
Dim mvstr As String
    Gtype(1).Enabled = False

    Gtype(2).Enabled = False
    playstr = playstr + CStr(Index)
    If Box(Index).Tag = "" Then
        drawX (Index)
        CPUplay = True
        NumBoxFilled = NumBoxFilled + 1
    End If
'If either side wins or draw, exit subroutine
    If checkforwin Then Exit Sub
'CPU Play
    If CPUplay Then
        mvstr = FindMoves
        rp = Int(Rnd * Len(mvstr)) + 1
        mvpos = CDec(Mid(mvstr, rp, 1))
        Debug.Print "Available Moves:"; mvstr; " Choose:"; mvpos
        drawO (mvpos)
        playstr = playstr + CStr(mvpos)
        CPUplay = False
        NumBoxFilled = NumBoxFilled + 1
    End If
    If checkforwin Then Exit Sub

End Sub

Private Function checkforwin() As Boolean

On Error Resume Next
    checkforwin = False
    If win Then
        MsgBox whowin + " Won", , "Game Over!"
        games_played = games_played + 1
        ScoreBoard.AddItem ("Game " + CStr(games_played) + " : " + whowin + " Win")
        playstr = playstr + whowin
        Debug.Print playstr
        If Gtype(1).Value Then
            NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = playstr
            NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = transpose(playstr)
            NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = rotate(playstr)
            NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = transpose(rotate(playstr))
            NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = rotate(rotate(playstr))
            NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = transpose(rotate(rotate(playstr)))
            NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = rotate(rotate(rotate(playstr)))
            NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = transpose(rotate(rotate(rotate(playstr))))
        End If
        Gtype(1).Enabled = True
        Gtype(2).Enabled = True
        Gtype(1).Value = True
        Call FreshButton1_Click
        checkforwin = True
        ElseIf NumBoxFilled = 9 Then
        MsgBox "Draw", , "Game Over!"
        games_played = games_played + 1
        ScoreBoard.AddItem ("Game " + CStr(games_played) + " : " + "Draw")
        playstr = playstr + "D"
        Debug.Print playstr
        Call FreshButton1_Click
        checkforwin = True
        Gtype(1).Enabled = True
        Gtype(2).Enabled = True
        Gtype(1).Value = True
    End If

End Function

Private Function transpose(tstr As String) As String

On Error Resume Next
Dim temp
Dim c
Dim c1 As String
Dim i

    For i = 1 To Len(tstr)
        c = Mid(tstr, i, 1)
        Select Case c
            Case Is = "1": c1 = "3"
            Case Is = "3": c1 = "1"
            Case Is = "2": c1 = "6"
            Case Is = "6": c1 = "2"
            Case Is = "5": c1 = "7"
            Case Is = "7": c1 = "5"
            Case Else: c1 = c
        End Select
        temp = temp + c1
    Next
    transpose = temp

End Function
'Matrix Manipulation Function
'Rotates the Matrix by 90 degrees, to rotate 180, just use the function twice
'for 270, three times
'
' 012    258
' 345 => 147
' 678    036
Private Function rotate(tstr As String) As String

On Error Resume Next
Dim temp
Dim c
Dim c1 As String
Dim i

    For i = 1 To Len(tstr)
        c = Mid(tstr, i, 1)
        Select Case c
            Case Is = "0": c1 = "2"
            Case Is = "1": c1 = "5"
            Case Is = "2": c1 = "8"
            Case Is = "3": c1 = "1"
            Case Is = "4": c1 = "4"
            Case Is = "5": c1 = "7"
            Case Is = "6": c1 = "0"
            Case Is = "7": c1 = "3"
            Case Is = "8": c1 = "6"
            Case Else: c1 = c
        End Select
        temp = temp + c1
    Next
    rotate = temp

End Function
'Finds the Appropriate move for the CPU, the function returns a string containing all
'possible moves
Private Function FindMoves() As String

On Error Resume Next
Dim i
Dim j
Dim l
Dim sml As Integer
Dim pst
Dim mvc As String

    For i = 0 To 8
        loseweight(i) = 0
    Next
    For j = 0 To 8
        pst = playstr + CStr(j)
        l = Len(pst)
        For i = 1 To NGP
            If l < Len(MovHis(i)) Then
                If pst = Left(MovHis(i), l) Then
                    If Right(MovHis(i), 1) = "X" Then loseweight(j) = loseweight(j) + 1
                End If
            End If
        Next i
    Next j
'Different weight calculation when the CPU moves first
    If Gtype(1).Value Then  'When CPU moves Second
        sml = 32767
        For i = 0 To 8
            If sml > loseweight(i) And Box(i).Tag = "" Then sml = loseweight(i)
        Next
        Else    'When CPU moves First
        sml = 0
        For i = 0 To 8
            If sml < loseweight(i) And Box(i).Tag = "" Then sml = loseweight(i)
        Next
    End If
    For i = 0 To 8
        If loseweight(i) = sml And Box(i).Tag = "" Then mvc = mvc + CStr(i)
    Next
'Checks wheter if the Human Player has a winning move, if there is, then block it
    For i = 0 To 8
        If Box(i).Tag = "" Then
            Box(i).Tag = "X"
            If win Then mvc = CStr(i)
            Box(i).Tag = ""
        End If
    Next
'Checks wheter the CPU has a winning move, if it does, then choose that move.
    For i = 0 To 8
        If Box(i).Tag = "" Then
            Box(i).Tag = "O"
            If win Then mvc = CStr(i)
            Box(i).Tag = ""
        End If
    Next
    FindMoves = mvc

End Function
Private Sub Form_Load()
ShapeResize Me

'reposition lblExit
lblExit.Left = Me.ScaleWidth - 360

On Error Resume Next
Dim i As Integer
    Randomize Timer

    On Error GoTo errHand
    Open App.path + "\MoveHis.txt" For Input As #1
        Input #1, NGP
        ReDim MovHis(NGP)
        For i = 1 To NGP
            Input #1, MovHis(i)
        Next
errHand:
    Close #1
End Sub

Private Sub freshbutton1_click_Click()

On Error Resume Next
    For i = 0 To 8
        Box(i).Cls
        Box(i).Tag = ""
    Next
    CPUplay = False
    whowin = ""
    playstr = ""
    NumBoxFilled = 0
    Gtype(1).Enabled = True
    Gtype(2).Enabled = True
    Gtype(1).Value = True

End Sub

Private Sub FreshButton1_Click()

On Error Resume Next
    For i = 0 To 8
        Box(i).Cls
        Box(i).Tag = ""
    Next
    CPUplay = False
    whowin = ""
    playstr = ""
    NumBoxFilled = 0
    Gtype(1).Enabled = True
    Gtype(2).Enabled = True
    Gtype(1).Value = True
End Sub

Private Sub FreshButton2_Click()

On Error Resume Next
    games_played = 0
    ScoreBoard.Clear

End Sub

Private Sub lblExit_Click()

On Error Resume Next
Dim CompactHis() As String
Dim CNum
Dim i

    For i = 1 To NGP
        If Right(MovHis(i), 1) = "X" Then   'Saves only losing moves
            CNum = CNum + 1
            ReDim Preserve CompactHis(CNum)
            CompactHis(CNum) = MovHis(i)
        End If
    Next
    If NGP = 0 Then Unload Me

'Open File for output
    Open App.path + "\MoveHis.txt" For Output As #1
        Print #1, CNum
        For i = 1 To CNum
            Print #1, CompactHis(i)
        Next
    Close #1

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

'Executes when the Options menu are Clicked
'Chooses between CPU play first or Human First
Private Sub Gtype_Click(Index As Integer)

On Error Resume Next
    If Gtype(2).Value Then
        mv = Int(Rnd * 9)
        drawO (mv)
        NumBoxFilled = NumBoxFilled + 1
        playstr = CStr(mv)
        Gtype(1).Enabled = False
        Gtype(2).Enabled = False
    End If

End Sub

