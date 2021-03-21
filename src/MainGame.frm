VERSION 5.00
Begin VB.Form MainGame 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnmainmenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton btnrestart 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton btnclear 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton btn5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton btn6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton btn7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton btn3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton btn8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton btn4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton btn0 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton btn1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton btn2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   8220
      X2              =   9420
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   8220
      X2              =   9420
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Label lblplayeroscore 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Score : 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8220
      TabIndex        =   12
      Top             =   4920
      Width           =   930
   End
   Begin VB.Label lblplayerxscore 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Score : 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8220
      TabIndex        =   11
      Top             =   3660
      Width           =   930
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Player O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8220
      TabIndex        =   10
      Top             =   4440
      Width           =   1020
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Player X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   8220
      TabIndex        =   9
      Top             =   3180
      Width           =   1005
   End
End
Attribute VB_Name = "MainGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Autuhor & Copyright (c)   : Kevin C. Magnifico
Rem Computer language         : Visual Basic 6
Rem Date Created              : 03/21/2021
Rem Game type                 : Tic tac toe
Rem License                   : Public Domain

Private is_x As Boolean
Private can_click_btn(8) As Boolean
Private pattern_lock(7) As Boolean
Private player_x_score As Integer
Private player_o_score As Integer

Private Sub InitializeAllData()
    is_x = False
    For i = 0 To UBound(can_click_btn)
        can_click_btn(i) = True
    Next
    For i = 0 To UBound(pattern_lock)
        pattern_lock(i) = False
    Next
    Me.btn0.Caption = ""
    Me.btn1.Caption = ""
    Me.btn2.Caption = ""
    Me.btn3.Caption = ""
    Me.btn4.Caption = ""
    Me.btn5.Caption = ""
    Me.btn6.Caption = ""
    Me.btn7.Caption = ""
    Me.btn8.Caption = ""
End Sub

Rem 0,1,2 = pattern_lock(0)
Rem 0,3,6 = pattern_lock(1)
Rem 0,4,8 = pattern_lock(2)
Rem 1,4,7 = pattern_lock(3)
Rem 2,5,8 = pattern_lock(4)
Rem 2,4,6 = pattern_lock(5)
Rem 3,4,5 = pattern_lock(6)
Rem 6,7,8 = pattern_lock(7)

Private Sub AddPlayerXScore()
    player_x_score = (player_x_score + 1)
    Me.lblplayerxscore.Caption = "Score : " & Str(player_x_score)
End Sub

Private Sub AddPlayerOScore()
    player_o_score = (player_o_score + 1)
    Me.lblplayeroscore.Caption = "Score : " & Str(player_o_score)
End Sub

Private Sub Checking()
    If is_x = True Then
        Rem Checking for Player X
        If Me.btn0.Caption = "X" And Me.btn1.Caption = "X" And Me.btn2.Caption = "X" And pattern_lock(0) = False Then
            MsgBox "Player X Tic Tac - 0,1,2", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(0) = True
            Call AddPlayerXScore
        ElseIf Me.btn0.Caption = "X" And Me.btn3.Caption = "X" And Me.btn6.Caption = "X" And pattern_lock(1) = False Then
            MsgBox "Player X Tic Tac - 0,3,6", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(1) = True
            Call AddPlayerXScore
        ElseIf Me.btn0.Caption = "X" And Me.btn4.Caption = "X" And Me.btn8.Caption = "X" And pattern_lock(2) = False Then
            MsgBox "Player X Tic Tac - 0,4,8", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(2) = True
            Call AddPlayerXScore
        ElseIf Me.btn1.Caption = "X" And Me.btn4.Caption = "X" And Me.btn7.Caption = "X" And pattern_lock(3) = False Then
            MsgBox "Player X Tic Tac - 1,4,7", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(3) = True
            Call AddPlayerXScore
        ElseIf Me.btn2.Caption = "X" And Me.btn5.Caption = "X" And Me.btn8.Caption = "X" And pattern_lock(4) = False Then
            MsgBox "Player X Tic Tac - 2,5,8", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(4) = True
            Call AddPlayerXScore
        ElseIf Me.btn2.Caption = "X" And Me.btn4.Caption = "X" And Me.btn6.Caption = "X" And pattern_lock(5) = False Then
            MsgBox "Player X Tic Tac - 2,4,6", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(5) = True
            Call AddPlayerXScore
        ElseIf Me.btn3.Caption = "X" And Me.btn4.Caption = "X" And Me.btn5.Caption = "X" And pattern_lock(6) = False Then
            MsgBox "Player X Tic Tac - 3,4,5", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(6) = True
            Call AddPlayerXScore
        ElseIf Me.btn6.Caption = "X" And Me.btn7.Caption = "X" And Me.btn8.Caption = "X" And pattern_lock(7) = False Then
            MsgBox "Player X Tic Tac - 6,7,8", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(7) = True
            Call AddPlayerXScore
        End If
    ElseIf is_x = False Then
        Rem Checking for Player O
        If Me.btn0.Caption = "O" And Me.btn1.Caption = "O" And Me.btn2.Caption = "O" And pattern_lock(0) = False Then
            MsgBox "Player O Tic Tac - 0,1,2", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(0) = True
            Call AddPlayerOScore
        ElseIf Me.btn0.Caption = "O" And Me.btn3.Caption = "O" And Me.btn6.Caption = "O" And pattern_lock(1) = False Then
            MsgBox "Player O Tic Tac - 0,3,6", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(1) = True
            Call AddPlayerOScore
        ElseIf Me.btn0.Caption = "O" And Me.btn4.Caption = "O" And Me.btn8.Caption = "O" And pattern_lock(2) = False Then
            MsgBox "Player O Tic Tac - 0,4,8", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(2) = True
            Call AddPlayerOScore
        ElseIf Me.btn1.Caption = "O" And Me.btn4.Caption = "O" And Me.btn7.Caption = "O" And pattern_lock(3) = False Then
            MsgBox "Player O Tic Tac - 1,4,7", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(3) = True
            Call AddPlayerOScore
        ElseIf Me.btn2.Caption = "O" And Me.btn5.Caption = "O" And Me.btn8.Caption = "O" And pattern_lock(4) = False Then
            MsgBox "Player O Tic Tac - 2,5,8", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(4) = True
            Call AddPlayerOScore
        ElseIf Me.btn2.Caption = "O" And Me.btn4.Caption = "O" And Me.btn6.Caption = "O" And pattern_lock(5) = False Then
            MsgBox "Player O Tic Tac - 2,4,6", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(5) = True
            Call AddPlayerOScore
        ElseIf Me.btn3.Caption = "O" And Me.btn4.Caption = "O" And Me.btn5.Caption = "O" And pattern_lock(6) = False Then
            MsgBox "Player O Tic Tac - 3,4,5", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(6) = True
            Call AddPlayerOScore
        ElseIf Me.btn6.Caption = "O" And Me.btn7.Caption = "O" And Me.btn8.Caption = "O" And pattern_lock(7) = False Then
            MsgBox "Player O Tic Tac - 6,7,8", vbInformation + vbOKOnly, "Tic tac toe"
            pattern_lock(7) = True
            Call AddPlayerOScore
        End If
    End If
End Sub

Private Sub btn0_Click()
    If can_click_btn(0) = True Then
        If is_x = False Then
            Me.btn0.Caption = "X"
            is_x = True
        Else
            Me.btn0.Caption = "O"
            is_x = False
        End If
        Call Checking
        can_click_btn(0) = False
    End If
End Sub

Private Sub btn1_Click()
    If can_click_btn(1) = True Then
        If is_x = False Then
            Me.btn1.Caption = "X"
            is_x = True
        Else
            Me.btn1.Caption = "O"
            is_x = False
        End If
        Call Checking
        can_click_btn(1) = False
    End If
End Sub

Private Sub btn2_Click()
    If can_click_btn(2) = True Then
        If is_x = False Then
            Me.btn2.Caption = "X"
            is_x = True
        Else
            Me.btn2.Caption = "O"
            is_x = False
        End If
        Call Checking
        can_click_btn(2) = False
    End If
End Sub

Private Sub btn3_Click()
    If can_click_btn(3) = True Then
        If is_x = False Then
            Me.btn3.Caption = "X"
            is_x = True
        Else
            Me.btn3.Caption = "O"
            is_x = False
        End If
        Call Checking
        can_click_btn(3) = False
    End If
End Sub

Private Sub btn4_Click()
    If can_click_btn(4) = True Then
        If is_x = False Then
            Me.btn4.Caption = "X"
            is_x = True
        Else
            Me.btn4.Caption = "O"
            is_x = False
        End If
        Call Checking
        can_click_btn(4) = False
    End If
End Sub

Private Sub btn5_Click()
    If can_click_btn(5) = True Then
        If is_x = False Then
            Me.btn5.Caption = "X"
            is_x = True
        Else
            Me.btn5.Caption = "O"
            is_x = False
        End If
        Call Checking
        can_click_btn(5) = False
    End If
End Sub

Private Sub btn6_Click()
    If can_click_btn(6) = True Then
        If is_x = False Then
            Me.btn6.Caption = "X"
            is_x = True
        Else
            Me.btn6.Caption = "O"
            is_x = False
        End If
        Call Checking
        can_click_btn(6) = False
    End If
End Sub

Private Sub btn7_Click()
    If can_click_btn(7) = True Then
        If is_x = False Then
            Me.btn7.Caption = "X"
            is_x = True
        Else
            Me.btn7.Caption = "O"
            is_x = False
        End If
        Call Checking
        can_click_btn(7) = False
    End If
End Sub

Private Sub btn8_Click()
    If can_click_btn(8) = True Then
        If is_x = False Then
            Me.btn8.Caption = "X"
            is_x = True
        Else
            Me.btn8.Caption = "O"
            is_x = False
        End If
        Call Checking
        can_click_btn(8) = False
    End If
End Sub

Private Sub btnclear_Click()
    Call InitializeAllData
End Sub

Private Sub btnmainmenu_Click()
    MenuGame.Show
    Unload Me
End Sub

Private Sub btnrestart_Click()
    If player_o_score = 0 Or player_x_score = 0 Then
        Dim answer As Integer
        answer = MsgBox("Are you sure want to restart the game", vbQuestion + vbYesNo, "Restart")
        If answer = vbYes Then
            player_x_score = 0
            player_o_score = 0
            Me.lblplayerxscore.Caption = "Score : " & Str(player_x_score)
            Me.lblplayeroscore.Caption = "Score : " & Str(player_o_score)
            Call InitializeAllData
        End If
    Else
        player_x_score = 0
        player_o_score = 0
        Me.lblplayerxscore.Caption = "Score : " & Str(player_x_score)
        Me.lblplayeroscore.Caption = "Score : " & Str(player_o_score)
        Call InitializeAllData
    End If
End Sub

Private Sub Form_Load()
    Call InitializeAllData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MenuGame.Show
    Me.Hide
End Sub
