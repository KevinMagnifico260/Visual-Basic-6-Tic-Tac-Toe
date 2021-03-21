VERSION 5.00
Begin VB.Form MenuGame 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnabout 
      BackColor       =   &H00C0E0FF&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   3435
   End
   Begin VB.CommandButton btnexit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   3435
   End
   Begin VB.CommandButton btnplay 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   3435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tic tac Toe"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   2640
      TabIndex        =   1
      Top             =   1020
      Width           =   5190
   End
End
Attribute VB_Name = "MenuGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnabout_Click()
    MsgBox "Develop : Kevin C. Magnifico" & vbCrLf & "Written in : Visual Basic 6.0" & vbCrLf & "License : Public Domain", vbInformation + vbOKOnly, "About"
End Sub

Private Sub btnexit_Click()
    Unload Me
    Unload MainGame
End Sub

Private Sub btnplay_Click()
    MainGame.Show
    Unload Me
End Sub
