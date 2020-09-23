VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rush Hour"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HSLevel 
      Height          =   255
      Left            =   360
      Max             =   40
      Min             =   1
      TabIndex        =   64
      Top             =   4320
      Value           =   1
      Width           =   3855
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Level : 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   65
      Top             =   4560
      Width           =   3855
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   63
      Left            =   3720
      TabIndex        =   63
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   62
      Left            =   3240
      TabIndex        =   62
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   61
      Left            =   2760
      TabIndex        =   61
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   60
      Left            =   2280
      TabIndex        =   60
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   59
      Left            =   1800
      TabIndex        =   59
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   58
      Left            =   1320
      TabIndex        =   58
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   57
      Left            =   840
      TabIndex        =   57
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   56
      Left            =   360
      TabIndex        =   56
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   55
      Left            =   3720
      TabIndex        =   55
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   54
      Left            =   3240
      TabIndex        =   54
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   53
      Left            =   2760
      TabIndex        =   53
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   52
      Left            =   2280
      TabIndex        =   52
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   51
      Left            =   1800
      TabIndex        =   51
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   50
      Left            =   1320
      TabIndex        =   50
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   49
      Left            =   840
      TabIndex        =   49
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   48
      Left            =   360
      TabIndex        =   48
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   47
      Left            =   3720
      TabIndex        =   47
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   46
      Left            =   3240
      TabIndex        =   46
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   45
      Left            =   2760
      TabIndex        =   45
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   44
      Left            =   2280
      TabIndex        =   44
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   43
      Left            =   1800
      TabIndex        =   43
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   42
      Left            =   1320
      TabIndex        =   42
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   41
      Left            =   840
      TabIndex        =   41
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   40
      Left            =   360
      TabIndex        =   40
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   39
      Left            =   3720
      TabIndex        =   39
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   38
      Left            =   3240
      TabIndex        =   38
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   37
      Left            =   2760
      TabIndex        =   37
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   36
      Left            =   2280
      TabIndex        =   36
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   35
      Left            =   1800
      TabIndex        =   35
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   34
      Left            =   1320
      TabIndex        =   34
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   33
      Left            =   840
      TabIndex        =   33
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   32
      Left            =   360
      TabIndex        =   32
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   31
      Left            =   3720
      TabIndex        =   31
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   30
      Left            =   3240
      TabIndex        =   30
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   29
      Left            =   2760
      TabIndex        =   29
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   28
      Left            =   2280
      TabIndex        =   28
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   27
      Left            =   1800
      TabIndex        =   27
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   26
      Left            =   1320
      TabIndex        =   26
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   25
      Left            =   840
      TabIndex        =   25
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   24
      Left            =   360
      TabIndex        =   24
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   23
      Left            =   3720
      TabIndex        =   23
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   22
      Left            =   3240
      TabIndex        =   22
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   21
      Left            =   2760
      TabIndex        =   21
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   20
      Left            =   2280
      TabIndex        =   20
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   19
      Left            =   1800
      TabIndex        =   19
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   18
      Left            =   1320
      TabIndex        =   18
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   17
      Left            =   840
      TabIndex        =   17
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   16
      Left            =   360
      TabIndex        =   16
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   15
      Left            =   3720
      TabIndex        =   15
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   14
      Left            =   3240
      TabIndex        =   14
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   13
      Left            =   2760
      TabIndex        =   13
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   12
      Left            =   2280
      TabIndex        =   12
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   11
      Left            =   1800
      TabIndex        =   11
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   10
      Left            =   1320
      TabIndex        =   10
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   9
      Left            =   840
      TabIndex        =   9
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   8
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   7
      Left            =   3720
      TabIndex        =   7
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   6
      Left            =   3240
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Cell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Levels(40) As String

Private Sub Form_Activate()
MsgBox "The purpose of the game is to move out the RED tile from the black box!." + vbCrLf + _
       "You may move each tile only in its longer direction." + vbCrLf + _
       "To move in horizontal direction click on Right or Left part of tile" + vbCrLf + _
       "To move in vertical direction click on Top or Bottom part of tile" + vbCrLf + vbCrLf + _
       "... Enjoy ... Mahdi Shakouri Rad, Mahdi_Rad@yahoo.com" + vbCrLf + vbCrLf + _
       "NOTE : There is also an user-friendly version of this game on" + vbCrLf + _
       "www.PlanetSourceCode.com, named 'Rush Hour' " _
       , vbOKOnly, "How to play ..."

End Sub

Private Sub Form_Load()
Dim i As Integer

Levels(1) = "A..BBBA..C.D##.CEDFFF.ED..G.HHIIGJJ."
Levels(2) = "AA...BC..D.BC##D.BC..D..E...FFE.GGG."
Levels(3) = "A..B..A..B..A##B....CDDD..C..E..FFFE"
Levels(4) = ".............##C...BBCD..E.CD..EFFD."
Levels(5) = "AA.B..CC.BDE.##FDEGGHFDEI.HF..I..JJJ"
Levels(6) = "AA.B.CD..BECD##BEFDGGGEFH...IIH...JJ"
Levels(7) = ".ABBCD.A.ECD.##E.F..GG.F...H.....H.."
Levels(8) = "...AAB..CCDB##EFDBGGEFHHIIJKKKLLJMMM"
Levels(9) = ".ABBCC.A.DEE##.DFGHIIIFGH.J.FKH.J..K"
Levels(10) = "AAB.CCDDB..FG##..FGHHH.FG..IJJKK.ILL"
Levels(11) = "ABBC..A..C..A##C....DEEE..D..F..GGGF"
Levels(12) = "ABB..CA.D..C##D..C..DEEE....F.GGG.F."
Levels(13) = "AABBC...D.CE.FD##EGF.HHEG..IJJGKKILL"
Levels(14) = "AAB.....B.CCDE##FGDEHHFG..I.JJKKI..."
Levels(15) = ".AABB.CCDDEFGH##EFGHIJEFGHIJKK.LLMM."
Levels(16) = "AABBCDE.FFCDEGH##D.GHIII..H...JJ...."
Levels(17) = "ABBB..A.CCDD##E...FFEG..HHHGIJKKKGIJ"
Levels(18) = "AABC..DDBC..E##C..EFFF..EGG...HHH..."
Levels(19) = "..ABB...A.C..D##C..DEEF..GGGF......."
Levels(20) = "A..BBBACCD..##ED.F..E..F..GHHF..GIII"
Levels(21) = "AABC..D.BC..D##C..DEEE...........FFF"
Levels(22) = "..ABBBC.ADEEC##D...F.DGGHFII.JHKKK.J"
Levels(23) = "..AAAB..CDDB..C##B..EFGG..EFHH..III."
Levels(24) = "..ABB..CA...DC##E.DFF.E.GGG.H.II..H."
Levels(25) = "AAB.CCDDB..EF##.GEFHHHGEFI.JKK.I.JLL"
Levels(26) = ".A.BBBCA.DE.C##DEFGHHHEFG.I..J..IKKJ"
Levels(27) = "ABBC..ADDC..##EC.F..EGGF..H..F..HIII"
Levels(28) = "AAAB....CBDD##C...EFCGGHEFIIIHJJKK.H"
Levels(29) = "AAA.B...C.B.##C.BDEFFGGDEHHI.JKKKI.J"
Levels(30) = "A.BCCCA.BD..A##D..EEFF.G.....GHHII.G"
Levels(31) = "AA.BBB...CDDE##C.GE.HIIGJJH..G..HFFF"
Levels(32) = "AABCDD..BC..##B...FGGHHIF..J.IKK.J.I"
Levels(33) = ".AB.CC.AB...##B...EFFGGHEIIJKHDDDJKH"
Levels(34) = "A..BBBA..C.D##.CEDFFFGED..HGIIJJHKK."
Levels(35) = "..ABBC..AD.C##AD.CEFFF..EGGHI.JJ.HI."
Levels(36) = "ABBBCCADEE.FAD##.FHHHI.F..JIKKLLJ..."
Levels(37) = "AAB.CCDDB.EFG##.EFGIIIEFG..JKKLL.JHH"
Levels(38) = "A..BBBACCD..##FD.G..FEEG..HIIG..HJJJ"
Levels(39) = "ABB.C.ADE.CLADE##LGGGH.L..IHJJKKIFF."
Levels(40) = "..ABBB..AC..##DC.EFFDGGEHIJJ.EHIKK.."

' Borders are Black
For i = 0 To 63
Cell(i).BorderStyle = 0
c = i Mod 8 ' Column of the Cell
r = Int(i / 8)
  If c = 0 Or c = 7 Or r = 0 Or r = 7 Then
    Cell(i).BackColor = vbBlack
    Else
    Cell(i).BackColor = vbWhite
  End If
Next i

Call Level(1)

End Sub


Private Sub Cell_Click(Index As Integer)
Dim r As Integer, c As Integer, i As Integer, j As Integer ' Row, Column
Dim Delta As Integer
' Don't do any thing if the cell is white ( Empty ) or black ( Border )
If Cell(Index).BackColor = vbWhite Or Cell(Index).BackColor = vbBlack Then Exit Sub

' First find the coordinate of this Cell
c = Index Mod 8      ' Column of the Cell = 0 .. 7
r = Int(Index / 8)   ' Row of the Cell = 0 .. 7
'Cell(Index).Caption = Str(c) + "," + Str(r)
If Cell(Index).BackColor = Cell(Index + 1).BackColor Then
  Delta = -1
  Else
  If Cell(Index).BackColor = Cell(Index - 1).BackColor Then
    Delta = 1
    Else
    If Cell(Index).BackColor = Cell(Index + 8).BackColor Then
      Delta = -8
    Else
    If Cell(Index).BackColor = Cell(Index - 8).BackColor Then
      Delta = 8
    End If
    End If
  End If
End If

If Cell(Index + Delta).BackColor = vbWhite Then ' Movement is possible
  If Cell(Index - 2 * Delta).BackColor = Cell(Index).BackColor Then ' 3 pieces
    Cell(Index - 2 * Delta).BackColor = vbWhite
    Cell(Index + Delta).BackColor = Cell(Index).BackColor
    Else
    If Cell(Index - Delta).BackColor = Cell(Index).BackColor Then ' 2 pieces
      Cell(Index - Delta).BackColor = vbWhite
      Cell(Index + Delta).BackColor = Cell(Index).BackColor
    End If
  End If
End If
  
' Check if the level is finished
If Cell(31).BackColor = vbRed Then
  MsgBox "Congratulations ... You finished the level", , "Congratulations ..."
  ' Load the next level
End If
End Sub

Private Sub Level(l As Integer)
Dim i As Integer, r As Integer, c As Integer, k As Integer
' Define Levels ....

k = 1
For i = 0 To 63
  c = i Mod 8      ' Column of the Cell = 0 .. 7
  r = Int(i / 8)   ' Row of the Cell = 0 .. 7
  If r = 0 Or r = 7 Or c = 0 Or c = 7 Then
    ' Its a border cell
    Else
    Select Case Mid(Levels(l), k, 1)
      Case "A": Cell(i).BackColor = &HFF00FF
      Case "B": Cell(i).BackColor = vbBlue
      Case "C": Cell(i).BackColor = vbCyan
      Case "D": Cell(i).BackColor = vbGreen
      Case "E": Cell(i).BackColor = vbYellow
      Case "F": Cell(i).BackColor = &H80FF&
      Case "G": Cell(i).BackColor = &HC0C0C0
      Case "H": Cell(i).BackColor = &H800080
      Case "I": Cell(i).BackColor = &H800000
      Case "J": Cell(i).BackColor = &HC0C0FF
      Case "K": Cell(i).BackColor = &H40C0&
      Case "L": Cell(i).BackColor = &H8000&
      Case "M": Cell(i).BackColor = &H8080&
      Case "#": Cell(i).BackColor = vbRed     ' the target
      Case Else: Cell(i).BackColor = vbWhite  ' empty
    End Select
    k = k + 1
  End If
Next i
Cell(31).BackColor = vbWhite  ' Exit way
End Sub

Private Sub HSLevel_Change()
  lblLevel.Caption = "Level : " + Str(HSLevel.Value)
  Call Level(HSLevel.Value)
End Sub
