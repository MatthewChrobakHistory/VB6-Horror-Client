VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   9900
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17385
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   17385
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000E&
      Height          =   9615
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3615
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         ForeColor       =   &H00FFFFFF&
         Height          =   8805
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdEditor 
         Caption         =   "Editor"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox picPoster 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   8
      Left            =   14040
      ScaleHeight     =   4755
      ScaleWidth      =   3210
      TabIndex        =   7
      Top             =   5040
      Width           =   3210
   End
   Begin VB.PictureBox picPoster 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   7
      Left            =   10680
      ScaleHeight     =   4755
      ScaleWidth      =   3210
      TabIndex        =   6
      Top             =   5040
      Width           =   3210
   End
   Begin VB.PictureBox picPoster 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   6
      Left            =   7320
      ScaleHeight     =   4755
      ScaleWidth      =   3210
      TabIndex        =   5
      Top             =   5040
      Width           =   3210
   End
   Begin VB.PictureBox picPoster 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   5
      Left            =   3960
      ScaleHeight     =   4755
      ScaleWidth      =   3210
      TabIndex        =   4
      Top             =   5040
      Width           =   3210
   End
   Begin VB.PictureBox picPoster 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   4
      Left            =   14040
      ScaleHeight     =   4755
      ScaleWidth      =   3210
      TabIndex        =   3
      Top             =   120
      Width           =   3210
   End
   Begin VB.PictureBox picPoster 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   1
      Left            =   3960
      ScaleHeight     =   4755
      ScaleWidth      =   3210
      TabIndex        =   2
      Top             =   120
      Width           =   3210
   End
   Begin VB.PictureBox picPoster 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   2
      Left            =   7320
      ScaleHeight     =   4755
      ScaleWidth      =   3210
      TabIndex        =   1
      Top             =   120
      Width           =   3210
   End
   Begin VB.PictureBox picPoster 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   3
      Left            =   10680
      ScaleHeight     =   4755
      ScaleWidth      =   3210
      TabIndex        =   0
      Top             =   120
      Width           =   3210
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEditor_Click()
    Call LoadEditor
End Sub

Private Sub Command1_Click()
Call PresentRandomMovies
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub picPoster_Click(Index As Integer)

MsgBox Trim$(Movie(Poster(Index).Movie).Name)

End Sub
