VERSION 5.00
Begin VB.Form frmMovieEditor 
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   13035
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkWatched 
      Caption         =   "Watched?"
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txtIMDbRating 
      Height          =   285
      Left            =   6840
      TabIndex        =   10
      Text            =   "0"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtDirector 
      Height          =   285
      Left            =   6480
      TabIndex        =   8
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   6720
      TabIndex        =   6
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   7200
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   6360
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.PictureBox picPoster 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   4755
      Left            =   2400
      ScaleHeight     =   4755
      ScaleWidth      =   3210
      TabIndex        =   1
      Top             =   120
      Width           =   3210
   End
   Begin VB.ListBox lstIndex 
      Height          =   7470
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "IMDb Rating:"
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Director:"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Year Made:"
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmMovieEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkWatched_Click()

If chkWatched.Value = 1 Then
    Movie(EditorIndex).Watched = True
Else
    Movie(EditorIndex).Watched = False
End If

End Sub

Private Sub cmdSave_Click()
Dim i As Long

For i = 1 To Max_Movies
    Call SaveMovie(i)
Next

frmMovieEditor.Visible = False

End Sub

Private Sub lstIndex_Click()

EditorIndex = lstIndex.ListIndex + 1
If EditorIndex = 0 Then lstIndex.ListIndex = 1 And EditorIndex = 1
Call EditorInit

End Sub

Private Sub txtDirector_Change()

Movie(EditorIndex).Director = txtDirector.Text

End Sub

Private Sub txtName_Change()

Movie(EditorIndex).Name = Trim$(txtName.Text)

End Sub

Private Sub txtIMDbRating_Change()

If IsNumeric(txtIMDbRating.Text) = False Then txtIMDbRating.Text = "0"
If txtIMDbRating.Text > 10 Or txtIMDbRating.Text < 0 Then txtIMDbRating.Text = "0"
Movie(EditorIndex).IMDBRating = txtIMDbRating.Text

End Sub

Private Sub txtYear_Change()

If IsNumeric(txtYear.Text) = False Then txtYear.Text = "1990"
Movie(EditorIndex).YearMade = txtYear.Text

End Sub
