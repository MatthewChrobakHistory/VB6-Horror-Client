Attribute VB_Name = "modEditor"
Option Explicit

Public EditorIndex As Long

Public Sub LoadEditor()
Dim i As Long

frmMovieEditor.lstIndex.Clear
For i = 1 To Max_Movies
     frmMovieEditor.lstIndex.AddItem Trim$(Movie(i).Name)
Next
frmMovieEditor.lstIndex = 0

frmMovieEditor.Visible = True

End Sub

Public Sub EditorInit()

If EditorIndex = 0 Then EditorIndex = 1

With frmMovieEditor
    .txtName.Text = Movie(EditorIndex).Name
    .picPoster.Picture = Nothing
    .picPoster.Picture = LoadPicture(App.Path & "/images/" & EditorIndex & ".jpg")
    .txtDirector.Text = Movie(EditorIndex).Director
    .txtIMDbRating.Text = Movie(EditorIndex).IMDBRating
    If Movie(EditorIndex).Watched = True Then
        .chkWatched.Value = 1
    Else
        .chkWatched.Value = 0
    End If
End With

End Sub
