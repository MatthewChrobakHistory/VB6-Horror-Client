Attribute VB_Name = "modData"
Option Explicit

Public Const Max_Movies As Long = 273

Public Movie(1 To Max_Movies) As MovieRec
Public MovieScreen As ScreenRec
Public Poster(1 To 8) As PosterRec

Private Type MovieRec
    Name As String
    YearMade As String
    Director As String
    IMDBRating As Byte
    Comments As String
    Picture As String
    Rating As String
    RatingReasons As String
    Genre As String
    Plot As String
    Sequal As String
    Prequal As String
    Watched As Boolean
    RemakeName As String
    RemakeYear As String
End Type

Private Type ScreenRec
    SCSequal As String
    SCPrequal As String
    SCRemakeName As String
    SCRemakeYear As String
End Type

Private Type PosterRec
    Movie As Long
End Type

Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Sub Main()
Dim i As Long

For i = 1 To Max_Movies
    If LenB(Dir(App.Path & "/Data/" & i & ".dat")) = 1 Then
        SaveMovie (i)
    Else
        LoadMovie (i)
    End If
Next

PresentRandomMovies

frmMain.Show

End Sub

Public Function RAND(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    RAND = Int((High - Low + 1) * Rnd) + Low
End Function

Sub ClearMovie(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Movie(Index)), LenB(Movie(Index)))
End Sub

Sub LoadMovie(ByVal Index As Long)
    Dim FileName As String
    Dim F As Long
    Call ClearMovie(Index)
    
    FileName = App.Path & "\Data\" & Index & ".bin"
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , Movie(Index)
    Close #F
End Sub

Sub SaveMovie(ByVal Index As Long)
    Dim FileName As String
    Dim F As Long

    FileName = App.Path & "\Data\" & Index & ".bin"
    
    F = FreeFile
    
    Open FileName For Binary As #F
    Put #F, , Movie(Index)
    Close #F
End Sub

Sub PresentRandomMovies()
Dim i As Long
Dim x As Long
Dim Movie As Long
Dim FoundMovie As Boolean

For i = 1 To 8
    Do While FoundMovie = False
        FoundMovie = True
        Movie = RAND(1, Max_Movies)
        For x = 1 To 8
            If Movie = Poster(i).Movie Then
                FoundMovie = False
            End If
        Next
    Loop
    Call SetPresentMovie(i, Movie)
    FoundMovie = False
Next

For i = 1 To 8
    For x = 1 To 8
        If Poster(x).Movie = Poster(i).Movie And x <> i Then
            'fix this later
            Call PresentRandomMovies
        End If
    Next
Next
        
End Sub

Sub SetPresentMovie(ByVal Screen As Long, ByVal Index As Long)

With frmMain
    .picPoster(Screen).Picture = Nothing
    .picPoster(Screen).Picture = LoadPicture(App.Path & "/images/" & Index & ".jpg")
    Poster(Screen).Movie = Index
End With
    
End Sub
