Imports System
Imports DataManager


Public Class RatingMasterData
    Private DB As AccessIO

    'Hardcoded rating scale, to be completed
    'Private AnadolubankRatings As String() = {"AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-", "BB+", "BB", "BB-", "B+", "B", "B-", "CCC+", "CCC", "CCC-", "CC", "C", "DDD", "DD", "D", "UNRATED"}
    'Private ScoreAnadolubank As Double() = {0.0003, 0.0003, 0.0003, 0.0003, 0.0004, 0.0005, 0.001, 0.0019, 0.0029, 0.0044, 0.0066, 0.0101, 0.0161, 0.0275, 0.0521, 0.1125, 0.2847, 0.2847, 0.2847, 0.2847, 0.2847, 1, 1, 1, 0.0415}

    'Ratings & Score for each is the same in the DB for the moment
    Private AnadolubankRatings As String()
    Private ScoreAnadolubank As Double()

    Private FitchRatings As String()
    Private ScoreFitch As Double()

    Private MoodysRatings As String()
    Private ScoreMoodyes As Double()

    Private SPRatings As String()
    Private ScoreSP As Double()

    Private InternalRatings As String() = {"1", "2", "3", "4"}
    Private MasterRatings As String() = InternalRatings
    Private MasterScores As Double() = {0.0, 2.0, 3.5, 4.5} 'Scores of master rating are lower bounds

    Private RatingDictionary As Dictionary(Of RatingDefinition, Double)
    Private MasterDictionary As Dictionary(Of Double, String)

    Public Enum RatingSystem
        Master = -1
        Undefined = 0
        Fitch = 1
        Moodys = 2
        SP = 3
        Internal = 4

        Anadolubank = 5
    End Enum

    Public Structure RatingDefinition
        Public RatingSystem As RatingSystem
        Public RatingName As String
    End Structure

    Public Sub New(DB As AccessIO)
        Me.DB = DB

        'Builds a Dictionary of all known ratings
        RatingDictionary = New Dictionary(Of RatingDefinition, Double)

        'ONLY ANADOLUBANK RATING IS USED
        RetrieveRatingsFromDB(RatingSystem.Anadolubank)

        RetrieveRatingsFromDB(RatingSystem.Fitch)

        RetrieveRatingsFromDB(RatingSystem.Moodys)

        RetrieveRatingsFromDB(RatingSystem.SP)

        'Create MasterDictionary
        MasterDictionary = New Dictionary(Of Double, String)
        AddMasterDictionary(MasterScores, MasterRatings)
    End Sub

    Private Sub AddMasterDictionary(Scores As Double(), Ratings As String())
        For i = 0 To Ratings.Count - 1
            MasterDictionary.Add(Scores(i), Ratings(i))
        Next
    End Sub

    'rating with PD only available for Anadolubank
    Private Sub RetrieveRatingsFromDB(ratingsystem As RatingSystem)
        Dim strRatingName As String
        Dim Key As RatingDefinition

        Select Case RatingSystem.Anadolubank
            Case RatingSystem.Anadolubank
                strRatingName = "Dat_RatingsPDAnadolubank"
            Case RatingSystem.Moodys
                strRatingName = "Dat_RatingsPDMoodys"
            Case RatingSystem.Fitch
                strRatingName = "Dat_RatingsPDFitch"
            Case RatingSystem.SP
                strRatingName = "Dat_RatingsPDS&P"
            Case Else
                strRatingName = "Dat_RatingsPDAnadolubank"
        End Select

        'read data from DB
        DB.OpenConnection()

        Dim Query As String = "Select *
                               FROM " & strRatingName & ""

        DB.CreateCommand(Query)
        DB.ExecuteReader()

        While (DB.Reader.Read())
            Key = New RatingDefinition With {.RatingSystem = ratingsystem, .RatingName = DB.Reader("Rating")}
            RatingDictionary.Add(Key, DB.Reader("PD"))
        End While

        DB.CloseReader()
        DB.CloseConnection()

    End Sub

    Public Function GetMasterRating(MasterScore As Double) As String

        Select Case MasterScore
            Case MasterScore < 0 'case negative
                Throw New Exception("ERROR: Negative MasterScore")

            Case MasterScore >= MasterScores(MasterScores.Length - 1) 'case: higher than the highest score
                Return MasterRatings(MasterScores.Length - 1)

            Case Else 'case between worst and best
                For i = 0 To MasterScores.Length - 2
                    If MasterScore >= MasterScores(i) And MasterScore < MasterScores(i + 1) Then
                        Return MasterRatings(i)
                    End If
                Next
        End Select

    End Function

    Public Function GetScore(Key As RatingDefinition) As Double

        Dim Score As Double
        If (RatingDictionary.TryGetValue(Key, Score)) Then
            RatingDictionary(Key) = Score
        End If

        Return Score
    End Function

End Class
