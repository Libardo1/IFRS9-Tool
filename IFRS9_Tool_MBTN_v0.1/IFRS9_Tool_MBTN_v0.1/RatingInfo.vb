Imports IFRS9_Tool_MBTN_v0._1.RatingMasterData
Imports IFRS9_Tool_MBTN_v0._1.Method
Imports DataManager

Public Class RatingInfo

    Private DB As AccessIO
    Private RatingData As RatingMasterData

    Private Ratings As List(Of RatingDefinition)
    Private MasterRatingMethod As MasterRatingMethod

    Private MasterScore As Double
    Private MasterRating As String

    Public Property HasFitchRating

    Public Sub New(DB As AccessIO, RatingData As RatingMasterData, RatingMethod As MasterRatingMethod)
        Me.DB = DB
        Me.RatingData = RatingData
        MasterRatingMethod = RatingMethod
    End Sub

    Public Sub Retrieve(CP As String, RefDate As Date)
        'Retrieve from Dat_Rating
        'fill the listOfRatings with data in the db
        Ratings = New List(Of RatingDefinition)

        DB.OpenConnection()

        'select the ratings that matches the counterparty name and RefDate
        Dim Query As String = "Select *
                               FROM [Dat_Rating]
                               WHERE CUSTOMER_NAME = '" & CP.ToString & "' AND VALUE_DATE = #" & RefDate & "#"

        DB.CreateCommand(Query)
        DB.ExecuteReader()

        Dim RatingDef As RatingDefinition
        Dim strRatingName As String
        While (DB.Reader.Read())

            Dim items As Array
            items = System.Enum.GetValues(GetType(RatingSystem))
            Dim item As String
            For Each item In items
                RatingDef = New RatingDefinition
                If item = RatingSystem.Fitch Then
                    strRatingName = "RatingFitch"
                    AddRating(RatingDef, item, strRatingName, RefDate)

                ElseIf item = RatingSystem.Moodys Then
                    RatingDef = New RatingDefinition
                    strRatingName = "RatingMoodys"
                    AddRating(RatingDef, item, strRatingName, RefDate)

                ElseIf item = RatingSystem.SP Then
                    RatingDef = New RatingDefinition
                    strRatingName = "RatingS&P"
                    AddRating(RatingDef, item, strRatingName, RefDate)
                End If
            Next

        End While

        DB.CloseReader()
        DB.CloseConnection()

    End Sub

    Public Sub AddRating(RatingDef As RatingDefinition, RatingSystem As RatingSystem, strRatingName As String, RefDate As Date)
        'if empty, returns the rating of the parent company
        'if no parent company, no rating will be added
        If DB.Reader(strRatingName).ToString <> "" Then
            RatingDef.RatingSystem = RatingSystem
            RatingDef.RatingName = DB.Reader(strRatingName).ToString
            Ratings.Add(RatingDef)
        Else
            If DB.Reader("Parent").ToString <> "" Then
                RatingDef.RatingSystem = RatingSystem
                RatingDef.RatingName = DB.Reader("Parent").ToString
                Ratings.Add(RatingDef)
            End If
        End If

    End Sub

    Public ReadOnly Property GetRatingBySystem(Method As RatingSystem) As RatingDefinition
        Get
            For Each Item In Ratings
                If Item.RatingSystem = Method Then
                    Return Item
                End If
            Next

            'downgrade (not implemented)
            'For i As Integer = 0 To Ratings.Count - 1
            '    If Ratings(i).RatingSystem = Method Then

            '        'check if need to downgrade: USELESS FOR THE MOMENT
            '        If Downgrade = DowngradeNotches.None Then
            '            Return Ratings(i)

            '        ElseIf Downgrade = DowngradeNotches.ThreeNotchesDown Then
            '            Return Ratings(i)

            '        End If
            '    End If
            'Next

            'If this code executes, no rating has been found, return undefined
            Return New RatingDefinition With {.RatingSystem = RatingSystem.Undefined, .RatingName = "Unknown"}
        End Get
    End Property

    Public Function GetOrCalculateMasterRating()
        If MasterRating = "" Then
            CalculateMasterRating()
        End If

        Return MasterRating
    End Function

    Private Function CalculateMasterRating() As RatingDefinition
        'get master score
        MasterScore = CalculateMasterScore()

        'get master rating
        MasterRating = RatingData.GetMasterRating(MasterScore)

        'create new rating defintion for master rating
        Dim RatingDefinition As New RatingDefinition With {.RatingSystem = RatingSystem.Master, .RatingName = MasterRating}
        Ratings.Add(RatingDefinition)

        Return RatingDefinition
    End Function

    Public Function CalculateMasterScore() As Double
        Dim Score, LocalScore As Double, Count As Long

        Select Case MasterRatingMethod
            Case MasterRatingMethod.Average
                Score = 0
                Count = 0

                'Loop trough all Enum values
                Dim items As Array
                items = System.Enum.GetValues(GetType(RatingSystem))
                Dim item As String
                For Each item In items
                    If item > 0 Then
                        If Not GetRatingBySystem(item).RatingSystem = RatingSystem.Undefined Then
                            'go though all RatingDefinition and get the corresponding hard coded score 
                            Score += RatingData.GetScore(GetRatingBySystem(item))
                            Count += 1
                        End If
                    End If
                Next

                Return Score / Count

            Case MasterRatingMethod.Min 'Lowest rating = highest score
                Score = Double.MinValue

                'Loop trough all Enum values Dim items As Array
                Dim items As Array
                items = System.Enum.GetValues(GetType(RatingSystem))
                Dim item As String
                For Each item In items
                    If item > 0 Then 'Excludes Master And Undefined
                        'downgrade set to 0. Change late by passing the correct downgrade in the function
                        If Not GetRatingBySystem(item).RatingSystem = RatingSystem.Undefined Then
                            LocalScore = RatingData.GetScore(GetRatingBySystem(item))
                            If LocalScore > Score Then Score = LocalScore
                        End If
                    End If
                Next

                Return Score

            Case Else
                Throw New Exception("ERROR in MasterScore calculation")

        End Select

    End Function

End Class
