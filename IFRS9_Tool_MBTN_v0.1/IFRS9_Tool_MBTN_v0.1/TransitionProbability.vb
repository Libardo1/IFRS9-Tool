Imports DataManager

Public Class TransitionProbability
    Private ID As Long
    Private Start_Date As Date
    Private End_Date As Date
    Private ID_Matrix As Long
    Private RatingFrom As String
    Private RatingTo As String

    Private Prob As Double

    Public Function RetrieveTP(DB As AccessIO) As TransitionProbability
        'TO DO: update corresponding elements to the one in the DB
        ID = DB.Reader("ID")
        Start_Date = DB.Reader("Date_Start")
        End_Date = DB.Reader("Date_End")
        ID_Matrix = DB.Reader("ID_Matrix")
        RatingFrom = DB.Reader("RatingFrom")
        RatingTo = DB.Reader("RatingTo")
        Prob = DB.Reader("Prob")

        Return Me
    End Function

    Public ReadOnly Property GetProb() As Double
        Get
            Return Prob
        End Get
    End Property

End Class
