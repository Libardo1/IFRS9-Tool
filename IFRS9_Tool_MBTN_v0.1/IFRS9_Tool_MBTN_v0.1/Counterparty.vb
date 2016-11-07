Imports DataManager

Public Class Counterparty
    Private ID As Long
    Private Name As String
    Private ID_Country As Long
    Private ID_Sector As Long

    Public Sub New(ID As Integer)
        Me.ID = ID
    End Sub

    Public Function Retrieve(ID As Integer, DB As AccessIO)
        DB.OpenConnection()

        'Select all counterparty info
        Dim Query As String = "Select *
                               FROM [Pfo_Counterparty]
                               WHERE ID = " + ID.ToString + ""

        DB.CreateCommand(Query)
        DB.ExecuteReader()
        While (DB.Reader.Read())
            Name = DB.Reader("Description")
            ID_Country = DB.Reader("ID_Country")
            ID_Sector = DB.Reader("ID_Sector")
        End While

        DB.CloseReader()
        DB.CloseConnection()

        Return Me
    End Function

    'Return Rating S&P
    Public ReadOnly Property GetRatingSnP(RefDate As Date, DB As AccessIO) As String
        Get
            Dim CounterpartyRating As New CounterpartyRating
            GetRatingSnP = CounterpartyRating.RetrieveRatingSnP(ID, RefDate, DB)
        End Get
    End Property

    'Return Rating Zanders
    Public ReadOnly Property GetRatingZanders(RefDate As Date, DB As AccessIO) As String
        Get
            Dim CounterpartyRating As New CounterpartyRating
            GetRatingZanders = CounterpartyRating.RetrieveRatingZanders(ID, RefDate, DB)
        End Get
    End Property

    Public ReadOnly Property GetID() As Long
        Get
            Return ID
        End Get
    End Property

    Public ReadOnly Property GetName() As String
        Get
            Return Name
        End Get
    End Property

    Public ReadOnly Property GetID_Country() As Long
        Get
            Return ID_Country
        End Get
    End Property

    Public ReadOnly Property GetID_Sector() As Long
        Get
            Return ID_Sector
        End Get
    End Property

End Class
