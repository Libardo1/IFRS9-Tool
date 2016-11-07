Imports DataManager

Public Class CounterpartyRating
    Private ID As Long
    Private ID_Counterparty As Long
    Private Rating_SnP As String
    Private Rating_Zanders As String
    Private Date_Start As Date
    Private Date_End As Date

    Public Function RetrieveRatingSnP(ID As Long, RefDate As Date, DB As AccessIO)
        DB.OpenConnection()

        'Get Rating_S&P
        Dim Query As String = "Select *
                               FROM [Pfo_CounterpartyRating_X]
                               WHERE (ID_Counterparty = " + ID.ToString + ") AND 
                                Date_Start <= #" + RefDate.ToString + "# AND Date_End > #" + RefDate.ToString + "#"

        DB.CreateCommand(Query)
        DB.ExecuteReader()
        While (DB.Reader.Read())
            Rating_SnP = DB.Reader("Rating_S&P")
        End While

        DB.CloseReader()
        DB.CloseConnection()

        Return Rating_SnP
    End Function

    Public Function RetrieveRatingZanders(ID As Long, RefDate As Date, DB As AccessIO)
        DB.OpenConnection()

        'Get Rating_S&P
        Dim Query As String = "Select *
                               FROM [Pfo_CounterpartyRating_X]
                               WHERE (ID_Counterparty = " + ID.ToString + ") AND (Date_Start = #" + RefDate + "#)"

        DB.CreateCommand(Query)
        DB.ExecuteReader()
        While (DB.Reader.Read())
            Rating_Zanders = DB.Reader("Rating_Zanders")
        End While

        DB.CloseReader()
        DB.CloseConnection()

        Return Rating_Zanders
    End Function

End Class
