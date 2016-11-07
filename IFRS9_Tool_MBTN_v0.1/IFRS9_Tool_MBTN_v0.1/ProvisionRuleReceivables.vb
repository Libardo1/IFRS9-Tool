Imports DataManager

Public Class ProvisionRuleReceivables
    Private ID As Long
    Private ID_ParameterSet As Integer
    Private SeqNr As Integer
    Private DaysPastDue As Integer
    Private LossRate As Double
    Private Rating As String

    Public Function RetrieveRule(DB As AccessIO) As ProvisionRuleReceivables
        ID = DB.Reader("ID")
        ID_ParameterSet = DB.Reader("ID_ParameterSet")
        SeqNr = DB.Reader("SeqNr")
        DaysPastDue = DB.Reader("DaysPastDue")
        LossRate = DB.Reader("LossRate")
        Rating = DB.Reader("Rating")

        Return Me
    End Function

    Public Function GetLossRate(DaysPastDue As Integer, Rating As String, DB As AccessIO) As Double
        DB.OpenConnection()

        'Retrieve LossRate w.r.t DaysPastDue and Rating
        Dim Query As String = "SELECT LossRate
                               FROM [Par_Parameters_Receivables]
                               WHERE DaysPastDue <= " + DaysPastDue.ToString + " AND Rating = '" + Rating.ToString + "'
                               ORDER BY SeqNr DESC"

        DB.CreateCommand(Query)
        DB.ExecuteReader()

        While (DB.Reader.Read())
            LossRate = DB.Reader("LossRate")
            Exit While
        End While

        'Set LossRate to 0 if no DaysPastDue
        If DaysPastDue <= 0 Then
            LossRate = 0
        End If

        DB.CloseReader()
        DB.CloseConnection()

        Return LossRate
    End Function

End Class
