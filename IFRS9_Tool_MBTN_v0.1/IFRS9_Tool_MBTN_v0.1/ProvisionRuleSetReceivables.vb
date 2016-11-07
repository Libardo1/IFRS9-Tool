Imports DataManager

Public Class ProvisionRuleSetReceivables
    Private MethodID As Long
    Private MethodName As String
    Private Instrument_Type As String
    Private Rating As String
    Private Receivables As ReceivablePortFolio

    Private Rules As List(Of ProvisionRuleReceivables)

    Public Sub New(MethodID As Integer, Receivables As ReceivablePortFolio)
        Me.MethodID = MethodID + 1 'index starts at 0 but 1 is used in the DB
        Me.Receivables = Receivables
    End Sub

    'Retrieve parametersets description: useful ?
    Public Sub RetrieveProvSetDes(DB As AccessIO)
        DB.OpenConnection()

        Dim Query As String = "Select * 
                               FROM [Par_Receivale_ParameterSets]
                               WHERE ID =" + MethodID.ToString + ""

        DB.CreateCommand(Query)
        DB.ExecuteReader()

        While (DB.Reader.Read())
            MethodID = DB.Reader("ID")
            Instrument_Type = DB.Reader("Instrument_Type")
            MethodName = DB.Reader("Method")
            Rating = DB.Reader("Rating_Type")
        End While

        DB.CloseReader()
        DB.CloseConnection()
    End Sub

    'Retrieve provision rules based on method
    Public Sub RetrieveProvRules(DB As AccessIO)
        Rules = New List(Of ProvisionRuleReceivables)
        DB.OpenConnection()

        Dim Query As String = "Select * 
                               FROM [Par_Parameters_Receivables]
                               WHERE ID_ParameterSet =" + MethodID.ToString + ""

        DB.CreateCommand(Query)
        DB.ExecuteReader()

        Dim ProvisionRule As ProvisionRuleReceivables
        While (DB.Reader.Read())
            ProvisionRule = New ProvisionRuleReceivables
            ProvisionRule.RetrieveRule(DB)
            Rules.Add(ProvisionRule)
        End While

        DB.CloseReader()
        DB.CloseConnection()
    End Sub

    Public ReadOnly Property GetAllRules() As List(Of ProvisionRuleReceivables)
        Get
            Return Rules
        End Get
    End Property

    Public ReadOnly Property GetLossRate(DaysPastDue As Integer, Rating As String, DB As AccessIO) As Double
        Get
            Dim Rule As New ProvisionRuleReceivables
            GetLossRate = Rule.GetLossRate(DaysPastDue, Rating, DB)
        End Get
    End Property

End Class
