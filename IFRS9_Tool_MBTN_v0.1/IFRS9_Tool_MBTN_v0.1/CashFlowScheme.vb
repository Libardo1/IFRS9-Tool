Imports DataManager

Public Class CashFlowScheme
    Private ID As Long
    Private CashFlows As List(Of CashFlow)
    Private Type As CFSchemeType

    Public Enum CFSchemeType
        LiquidityTypical = 1
        InterestRateTypical = 2
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(CFSchemeType As CFSchemeType)
        CashFlows = New List(Of CashFlow)
        Type = CFSchemeType
    End Sub


    Public Sub RetrieveCFs(ID_Instrument As Long, DB As AccessIO)
        CashFlows = New List(Of CashFlow)


        DB.OpenConnection()

        Dim QueryCF As String = "Select * 
                               FROM [CashFlow_Bloomberg]
                               WHERE ID_Instrument = " + ID_Instrument.ToString + ""

        DB.CreateCommand(QueryCF)
        DB.ExecuteReader()

        Dim CashFLow As CashFlow
        While (DB.Reader.Read())
            CashFLow = New CashFlow
            CashFLow.Retrieve(DB)
            CashFlows.Add(CashFLow)
        End While

        DB.CloseReader()
        DB.CloseConnection()
    End Sub

    Public ReadOnly Property GetCFSchemeID() As Long
        Get
            Return ID
        End Get
    End Property

    Public ReadOnly Property GetCashFlowByDate(RefDate As Date) As CashFlow

        Get
            GetCashFlowByDate = New CashFlow
            For Each CashFlow In CashFlows
                If CashFlow.GetCFDate = RefDate Then Return CashFlow
            Next
        End Get
    End Property

    Public ReadOnly Property GetCashFlows() As List(Of CashFlow)
        Get
            Return CashFlows
        End Get
    End Property

    'MAXIME: ugly
    Public ReadOnly Property GetCFSchemeType() As CFSchemeType
        Get
            Return Type = CashFlows(0).GetID_CFScheme
        End Get
    End Property

    Public Sub SetCashFlowsAmounts(value As Double)
        For Each CashFlow In CashFlows

            CashFlow.SetCashFlowAmount() = value
        Next
    End Sub

    Public Sub AddCashFlows(CF As CashFlow)
        CashFlows.Add(CF)
    End Sub

End Class
