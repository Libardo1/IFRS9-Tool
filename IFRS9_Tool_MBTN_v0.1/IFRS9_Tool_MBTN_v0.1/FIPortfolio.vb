Imports DataManager
Imports IFRS9_Tool_MBTN_v0._1.Method

Public Class FIPortfolio
    Inherits Portfolio

    Private FixedIncomeTransactions As List(Of FixedIncomeTransaction)

    Private PortfolioGrossCarryingValue As Double
    Private PortfolioNetCarryingValue As Decimal

    Private CalculationMethod As CalculationMethod
    Private Timing As ProvisioningMoments
    Private PDMethod As PDMethod
    Private EADMethod As EADMethod
    Private LGDMethod As LGDMethod
    Private StagingMethod As StagingMethod
    Private ThresholdX As Double

    Private DownGrade As DowngradeNotches

    Private PortfolioProvision As ECL

    Public Sub New(ReportDate As Date)
        MyBase.TheReportDate = ReportDate

        PortfolioGrossCarryingValue = Decimal.MinValue
        PortfolioNetCarryingValue = Decimal.MinValue
    End Sub

    'put in base class?
    Public Sub SetMethods(CalculationMethod As CalculationMethod, TSet As ProvisioningMoments, PDM As PDMethod, EADM As EADMethod, LGDM As LGDMethod, MasterRatingMethod As MasterRatingMethod,
                          StagingM As StagingMethod, ThresholdX As Double, DownGrade As DowngradeNotches)
        Me.CalculationMethod = CalculationMethod
        Timing = TSet
        PDMethod = PDM
        EADMethod = EADM
        LGDMethod = LGDM
        MyBase.TheMasterRatingMethod = MasterRatingMethod
        StagingMethod = StagingM
        Me.ThresholdX = ThresholdX
        Me.DownGrade = DownGrade
    End Sub

    Public Sub RetrieveData()
        Dim DB As AccessIO = DBConnection

        DB.OpenConnection()
        FixedIncomeTransactions = New List(Of FixedIncomeTransaction)

        Dim Query As String = "Select * FROM Dat_Securities"
        'Dim Query As String = "Select * FROM Dat_LoansSmallDB"
        'Dim Query As String = "Select * FROM Dat_CashEquivalents"

        DB.CreateCommand(Query)
        DB.ExecuteReader()

        'create FixeIncomeTransaction object
        Dim NewFixedIncomeTransaction As FixedIncomeTransaction
        Dim count = 0
        While (DB.Reader.Read())
            NewFixedIncomeTransaction = New FixedIncomeTransaction(CalculationMethod, Timing, PDMethod, EADMethod, LGDMethod, TheMasterRatingMethod, StagingMethod,
                                                                   ThresholdX)
            NewFixedIncomeTransaction.Retrieve(MyBase.TheReportDate, DB)

            FixedIncomeTransactions.Add(NewFixedIncomeTransaction)
            count += 1
        End While

        'find collateral value
        For Each Transaction In FixedIncomeTransactions
            'Query = "Select I.Collateral FROM In_Loans I , Dat_Loans D WHERE I.[LOAN REF NO] = D.[LOAN_REF_NO] And D.[LOAN_REF_NO] = '" & CStr(Transaction.GetID_Perso) & "'"
            Query = "Select I.Collateral 
                     FROM In_Securities I , Dat_Securities D 
                     WHERE I.[LOAN REF NO] = D.[LOAN_REF_NO] And D.[LOAN_REF_NO] = '" & CStr(Transaction.GetID_Perso) & "'"

            'Query = "Select 0 As Collateral"
            DB.CreateCommand(Query)
            DB.ExecuteReader()
            If DB.Reader.Read Then
                Transaction.CollateralValue = DB.Reader("Collateral")
            Else
                Transaction.CollateralValue = 0
            End If
        Next

        'find origination date
        'set it to either the Transaction date (if not available) or to the date Yorick inputted in the DB (for the ratings) 
        For Each Transaction In FixedIncomeTransactions
            Query = "Select FIRST(VALUE_DATE) AS FirstDate
                     From Dat_Rating
                     Where CUSTOMER_NAME = '" + Transaction.TheName_Counterparty + "'"

            DB.CreateCommand(Query)
            DB.ExecuteReader()
            If DB.Reader.Read Then

                Dim type As VariantType = VarType(DB.Reader("FirstDate"))
                Dim str As String
                str = type.ToString

                If VarType(DB.Reader("FirstDate")).ToString = "Null" Then
                    Transaction.TheOriginationDate = Transaction.TheTransactionDate
                Else
                    Transaction.TheOriginationDate = DB.Reader("FirstDate")
                End If
            End If
        Next

        DB.CloseReader()
        DB.CloseConnection()

    End Sub

    Public Sub CalculateCashFlowSchemes()
        Dim DB As AccessIO = DBConnection
        For Each Instrument In FixedIncomeTransactions
            Instrument.CalculateCashFlowScheme()
        Next
    End Sub

    Public Sub RetrieveRating()
        Dim DB As AccessIO = DBConnection
        Dim MasterData As New RatingMasterData(DB)

        For Each Instrument In FixedIncomeTransactions
            Instrument.RetrieveRatingInfo(DB, MasterData, TheMasterRatingMethod)
        Next
    End Sub


    'GrossCaryingValue for the portfolio
    Public Sub CalculatePortfolioGrossCarryingValue()
        PortfolioGrossCarryingValue = 0

        For Each FITransaction In FixedIncomeTransactions
            FITransaction.CalculateEIR()
            FITransaction.CalculateGrossCarryingValue()
            PortfolioGrossCarryingValue += FITransaction.GetGrossCarryingValue
        Next
    End Sub

    'Provision for the portfolio
    Public Sub CalculatePortfolioProvision()
        PortfolioProvision = New ECL(TheReportDate)
        Dim PortfolioProvisionAmount = 0

        For Each FITransaction In FixedIncomeTransactions
            FITransaction.CalculateProvision(DBConnection)
            PortfolioProvisionAmount += FITransaction.TheProvision
        Next

        PortfolioProvision.TheECLValue = PortfolioProvisionAmount
    End Sub

    'NetCarryingValue for the portfolio
    Public Sub CalculatePortfolioNetCarryingValue()
        PortfolioNetCarryingValue = PortfolioGrossCarryingValue - PortfolioProvision.TheECLValue
    End Sub

    'save results
    Public Sub SavePortfolioResultAccess(DB As AccessIO, ID_Report As Integer)
        Dim ID_Portfolio As Integer = 1 'used in database. Change later

        DB.OpenConnection()

        Dim query As String = "INSERT INTO Res_FIPortfolio(ReportDate, ID_Portfolio, Provisioning_Method, PD_Method, EAD_Method, LGD_Method, GrossCarryingValue, Provision, NetCarryingValue)
                               values (#" & TheReportDate & "#," & ID_Portfolio & "," & Timing & "," & PDMethod & "," & EADMethod & "," & LGDMethod & "," & PortfolioGrossCarryingValue & "," & PortfolioProvision.TheECLValue & "," & PortfolioNetCarryingValue & ")"

        DB.CreateCommand(query)
        DB.ExecuteNonQuery()

        DB.CloseConnection()

    End Sub

    Public Sub SavePortfolioResultExcel(ID_Report As Integer)


    End Sub

    'save individual instrument results
    Public Sub SaveInstrumentResultAccess(DB As AccessIO, filepath As String, ID_Report As Integer)
        For Each Instrument In FixedIncomeTransactions
            Instrument.SaveInstrumentResultAccess(DB, filepath, ID_Report)
        Next
    End Sub

    Public Sub SaveInstrumentResultExcel()
        For Each Instrument In FixedIncomeTransactions
            Instrument.SaveInstrumentResultExcel()
        Next

    End Sub

    'Gets functions
    Public ReadOnly Property GetPortfolioGrossCarryingValue()
        Get
            Return PortfolioGrossCarryingValue
        End Get
    End Property

    Public ReadOnly Property GetPortfolioProvision()
        Get
            Return PortfolioProvision.TheECLValue
        End Get
    End Property

    Public ReadOnly Property GetPortfolioNetCarryingValue()
        Get
            Return PortfolioNetCarryingValue
        End Get
    End Property

    Public ReadOnly Property GetPDMethod() As PDMethod
        Get
            GetPDMethod = PDMethod
        End Get
    End Property

    Public ReadOnly Property GetEADMethod() As EADMethod
        Get
            GetEADMethod = EADMethod
        End Get
    End Property

    Public ReadOnly Property GetLGDMethod() As LGDMethod
        Get
            GetLGDMethod = LGDMethod
        End Get
    End Property

    Public ReadOnly Property GetTiming() As LGDMethod
        Get
            GetTiming = Timing
        End Get
    End Property

    Public ReadOnly Property GetAllFixedIncomes() As List(Of FixedIncomeTransaction)
        Get
            GetAllFixedIncomes = FixedIncomeTransactions
        End Get
    End Property

End Class