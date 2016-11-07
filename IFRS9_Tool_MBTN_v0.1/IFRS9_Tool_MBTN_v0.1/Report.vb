Imports DataManager
Imports IFRS9_Tool_MBTN_v0._1.Method

Public Class Report
    Private ID As Long
    Private ReportDate As Date
    Private strDataPath As String

    ' still used for Receivables. Remove afterward
    Private DBConnection As AccessIO

    'Private Guarantees As GuaranteePortfolio
    'Private LoanCommitments As CommitmentPortfolio

    Private Receivables As ReceivablePortFolio
    Private FixedIncomeTransactions As FIPortfolio

    Private ProvisionCalculatorReceivables As ProvisionCalculatorReceivables
    Private ProvisionCalculatorFixedIncomeTransactions As ProvisionCalculatorFI

    'later: construc generic constructor for provisions, FI, ...
    Public Sub New(ID As Long, ReportDate As Date, strDataPath As String)
        Me.ID = ID
        Me.ReportDate = ReportDate
        Me.strDataPath = strDataPath

        FixedIncomeTransactions = New FIPortfolio(ReportDate)
        DBConnection = New AccessIO(strDataPath)
    End Sub

    Public Sub Initialize()
        'dummy
    End Sub

    Public Sub CalcReceivablesProvisions()
        'Retrieve receivables data
        Receivables = New ReceivablePortFolio(ReportDate)
        Receivables.RetrieveData(DBConnection)

        'Retrieve Rules
        ProvisionCalculatorReceivables = New ProvisionCalculatorReceivables(ReportDate, Receivables)
        ProvisionCalculatorReceivables.RetrieveRules(DBConnection)

        'Calc and Save calculated provisions
        ProvisionCalculatorReceivables.CalcProvision(DBConnection)
    End Sub

    Public Sub SetMehods(CalculationMethod As CalculationMethod, Timing As ProvisioningMoments, PDMethod As PDMethod, EADMethod As EADMethod, LGDMethod As LGDMethod,
                         MasterRatingMethod As MasterRatingMethod, StagingMethod As StagingMethod, ThresholdX As Double, DownGrade As DowngradeNotches)
        'set methods
        FixedIncomeTransactions.SetMethods(CalculationMethod, Timing, PDMethod, EADMethod, LGDMethod, MasterRatingMethod, StagingMethod, ThresholdX, DownGrade)
    End Sub

    Public Sub RetrieveDataFI()
        'set DB connection
        FixedIncomeTransactions.DBConnection = DBConnection

        'retrieve data
        FixedIncomeTransactions.RetrieveData()
    End Sub

    Public Sub CalculateCashFlowScheme()
        FixedIncomeTransactions.CalculateCashFlowSchemes()
    End Sub

    Public Sub RetrieveRating()
        FixedIncomeTransactions.RetrieveRating()

        'Dim Transaction As FixedIncomeTransaction
        'For Each Transaction In FixedIncomeTransactions.GetAllFixedIncomes
        '    Debug.Print("ID: " & CStr(Transaction.GetID_Perso) & "; " & "Rating: " & CStr(Transaction.GetRatingBySystem(RatingMasterData.RatingSystem.Anadolubank, Transaction.GetDowngrade)))
        'Next
    End Sub

    Public Sub CalcFIGrossCarryingValue()
        FixedIncomeTransactions.CalculatePortfolioGrossCarryingValue()
    End Sub

    Public Sub CalcFIProvisions()
        FixedIncomeTransactions.CalculatePortfolioProvision()
    End Sub

    Public Sub CalcFINetCarryingValue()
        FixedIncomeTransactions.CalculatePortfolioNetCarryingValue()
    End Sub


    Public Sub SaveResults()
        'Access
        'FixedIncomeTransactions.SavePortfolioResultAccess(DBConnection, ID)
        FixedIncomeTransactions.SaveInstrumentResultAccess(DBConnection, strDataPath, ID)

        'Excel
        FixedIncomeTransactions.SavePortfolioResultExcel(ID)
        FixedIncomeTransactions.SaveInstrumentResultExcel()
    End Sub


    Public ReadOnly Property GetID() As Double
        Get
            Return ID
        End Get
    End Property

    Public ReadOnly Property GetReportDate() As Date
        Get
            Return ReportDate
        End Get
    End Property
End Class
